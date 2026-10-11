# lib/win32ole/jruby/win32ole.rb
require 'fiddle'
require 'win32ole/jruby/win32'
require 'win32ole/jruby/dispatch'
require 'win32ole/jruby/array'
require 'win32ole/jruby/record'
require 'win32ole/jruby/variant'

class WIN32OLE
  RuntimeError = Class.new(::RuntimeError)
  QueryInterfaceError = Class.new(RuntimeError)

  ::Object.const_set(:WIN32OLERuntimeError, RuntimeError)
  ::Object.deprecate_constant(:WIN32OLERuntimeError)

  ::Object.const_set(:WIN32OLEQueryInterfaceError, QueryInterfaceError)
  ::Object.deprecate_constant(:WIN32OLEQueryInterfaceError)

  include Dispatch

  W = Win32
  private_constant :W

  LOCALE_SYSTEM_DEFAULT = W::LOCALE_SYSTEM_DEFAULT
  LOCALE_USER_DEFAULT = W::LOCALE_USER_DEFAULT

  CP_ACP = W::CP_ACP
  CP_OEMCP = W::CP_OEMCP
  CP_MACCP = W::CP_MACCP
  CP_THREAD_ACP = W::CP_THREAD_ACP
  CP_SYMBOL = W::CP_SYMBOL
  CP_UTF7 = W::CP_UTF7
  CP_UTF8 = W::CP_UTF8

  @lcid = LOCALE_SYSTEM_DEFAULT

  class << self
    # ext/win32ole/win32ole.c's fole_s_get_code_page / fole_s_set_code_page:
    # the codepage strings coming back from COM are converted to (initially
    # derived from Encoding.default_internal || default_external).
    def codepage
      W.codepage
    end

    def codepage=(cp)
      W.codepage = cp
    end

    def locale
      @lcid
    end

    # ext/win32ole/win32ole.c's fole_s_set_locale: the two sentinel LCIDs
    # are always accepted (they resolve dynamically at call time, so
    # EnumSystemLocalesA can't confirm them up front); any other LCID must
    # name an installed locale.
    def locale=(lcid)
      unless lcid == LOCALE_SYSTEM_DEFAULT || lcid == LOCALE_USER_DEFAULT || W.locale_installed?(lcid)
        raise WIN32OLE::RuntimeError, "not installed locale: #{lcid}"
      end

      @lcid = lcid
      nil
    end

    # ole_initialize (fole_s_ole_initialize): OleInitialize once per thread
    # -- COM initialization is per OS thread, hence a thread variable, not
    # a fiber-local Thread#[]. (C also installs an OLE message filter; this
    # port doesn't.)
    def ole_initialize
      return nil if Thread.current.thread_variable_get(:win32ole_ole_initialized)

      hr = W.ole_initialize_fn.call(nil)
      raise ::RuntimeError, "fail: OLE initialize\n#{hresult_detail(hr)}" if W.failed?(hr)

      Thread.current.thread_variable_set(:win32ole_ole_initialized, true)
      nil
    end

    def ole_uninitialize
      return nil unless Thread.current.thread_variable_get(:win32ole_ole_initialized)

      W.ole_uninitialize_fn.call
      Thread.current.thread_variable_set(:win32ole_ole_initialized, nil)
      nil
    end

    # fole_s_show_help: open +target+'s help -- a help file path, or a
    # WIN32OLE::Type / WIN32OLE::Method's own help file at its context.
    def ole_show_help(target, helpcontext = nil)
      if target.is_a?(WIN32OLE::Type) || target.is_a?(WIN32OLE::Method)
        helpfile = target.helpfile
        raise ::RuntimeError, "no helpfile of `#{target.name}'" if helpfile.to_s.empty?

        helpcontext = target.helpcontext
      else
        helpfile = target
      end
      raise TypeError, '1st parameter must be (String|WIN32OLE::Type|WIN32OLE::Method)' unless helpfile.is_a?(::String)

      hwnd = W.show_html_help(helpfile, Integer(helpcontext))
      raise ::RuntimeError, "failed to open help file `#{helpfile}'" if hwnd.zero?

      nil
    end

    # fole_s_reference_count (a debugging aid): the IDispatch's reference
    # count, read with an AddRef/Release pair.
    def ole_reference_count(obj)
      ptr = obj.dispatch_ptr
      return 0 if ptr.nil? || ptr.zero?

      W.add_ref(ptr)
      release_com_pointer(ptr)
    end

    # fole_s_free (a debugging aid): one Release while a reference is
    # left, returning the new count. Unlike WIN32OLE#ole_free, the object
    # keeps its pointer.
    def ole_free(obj)
      ptr = obj.dispatch_ptr
      return 0 if ptr.nil? || ptr.zero? || ole_reference_count(obj).zero?

      release_com_pointer(ptr)
    end

    def wrap_dispatch_pointer(ptr)
      obj = allocate
      obj.instance_variable_set(:@ptr, ptr)
      obj.send(:install_finalizer)
      obj
    end

    def connect(server, host = nil)
      server = string_value(server)
      raise NotImplementedError, 'remote OLE (host) is not supported yet' unless host.nil?

      ole_initialize

      # fole_s_connect: a name that is neither a ProgID nor a CLSID is
      # bound as a moniker (e.g. "winmgmts:").
      clsid = clsid_for(server)
      return bind_moniker(server) if clsid.nil?

      ppv = ("\x00" * W::PTR_SIZE).b
      hr = W.get_active_object.call(clsid, nil, ppv)
      if W.failed?(hr)
        raise WIN32OLE::RuntimeError, "OLE server `#{server}' not running\n#{hresult_detail(hr)}"
      end

      # GetActiveObject's out-parameter is IUnknown**, not IDispatch** --
      # ext/win32ole/win32ole.c:1948-1957 QueryInterfaces for IID_IDispatch
      # and releases the original IUnknown* before wrapping the result.
      iunknown_ptr = ppv.unpack1(W::PACK_PTR)
      idispatch_ptr = W.query_interface(iunknown_ptr, W::IID_IDISPATCH)
      release_com_pointer(iunknown_ptr)
      if idispatch_ptr.nil?
        raise WIN32OLE::RuntimeError, W.unknown_server_error_message(server)
      end

      wrap_dispatch_pointer(idispatch_ptr)
    end

    def release_com_pointer(ptr)
      return if ptr.nil? || ptr.zero?

      W.vtable_function(ptr, 2, [W::VOIDP], W::DWORD).call(ptr)
    end

    # +to_free+ collects a proc per allocation behind the VARIANT (C's
    # VariantClear of it); call them once the VARIANT is no longer needed.
    def ruby_value_to_variant_bytes(value, to_free)
      case value
      when ::Array
        psa = SafeArray.ruby_array_to_safearray(value, W::VT_VARIANT, to_free)
        to_free << -> { SafeArray.safe_array_destroy.call(psa) }
        return W.pack_variant(W::VT_VARIANT | W::VT_ARRAY, W.pack_pointer(psa.to_i))
      when WIN32OLE::Record
        return value.to_variant_bytes
      when WIN32OLE::Variant
        return value.instance_variable_get(:@var)
      end

      type = W.ruby_to_variant_type(value)
      payload =
        case type
        when :i4 then W.pack_i4(value)
        when :i8 then W.pack_i8(value)
        when :r8 then W.pack_r8(value)
        when :bool then W.pack_bool(value)
        when :empty then W.pack_empty
        when :bstr
          bstr = W.sys_alloc_string.call(W.wstr(value))
          to_free << -> { W.sys_free_string.call(bstr) }
          W.pack_pointer(bstr)
        when :dispatch
          W.pack_pointer(value.instance_variable_get(:@ptr))
        when :date
          W.pack_r8(W.time_to_vtdate(value))
        end
      W.pack_variant(W::VT_FOR_TYPE.fetch(type), payload)
    end

    # ole_variant2val. owned: true (a call's own result) hands its BSTR /
    # IDispatch reference over to the Ruby value; owned: false (a
    # Variant's own value, a SAFEARRAY element, anything behind VT_BYREF)
    # leaves the VARIANT intact -- the BSTR is not freed and the IDispatch
    # is AddRef'd before being wrapped.
    def variant_bytes_to_ruby_value(bytes, owned: true)
      vt, = W.unpack_variant(bytes)
      if (vt & W::VT_BYREF) != 0
        owned = false
        bytes = W.deref_variant(bytes)
        vt, = W.unpack_variant(bytes)
      end
      base_vt = vt & W::VT_TYPEMASK

      if (vt & W::VT_ARRAY) != 0
        _vt, payload = W.unpack_variant(bytes)
        psa = W.unpack_pointer(payload)
        return nil if psa.zero?

        return SafeArray.safearray_to_ruby_array(psa, base_vt)
      end

      if base_vt == W::VT_RECORD
        _vt, body = W.unpack_variant(bytes, body_size: WIN32OLE::Record::VT_RECORD_BODY_SIZE)
        buffer_ptr, pri = body.unpack("#{W::PACK_PTR}2")
        return WIN32OLE::Record.from_irecordinfo_and_buffer(pri, buffer_ptr)
      end

      type = W.variant_ruby_type(vt)
      _vt2, payload = W.unpack_variant(bytes)
      case type
      when :empty then nil
      when :i4 then W.unpack_i4(payload)
      when :i8 then W.unpack_i8(payload)
      when :r8 then W.unpack_r8(payload)
      when :bool then W.unpack_bool(payload)
      when :i1 then W.unpack_i1(payload)
      when :ui1 then W.unpack_ui1(payload)
      when :i2 then W.unpack_i2(payload)
      when :ui2 then W.unpack_ui2(payload)
      when :ui4 then W.unpack_ui4(payload)
      when :ui8 then W.unpack_ui8(payload)
      when :int then W.unpack_int(payload)
      when :uint then W.unpack_uint(payload)
      when :r4 then W.unpack_r4(payload)
      when :error then W.unpack_error(payload)
      when :date then W.vtdate_to_time(W.unpack_r8(payload))
      when :bstr
        addr = W.unpack_pointer(payload)
        str = W.bstr_to_s(addr) || '' # a NULL BSTR is "", as in ole_variant2val
        W.sys_free_string.call(addr) if owned && !addr.zero?
        str
      when :dispatch
        ptr = W.unpack_pointer(payload)
        unless ptr.zero?
          W.add_ref(ptr) unless owned
          wrap_dispatch_pointer(ptr)
        end
      when :other then variant_to_bstr_value(bytes)
      end
    end

    # ole_variant2val's default branch (VT_CY, VT_DECIMAL, ...): the
    # VariantChangeTypeEx(VT_BSTR) rendering, or nil if it can't convert.
    def variant_to_bstr_value(bytes)
      dest = ("\x00" * W::VARIANT_SIZE).b
      hr = W.variant_change_type.call(dest, bytes, locale, 0, W::VT_BSTR)
      return nil if W.failed?(hr)

      addr = W.unpack_pointer(W.unpack_variant(dest)[1])
      str = W.bstr_to_s(addr) || ''
      W.sys_free_string.call(addr) unless addr.zero?
      str
    end

    def resolve_clsid(server)
      clsid, hr = clsid_and_hr(server)
      return clsid if clsid

      raise WIN32OLE::RuntimeError, "#{W.unknown_server_error_message(server)}\n#{hresult_detail(hr)}"
    end

    # The CLSID for a ProgID or "{...}" CLSID string, or nil.
    def clsid_for(server)
      clsid_and_hr(server).first
    end

    def clsid_and_hr(server)
      wide = W.wstr(server)
      clsid = ("\x00" * 16).b
      hr = W.clsid_from_progid.call(wide, clsid)
      # ole_create_dcom tries the remote host's registry in between.
      if W.failed?(hr) && block_given? && (remote = yield)
        return [remote, 0]
      end
      hr = W.clsid_from_string.call(wide, clsid) if W.failed?(hr)
      [W.failed?(hr) ? nil : clsid, hr]
    end

    IMONIKER_BIND_TO_OBJECT = 8 # IUnknown(3) + IPersist(1) + IPersistStream(4)

    # ole_bind_obj: CreateBindCtx + MkParseDisplayName + IMoniker::BindToObject
    # for IDispatch.
    def bind_moniker(moniker)
      out = ("\x00" * W::PTR_SIZE).b
      hr = W.create_bind_ctx.call(0, out)
      raise WIN32OLE::RuntimeError, "failed to create bind context\n#{hresult_detail(hr)}" if W.failed?(hr)

      bindctx = out.unpack1(W::PACK_PTR)
      begin
        eaten = ("\x00" * 4).b
        out = ("\x00" * W::PTR_SIZE).b
        hr = W.mk_parse_display_name.call(bindctx, W.wstr(moniker), eaten, out)
        if W.failed?(hr)
          raise WIN32OLE::RuntimeError, "failed to parse display name of moniker `#{moniker}'\n#{hresult_detail(hr)}"
        end

        pmoniker = out.unpack1(W::PACK_PTR)
        begin
          out = ("\x00" * W::PTR_SIZE).b
          hr = W.vtable_function(pmoniker, IMONIKER_BIND_TO_OBJECT, [W::VOIDP] * 5, W::LONG)
                .call(pmoniker, bindctx, nil, W::IID_IDISPATCH, out)
          raise WIN32OLE::RuntimeError, "failed to bind moniker `#{moniker}'\n#{hresult_detail(hr)}" if W.failed?(hr)
        ensure
          release_com_pointer(pmoniker)
        end
      ensure
        release_com_pointer(bindctx)
      end
      wrap_dispatch_pointer(out.unpack1(W::PACK_PTR))
    end

    def hresult_detail(hr)
      "    HRESULT error code:#{W.hr_hex(hr)}\n      #{W.hresult_system_message(hr)}"
    end

    def create_guid
      buf = ("\x00" * 16).b
      hr = W.co_create_guid.call(buf)
      raise WIN32OLE::RuntimeError, "failed to create GUID\n#{hresult_detail(hr)}" if W.failed?(hr)

      d1, d2, d3 = buf.unpack('LSS')
      d4 = buf[8, 8].unpack('C8')
      format('{%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}', d1, d2, d3, *d4)
    end

    # ext/win32ole/win32ole.c's fole_s_const_load / ole_const_load. +ole+ is
    # a WIN32OLE (its containing typelib) or a String naming a typelib the
    # way typelib_file resolves it; +mod+ defaults to WIN32OLE itself. Like
    # C, each visible constant's first letter is ASCII-upcased
    # (ssfWINDOWS -> SsfWINDOWS); names that still aren't valid constant
    # names go into a CONSTANTS hash on +mod+ instead.
    def const_load(ole, mod = nil)
      raise TypeError, '2nd parameter must be Class or Module' unless mod.nil? || mod.is_a?(::Module)

      mod ||= self
      tlib =
        case ole
        when WIN32OLE then ole.ole_type.ole_typelib
        when ::String
          ptr = WIN32OLE::TypeLib.load_typelib(WIN32OLE::TypeLib.file_for(ole) || ole)
          raise WIN32OLE::RuntimeError, 'failed to LoadTypeLibEx' if ptr.nil?

          WIN32OLE::TypeLib.from_itypelib_ptr(ptr)
        else
          raise TypeError, '1st parameter must be WIN32OLE instance'
        end

      constants = {}
      tlib.ole_types.each do |type|
        type.variables.each do |var|
          next unless var.variable_kind == 'CONSTANT' && var.visible?

          name = var.name.dup
          name[0] = name[0].upcase(:ascii)
          if name.match?(/\A[A-Z]\w*\z/)
            mod.const_set(name, var.value) unless mod.const_defined?(name, false)
          else
            constants[name] = var.value
          end
        end
      end
      mod.const_set(:CONSTANTS, constants)
      nil
    end

    # Ruby-level StringValue: a String, or anything with an implicit
    # #to_str, else TypeError -- which is what ext/win32ole raises for a
    # non-string server/host/method name.
    def string_value(val)
      String.try_convert(val) || raise(TypeError, "no implicit conversion of #{val.class} into String")
    end

    private :resolve_clsid, :clsid_for, :clsid_and_hr, :bind_moniker, :hresult_detail, :release_com_pointer,
            :string_value
  end

  def initialize(server, host = nil)
    server = self.class.send(:string_value, server)
    self.class.ole_initialize
    return create_dcom(server, self.class.send(:string_value, host)) unless host.nil?

    clsid = self.class.send(:resolve_clsid, server)
    ppv = ("\x00" * W::PTR_SIZE).b
    hr = W.co_create_instance.call(
      clsid, nil, W::CLSCTX_INPROC_SERVER | W::CLSCTX_LOCAL_SERVER, W::IID_IDISPATCH, ppv
    )
    if W.failed?(hr)
      raise WIN32OLE::RuntimeError, "#{W.unknown_server_error_message(server)}\n#{self.class.send(:hresult_detail, hr)}"
    end

    @ptr = ppv.unpack1(W::PTR_SIZE == 8 ? 'Q' : 'L')
    install_finalizer
  end

  private

  # ole_create_dcom: the CLSID from the ProgID here, then from the host's
  # registry, then as a CLSID string; the object from CoCreateInstanceEx
  # on +host+.
  def create_dcom(server, host)
    clsid, hr = self.class.send(:clsid_and_hr, server) { W.clsid_from_remote(host, server) }
    raise WIN32OLE::RuntimeError, "#{W.unknown_server_error_message(server)}\n#{self.class.send(:hresult_detail, hr)}" unless clsid

    host_name = W.wstr(host)
    # COSERVERINFO { DWORD dwReserved1; LPWSTR pwszName; COAUTHINFO *pAuthInfo; DWORD dwReserved2; }
    info = [0, Fiddle::Pointer.to_ptr(host_name).to_i, 0, 0].pack("L#{'x4' if W::PTR_SIZE == 8}#{W::PACK_PTR}#{W::PACK_PTR}L#{'x4' if W::PTR_SIZE == 8}")
    iid = W::IID_IDISPATCH.dup
    # MULTI_QI { const IID *pIID; IUnknown *pItf; HRESULT hr; }
    mqi = [Fiddle::Pointer.to_ptr(iid).to_i, 0, 0].pack("#{W::PACK_PTR}#{W::PACK_PTR}l#{'x4' if W::PTR_SIZE == 8}")
    hr = W.co_create_instance_ex.call(clsid, nil, W::CLSCTX_REMOTE_SERVER, info, 1, mqi)
    if W.failed?(hr)
      raise WIN32OLE::RuntimeError,
            "failed to create DCOM server `#{server}' in `#{host}'\n#{self.class.send(:hresult_detail, hr)}"
    end

    @ptr = mqi[W::PTR_SIZE, W::PTR_SIZE].unpack1(W::PACK_PTR)
    install_finalizer
  end

  def install_finalizer
    ptr = @ptr
    release_fn = W.vtable_function(ptr, 2, [W::VOIDP], W::DWORD)
    ObjectSpace.define_finalizer(self, self.class.finalizer(ptr, release_fn))
  end

  def self.finalizer(ptr, release_fn)
    proc { release_fn.call(ptr) unless ptr.zero? }
  end

  DISP_E_EXCEPTION = -2147352567 # 0x80020009

  # A Symbol (the normal path) or a String is accepted, as fole_missing's
  # rb_check_symbol + StringValueCStr does; anything else is a TypeError.
  def method_missing(name, *args)
    mname = name.is_a?(::Symbol) ? name.to_s : self.class.send(:string_value, name)
    plan = W.dispatch_plan(mname, args)
    dispid = dispid_for(plan[:name])
    if dispid.nil?
      if plan[:named_put]
        raise WIN32OLE::RuntimeError, W.property_put_error_message(plan[:name], "\nunknown property or method: `#{plan[:name]}'")
      end
      return super(name, *args)
    end

    args, named_args = split_named_args(plan[:name], args) unless plan[:named_put]
    hr, result_bytes, excepinfo_bytes =
      ole_invoke(dispid, args, plan[:wflags], named_put: plan[:named_put], named_args: named_args)

    if W.failed?(hr)
      detail = error_detail(hr, excepinfo_bytes)
      message = plan[:named_put] ? W.property_put_error_message(plan[:name], detail)
                                  : W.method_error_message(plan[:name], detail)
      raise WIN32OLE::RuntimeError, message
    end

    return nil if plan[:named_put]

    self.class.variant_bytes_to_ruby_value(result_bytes)
  end

  def respond_to_missing?(name, include_private = false)
    !dispid_for(name.to_s.sub(/=\z/, '')).nil? || super
  end

  public

  # ext/win32ole/win32ole.c's fole_invoke: the explicit-call form of
  # method_missing's regular (non property-put) path -- lets a caller
  # invoke a method whose name collides with a Ruby method (Object#send
  # works too, but this matches MRI's documented API).
  def invoke(name, *args)
    dispid = dispid_for(name.to_s)
    return super if dispid.nil?

    args, named_args = split_named_args(name.to_s, args)
    hr, result_bytes, excepinfo_bytes =
      ole_invoke(dispid, args, W::DISPATCH_METHOD | W::DISPATCH_PROPERTYGET, named_args: named_args)

    if W.failed?(hr)
      raise WIN32OLE::RuntimeError, W.method_error_message(name, error_detail(hr, excepinfo_bytes))
    end

    self.class.variant_bytes_to_ruby_value(result_bytes)
  end

  # ext/win32ole/win32ole.c's fole_invoke2/fole_getproperty2/fole_setproperty2:
  # the "early binding" siblings of #invoke/#setproperty -- dispid and
  # per-argument VARTYPE are given explicitly instead of resolved by name,
  # so no GetIDsOfNames round-trip happens.
  def _invoke(dispid, args, types)
    ole_invoke2(dispid, args, types, W::DISPATCH_METHOD)
  end

  def _getproperty(dispid, args, types)
    ole_invoke2(dispid, args, types, W::DISPATCH_PROPERTYGET)
  end

  def _setproperty(dispid, args, types)
    ole_invoke2(dispid, args, types, W::DISPATCH_PROPERTYPUT)
  end

  # ext/win32ole/win32ole.c's fole_setproperty: like the `name=(val)`
  # method_missing path, but for properties that also take index
  # arguments (e.g. sheet.setproperty('Cells', 1, 2, 10)) -- the last
  # argument is the value, everything before it is an index arg, and (per
  # ole_invoke) they're all invoked as one DISPATCH_PROPERTYPUT call with
  # the value carried as the DISPID_PROPERTYPUT named argument.
  def setproperty(name, *args)
    if args.empty?
      raise WIN32OLE::RuntimeError, W.property_put_error_message(name, "\nargument error")
    end

    dispid = dispid_for(name.to_s)
    return super if dispid.nil?

    hr, _result_bytes, excepinfo_bytes = ole_invoke(dispid, args, W::DISPATCH_PROPERTYPUT, named_put: true)

    if W.failed?(hr)
      raise WIN32OLE::RuntimeError, W.property_put_error_message(name, error_detail(hr, excepinfo_bytes))
    end

    nil
  end

  # ext/win32ole/win32ole.c's fole_each: gets the collection's enumerator
  # through DISPID_NEWENUM, QueryInterfaces it for IEnumVARIANT, and yields
  # each item IEnumVARIANT::Next returns until it stops returning S_OK.
  def each
    return to_enum(:each) unless block_given?

    hr, result_bytes, = ole_invoke(W::DISPID_NEWENUM, [], W::DISPATCH_METHOD | W::DISPATCH_PROPERTYGET)
    if W.failed?(hr)
      raise WIN32OLE::QueryInterfaceError, "failed to get IEnum Interface\n#{self.class.send(:hresult_detail, hr)}"
    end

    vt, payload = W.unpack_variant(result_bytes)
    unknown = [W::VT_UNKNOWN, W::VT_DISPATCH].include?(vt) ? W.unpack_pointer(payload) : 0
    raise WIN32OLE::QueryInterfaceError, 'failed to get IEnum Interface' if unknown.zero?

    enum = W.query_interface(unknown, W::IID_IENUMVARIANT)
    self.class.send(:release_com_pointer, unknown)
    raise WIN32OLE::QueryInterfaceError, 'failed to get IEnum Interface' if enum.nil?

    begin
      next_fn = W.vtable_function(enum, 3, [W::VOIDP, W::DWORD, W::VOIDP, W::VOIDP], W::LONG)
      loop do
        var = ("\x00" * W::VARIANT_SIZE).b
        break unless next_fn.call(enum, 1, var, nil).zero? # S_OK; S_FALSE (1) = no more

        yield self.class.variant_bytes_to_ruby_value(var)
      end
    ensure
      self.class.send(:release_com_pointer, enum)
    end
    nil
  end

  # ext/win32ole/win32ole.c's fole_getproperty_with_bracket: invokes the
  # object's default member (DISPID_VALUE) with the given arguments --
  # e.g. dict['ruby'] on a Scripting.Dictionary, whose default property
  # is Item. No name lookup happens here (unlike method_missing/#invoke):
  # all arguments are passed straight through to DISPID_VALUE.
  def [](*args)
    hr, result_bytes, excepinfo_bytes = ole_invoke(W::DISPID_VALUE, args, W::DISPATCH_PROPERTYGET)
    if W.failed?(hr)
      raise WIN32OLE::RuntimeError, W.method_error_message(args.first, error_detail(hr, excepinfo_bytes))
    end

    self.class.variant_bytes_to_ruby_value(result_bytes)
  end

  # ext/win32ole/win32ole.c's fole_setproperty_with_bracket: same
  # DISPID_VALUE target as #[], but DISPATCH_PROPERTYPUT -- the last
  # argument is the value, everything before it is an index argument
  # (e.g. dict[2] = 'TWO').
  def []=(*args)
    hr, _result_bytes, excepinfo_bytes = ole_invoke(W::DISPID_VALUE, args, W::DISPATCH_PROPERTYPUT, named_put: true)
    if W.failed?(hr)
      raise WIN32OLE::RuntimeError, W.property_put_error_message(args.first, error_detail(hr, excepinfo_bytes))
    end

    args.last
  end

  IPERSISTMEMORY_INIT_NEW = 8 # IUnknown(3) + IPersist(1) + IsDirty/Load/Save/GetSizeMax

  # fole_activex_initialize: initialize an ActiveX control through its
  # IPersistMemory::InitNew.
  def ole_activex_initialize
    persist = W.query_interface(@ptr, W::IID_IPERSISTMEMORY)
    hr = -2147467262 # E_NOINTERFACE
    if persist
      begin
        hr = W.vtable_function(persist, IPERSISTMEMORY_INIT_NEW, [W::VOIDP], W::LONG).call(persist)
      ensure
        self.class.send(:release_com_pointer, persist)
      end
    end
    if W.failed?(hr)
      raise WIN32OLE::RuntimeError, "fail to initialize ActiveX control\n#{self.class.send(:hresult_detail, hr)}"
    end

    nil
  end

  # fole_free: Release the IDispatch now instead of at GC time.
  def ole_free
    release
    nil
  end

  # WIN32OLE::Event needs this raw IDispatch* to build its own COM
  # connections (QueryInterface for IConnectionPointContainer,
  # IProvideClassInfo2, etc.) -- #ole_query_interface wraps whatever it
  # gets back as a WIN32OLE, which is not what Event needs, so this is
  # Event's one deliberate, documented crack in encapsulation.
  def dispatch_ptr
    @ptr
  end

  def ole_type
    type_info_ptr = get_type_info_ptr
    raise WIN32OLE::QueryInterfaceError, 'failed to GetTypeInfo' if type_info_ptr.nil?

    Type.from_typeinfo_ptr(type_info_ptr)
  end

  def ole_typelib
    ole_type.ole_typelib
  end

  def ole_methods
    ole_methods_by_invkind(nil)
  end

  def ole_get_methods
    ole_methods_by_invkind(TypeInfo::INVOKE_PROPERTYGET)
  end

  def ole_put_methods
    ole_methods_by_invkind(TypeInfo::INVOKE_PROPERTYPUT | TypeInfo::INVOKE_PROPERTYPUTREF)
  end

  def ole_func_methods
    ole_methods_by_invkind(TypeInfo::INVOKE_FUNC)
  end

  def ole_respond_to?(name)
    !dispid_for(name.to_s).nil?
  end

  # fole_method_help: the WIN32OLE::Method named +name+ on this object's
  # type. (ole_method_help is an alias of it, and ole_obj_help one of
  # ole_type -- neither launches a help file; only WIN32OLE.ole_show_help
  # does.)
  def ole_method(name)
    WIN32OLE::Method.new(type_via_containing_typelib, self.class.send(:string_value, name))
  end
  alias ole_method_help ole_method
  alias ole_obj_help ole_type

  # ext/win32ole/win32ole.c's fole_query_interface: like C, the returned
  # interface pointer is wrapped as a WIN32OLE (i.e. treated as IDispatch*).
  def ole_query_interface(iid)
    iid = self.class.send(:string_value, iid)
    raise WIN32OLE::RuntimeError, 'this WIN32OLE object has already been released' if @ptr.nil? || @ptr.zero?

    iid_bytes = ("\x00" * 16).b
    hr = W.clsid_from_string.call(W.wstr(iid), iid_bytes)
    if W.failed?(hr)
      raise WIN32OLE::RuntimeError, "invalid iid: `#{iid}'\n#{self.class.send(:hresult_detail, hr)}"
    end

    ppv = ("\x00" * W::PTR_SIZE).b
    hr = W.vtable_function(@ptr, 0, [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG).call(@ptr, iid_bytes, ppv)
    if W.failed?(hr)
      raise WIN32OLE::QueryInterfaceError, "failed to get interface `#{iid}'\n#{self.class.send(:hresult_detail, hr)}"
    end

    self.class.wrap_dispatch_pointer(ppv.unpack1(W::PACK_PTR))
  end

  private

  # ext/win32ole/win32ole.c's ole_invoke2: builds one DISPPARAMS entry per
  # (value, VARTYPE) pair -- VT_VARIANT means "pass as its own natural
  # type" (no coercion), matching C's skip of VariantChangeTypeEx for that
  # case; anything else is packed via WIN32OLE::Variant.new(val, vt),
  # which already implements the explicit-VARTYPE coercion (including
  # VT_ARRAY/VT_BYREF) that this needs. The Variant objects are kept alive
  # in keep_alive through the Invoke call, since VT_BYREF's backing buffer
  # is otherwise only referenced from their own instance state.
  def ole_invoke2(dispid, args, types, dispkind)
    raise WIN32OLE::RuntimeError, 'this WIN32OLE object has already been released' if @ptr.nil? || @ptr.zero?
    raise TypeError, 'wrong argument type (expected Array)' unless args.is_a?(::Array) && types.is_a?(::Array)

    to_free = []
    keep_alive = []
    arg_variants = args.zip(types).map do |val, vt|
      if vt == W::VT_VARIANT
        WIN32OLE.ruby_value_to_variant_bytes(val, to_free)
      else
        variant = WIN32OLE::Variant.new(val, vt)
        keep_alive << variant
        variant.instance_variable_get(:@var)
      end
    end.reverse
    args_blob = arg_variants.join
    args_ptr = args_blob.empty? ? nil : W.native_pointer_for(args_blob)

    named_put = (dispkind & W::DISPATCH_PROPERTYPUT) != 0
    named_blob = named_put ? [W::DISPID_PROPERTYPUT].pack('l') : ''
    named_ptr = named_blob.empty? ? nil : W.native_pointer_for(named_blob)

    dispparams = [
      args_ptr ? args_ptr.to_i : 0,
      named_ptr ? named_ptr.to_i : 0,
      arg_variants.size,
      named_put ? 1 : 0
    ].pack("#{W::PACK_PTR}#{W::PACK_PTR}LL")

    excepinfo = ("\x00" * W::EXCEPINFO_SIZE).b
    result = ("\x00" * W::VARIANT_SIZE).b

    hr = invoke_fn.call(@ptr, dispid, W::IID_NULL, 0, dispkind, dispparams, result, excepinfo, nil)
    to_free.each(&:call)

    if W.failed?(hr)
      raise WIN32OLE::RuntimeError, W.method_error_message("<dispatch id:#{dispid}>", error_detail(hr, excepinfo))
    end

    self.class.variant_bytes_to_ruby_value(result)
  end

  def error_detail(hr, excepinfo_bytes)
    if hr == DISP_E_EXCEPTION
      info = W.parse_excepinfo(excepinfo_bytes)
      source = W.bstr_to_s(info[:bstr_source_ptr]) || '<Unknown>'
      description = W.bstr_to_s(info[:bstr_description_ptr]) || '<No Description>'
      code = info[:w_code].zero? ? (info[:scode] & 0xFFFFFFFF).to_s(16).upcase : info[:w_code].to_s
      "\n    OLE error code:#{code} in #{source}\n      #{description}\n#{self.class.send(:hresult_detail, hr)}"
    else
      "\n#{self.class.send(:hresult_detail, hr)}"
    end
  end

  def get_type_info_ptr
    out = ("\x00" * W::PTR_SIZE).b
    hr = TypeInfo.get_type_info_fn(@ptr).call(@ptr, 0, W::LOCALE_SYSTEM_DEFAULT, out)
    return nil if W.failed?(hr)

    out.unpack1(W::PACK_PTR)
  end

  def ole_methods_by_invkind(mask)
    type_via_containing_typelib.ole_methods.select do |m|
      mask.nil? || (m.invkind & mask) != 0
    end
  end

  # Ports ext/win32ole/win32ole.c's typeinfo_from_ole: GetTypeInfo →
  # GetDocumentation (this type's own name) → GetContainingTypeLib →
  # scan the typelib for the entry with a matching name → GetTypeInfo(i)
  # again. See design spec §1.1/§8 risk #3 for why this round-trip exists
  # instead of just reusing the first ITypeInfo* the way #ole_type does —
  # ported as-is rather than "simplified" without understanding it.
  def type_via_containing_typelib
    first_type = ole_type
    target_name = first_type.name
    tlib = first_type.ole_typelib
    match = tlib.ole_types.find { |t| t.name == target_name }
    match || first_type
  end
end
