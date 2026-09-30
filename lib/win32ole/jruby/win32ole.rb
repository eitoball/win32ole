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

  @lcid = LOCALE_SYSTEM_DEFAULT

  class << self
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

    def wrap_dispatch_pointer(ptr)
      obj = allocate
      obj.instance_variable_set(:@ptr, ptr)
      obj.send(:install_finalizer)
      obj
    end

    def connect(server, host = nil)
      raise NotImplementedError, 'remote OLE (host) is not supported yet' unless host.nil?

      hr = W.co_initialize.call(nil)
      raise 'fail: OLE initialize' unless hr.zero? || hr == 1

      clsid = resolve_clsid(server)
      ppv = ("\x00" * W::PTR_SIZE).b
      hr = W.get_active_object.call(clsid, nil, ppv)
      if W.failed?(hr)
        raise WIN32OLE::RuntimeError, "#{W.unknown_server_error_message(server)}\n#{hresult_detail(hr)}"
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

    def ruby_value_to_variant_bytes(value, bstrs_to_free)
      case value
      when ::Array
        psa = SafeArray.ruby_array_to_safearray(value, W::VT_VARIANT, bstrs_to_free)
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
          bstrs_to_free << bstr
          W.pack_pointer(bstr)
        when :dispatch
          W.pack_pointer(value.instance_variable_get(:@ptr))
        end
      W.pack_variant(W::VT_FOR_TYPE.fetch(type), payload)
    end

    def variant_bytes_to_ruby_value(bytes)
      vt, = W.unpack_variant(bytes)
      base_vt = vt & W::VT_TYPEMASK

      if (vt & W::VT_ARRAY) != 0
        _vt, payload = W.unpack_variant(bytes)
        psa = W.unpack_pointer(payload)
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
      when :bstr
        addr = W.unpack_pointer(payload)
        str = W.bstr_to_s(addr)
        W.sys_free_string.call(addr) unless addr.zero?
        str
      when :dispatch
        ptr = W.unpack_pointer(payload)
        ptr.zero? ? nil : wrap_dispatch_pointer(ptr)
      end
    end

    def resolve_clsid(server)
      wide = W.wstr(server)
      clsid = ("\x00" * 16).b
      hr = W.clsid_from_progid.call(wide, clsid)
      hr = W.clsid_from_string.call(wide, clsid) if W.failed?(hr)
      if W.failed?(hr)
        raise WIN32OLE::RuntimeError, "#{W.unknown_server_error_message(server)}\n#{hresult_detail(hr)}"
      end

      clsid
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

    def const_load(ole, mod)
      ole.ole_type.ole_typelib.ole_types.each do |type|
        type.variables.each do |var|
          next unless var.variable_kind == 'CONSTANT'
          next if mod.const_defined?(var.name, false)

          mod.const_set(var.name, var.value)
        end
      end
      nil
    end

    private :resolve_clsid, :hresult_detail, :release_com_pointer
  end

  def initialize(server, host = nil)
    raise NotImplementedError, 'remote OLE (host) is not supported yet' unless host.nil?

    hr = W.co_initialize.call(nil)
    raise 'fail: OLE initialize' unless hr.zero? || hr == 1

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

  def install_finalizer
    ptr = @ptr
    release_fn = W.vtable_function(ptr, 2, [W::VOIDP], W::DWORD)
    ObjectSpace.define_finalizer(self, self.class.finalizer(ptr, release_fn))
  end

  def self.finalizer(ptr, release_fn)
    proc { release_fn.call(ptr) unless ptr.zero? }
  end

  DISP_E_EXCEPTION = -2147352567 # 0x80020009

  def method_missing(name, *args)
    plan = W.dispatch_plan(name.to_s, args)
    dispid = dispid_for(plan[:name])
    if dispid.nil?
      if plan[:named_put]
        raise WIN32OLE::RuntimeError, W.property_put_error_message(plan[:name], "\nunknown property or method: `#{plan[:name]}'")
      end
      return super(name, *args)
    end

    hr, result_bytes, excepinfo_bytes = ole_invoke(dispid, args, plan[:wflags], named_put: plan[:named_put])

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

    hr, result_bytes, excepinfo_bytes = ole_invoke(dispid, args, W::DISPATCH_METHOD | W::DISPATCH_PROPERTYGET)

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

  # WIN32OLE::Event needs this raw IDispatch* to build its own COM
  # connections (QueryInterface for IConnectionPointContainer,
  # IProvideClassInfo2, etc.) -- there is no general-purpose
  # QueryInterface API (#ole_query_interface is a Phase 2 non-goal), so
  # this is Event's one deliberate, documented crack in encapsulation.
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

  def ole_method_help(*)
    raise NotImplementedError, 'launching help files is not implemented (Phase 2 non-goal)'
  end

  def ole_obj_help
    raise NotImplementedError, 'launching help files is not implemented (Phase 2 non-goal)'
  end

  def ole_query_interface(*)
    raise NotImplementedError, 'arbitrary QueryInterface is not implemented (Phase 2 non-goal)'
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

    bstrs_to_free = []
    keep_alive = []
    arg_variants = args.zip(types).map do |val, vt|
      if vt == W::VT_VARIANT
        WIN32OLE.ruby_value_to_variant_bytes(val, bstrs_to_free)
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
    bstrs_to_free.each { |bstr| W.sys_free_string.call(bstr) unless bstr.zero? }

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
