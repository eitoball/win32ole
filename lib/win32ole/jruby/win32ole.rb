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

  class << self
    def wrap_dispatch_pointer(ptr)
      obj = allocate
      obj.instance_variable_set(:@ptr, ptr)
      obj.send(:install_finalizer)
      obj
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
  end

  def initialize(server, host = nil)
    raise NotImplementedError, 'remote OLE (host) is not supported yet' unless host.nil?

    hr = W.co_initialize.call(nil)
    raise 'fail: OLE initialize' unless hr.zero? || hr == 1

    clsid = resolve_clsid(server)
    ppv = ("\x00" * W::PTR_SIZE).b
    hr = W.co_create_instance.call(
      clsid, nil, W::CLSCTX_INPROC_SERVER | W::CLSCTX_LOCAL_SERVER, W::IID_IDISPATCH, ppv
    )
    if W.failed?(hr)
      raise WIN32OLE::RuntimeError, "#{W.unknown_server_error_message(server)}\n#{hresult_detail(hr)}"
    end

    @ptr = ppv.unpack1(W::PTR_SIZE == 8 ? 'Q' : 'L')
    install_finalizer
  end

  private

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

  def error_detail(hr, excepinfo_bytes)
    if hr == DISP_E_EXCEPTION
      info = W.parse_excepinfo(excepinfo_bytes)
      source = W.bstr_to_s(info[:bstr_source_ptr]) || '<Unknown>'
      description = W.bstr_to_s(info[:bstr_description_ptr]) || '<No Description>'
      code = info[:w_code].zero? ? (info[:scode] & 0xFFFFFFFF).to_s(16).upcase : info[:w_code].to_s
      "\n    OLE error code:#{code} in #{source}\n      #{description}\n#{hresult_detail(hr)}"
    else
      "\n#{hresult_detail(hr)}"
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
