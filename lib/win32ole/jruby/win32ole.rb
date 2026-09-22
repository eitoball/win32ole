# lib/win32ole/jruby/win32ole.rb
require 'fiddle'
require 'win32ole/jruby/win32'
require 'win32ole/jruby/dispatch'

class WIN32OLE
  RuntimeError = Class.new(::RuntimeError)

  include Dispatch

  W = Win32
  private_constant :W

  def initialize(server, host = nil)
    raise NotImplementedError, 'remote OLE (host) is not supported yet' unless host.nil?

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

  def native_buffers
    @native_buffers ||= []
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

    variant_bytes_to_ruby_value(result_bytes)
  end

  def respond_to_missing?(name, include_private = false)
    !dispid_for(name.to_s.sub(/=\z/, '')).nil? || super
  end

  def ruby_value_to_variant_bytes(value, keep_alive)
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
        keep_alive << bstr
        W.pack_pointer(bstr)
      when :dispatch
        W.pack_pointer(value.instance_variable_get(:@ptr))
      end
    W.pack_variant(W::VT_FOR_TYPE.fetch(type), payload)
  end

  def variant_bytes_to_ruby_value(bytes)
    vt, payload = W.unpack_variant(bytes)
    type = W.variant_ruby_type(vt)
    case type
    when :empty then nil
    when :i4 then W.unpack_i4(payload)
    when :i8 then W.unpack_i8(payload)
    when :r8 then W.unpack_r8(payload)
    when :bool then W.unpack_bool(payload)
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

  private

  def wrap_dispatch_pointer(ptr)
    obj = self.class.allocate
    obj.instance_variable_set(:@ptr, ptr)
    obj.send(:install_finalizer)
    obj
  end

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
end
