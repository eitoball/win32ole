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
end
