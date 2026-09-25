# lib/win32ole/jruby/dispatch.rb
require 'win32ole/jruby/win32'

class WIN32OLE
  module Dispatch
    W = WIN32OLE::Win32

    def dispid_for(name)
      raise WIN32OLE::RuntimeError, 'this WIN32OLE object has already been released' if @ptr.nil? || @ptr.zero?

      name_buf = W.wstr(name)
      name_ptr = W.native_pointer_for(name_buf)
      names = [name_ptr.to_i].pack(W::PACK_PTR)
      dispids = ("\x00" * 4).b
      hr = get_ids_of_names_fn.call(@ptr, W::IID_NULL, names, 1, 0, dispids)
      return nil if W.failed?(hr)

      dispids.unpack1('l')
    end

    # Returns [hr, result_variant_bytes, excepinfo_bytes]. Never raises —
    # callers build the user-facing error message (they know whether this
    # was a method call or a property-put, which changes the message).
    def ole_invoke(dispid, arg_values, wflags, named_put: false)
      raise WIN32OLE::RuntimeError, 'this WIN32OLE object has already been released' if @ptr.nil? || @ptr.zero?

      keep_alive = []
      bstrs_to_free = []

      arg_variants = arg_values.reverse.map { |v| WIN32OLE.ruby_value_to_variant_bytes(v, bstrs_to_free) }
      args_blob = arg_variants.join
      args_ptr = args_blob.empty? ? nil : W.native_pointer_for(args_blob)
      keep_alive << args_ptr if args_ptr

      named_blob = named_put ? [W::DISPID_PROPERTYPUT].pack('l') : ''
      named_ptr = named_blob.empty? ? nil : W.native_pointer_for(named_blob)
      keep_alive << named_ptr if named_ptr

      # DISPPARAMS is a real native struct (VARIANTARG *rgvarg; DISPID
      # *rgdispidNamedArgs; UINT cArgs; UINT cNamedArgs;) read positionally
      # by the callee, so — unlike a VARIANT's self-designed padding — the
      # two pointer fields MUST be packed at native pointer width (PACK_PTR),
      # not a hardcoded 8 bytes, or this silently misaligns cArgs/cNamedArgs
      # on x86.
      dispparams = [
        args_ptr ? args_ptr.to_i : 0,
        named_ptr ? named_ptr.to_i : 0,
        arg_variants.size,
        named_put ? 1 : 0
      ].pack("#{W::PACK_PTR}#{W::PACK_PTR}LL")
      keep_alive << dispparams

      excepinfo = ("\x00" * W::EXCEPINFO_SIZE).b
      result = ("\x00" * W::VARIANT_SIZE).b

      hr = invoke_fn.call(@ptr, dispid, W::IID_NULL, 0, wflags, dispparams, result, excepinfo, nil)
      bstrs_to_free.each { |bstr| W.sys_free_string.call(bstr) }
      [hr, result, excepinfo]
    end

    def release
      return unless @ptr && !@ptr.zero?

      release_fn.call(@ptr)
      @ptr = nil
      ObjectSpace.undefine_finalizer(self)
    end

    private

    def get_ids_of_names_fn
      @get_ids_of_names_fn ||= W.vtable_function(
        @ptr, 5, [W::VOIDP, W::VOIDP, W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def invoke_fn
      @invoke_fn ||= W.vtable_function(
        @ptr, 6,
        [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::WORD, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP],
        W::LONG
      )
    end

    def release_fn
      @release_fn ||= W.vtable_function(@ptr, 2, [W::VOIDP], W::DWORD)
    end
  end
end
