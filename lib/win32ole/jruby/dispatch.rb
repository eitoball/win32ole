# lib/win32ole/jruby/dispatch.rb
require 'win32ole/jruby/win32'

class WIN32OLE
  module Dispatch
    W = WIN32OLE::Win32

    def dispid_for(name)
      ids = dispids_for([name])
      ids && ids.first
    end

    # One GetIDsOfNames call for several names. Given a method name followed
    # by parameter names, the parameters are resolved in that method's
    # scope (what named arguments need). nil if the call fails.
    def dispids_for(names)
      raise WIN32OLE::RuntimeError, 'this WIN32OLE object has already been released' if @ptr.nil? || @ptr.zero?

      name_bufs = names.map { |name| W.wstr(name) }
      name_ptrs = name_bufs.map { |buf| W.native_pointer_for(buf) }
      names_blob = name_ptrs.map(&:to_i).pack("#{W::PACK_PTR}*")
      dispids = ("\x00" * (4 * names.size)).b
      hr = get_ids_of_names_fn.call(@ptr, W::IID_NULL, names_blob, names.size, 0, dispids)
      return nil if W.failed?(hr)

      dispids.unpack('l*')
    end

    # ole_invoke's treatment of a trailing Hash: its entries become named
    # arguments, their DISPIDs looked up together with the method name.
    # Returns [positional_args, named_args], named_args being
    # [[dispid, value], ...] or nil when there is no trailing Hash.
    def split_named_args(name, args)
      hash = args.last
      return [args, nil] unless hash.is_a?(::Hash)

      keys = hash.keys.map do |key|
        case key
        when ::String then key
        when ::Symbol then key.to_s
        else raise TypeError, 'wrong argument type (expected String or Symbol)'
        end
      end
      ids = dispids_for([name, *keys])
      raise WIN32OLE::RuntimeError, "failed to get named argument info: `#{name}'" if ids.nil?

      [args[0...-1], ids.drop(1).zip(hash.values)]
    end

    # Returns [hr, result_variant_bytes, excepinfo_bytes]. Never raises —
    # callers build the user-facing error message (they know whether this
    # was a method call or a property-put, which changes the message).
    # named_args ([[dispid, value], ...]) occupy the front of rgvarg, in
    # order, ahead of the reversed positional arguments -- the layout
    # ole_invoke builds; named_put marks the value as DISPID_PROPERTYPUT.
    def ole_invoke(dispid, arg_values, wflags, named_put: false, named_args: nil)
      raise WIN32OLE::RuntimeError, 'this WIN32OLE object has already been released' if @ptr.nil? || @ptr.zero?

      keep_alive = []
      bstrs_to_free = []

      named_args ||= []
      values = named_args.map(&:last) + arg_values.reverse
      arg_variants = values.map { |v| WIN32OLE.ruby_value_to_variant_bytes(v, bstrs_to_free) }
      args_blob = arg_variants.join
      args_ptr = args_blob.empty? ? nil : W.native_pointer_for(args_blob)
      keep_alive << args_ptr if args_ptr

      named_ids = named_put ? [W::DISPID_PROPERTYPUT] : named_args.map(&:first)
      named_blob = named_ids.pack('l*')
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
        named_ids.size
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
