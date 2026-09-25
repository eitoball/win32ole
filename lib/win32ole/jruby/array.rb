# lib/win32ole/jruby/array.rb
require 'fiddle'
require 'win32ole/jruby/win32'

class WIN32OLE
  module SafeArray
    W = Win32
    private_constant :W

    module_function

    def oleaut32
      W.oleaut32
    end

    def safe_array_create
      @safe_array_create ||= Fiddle::Function.new(
        oleaut32['SafeArrayCreate'], [W::WORD, W::DWORD, W::VOIDP], W::VOIDP, W::STDCALL
      )
    end

    def safe_array_create_vector
      @safe_array_create_vector ||= Fiddle::Function.new(
        oleaut32['SafeArrayCreateVector'], [W::WORD, W::LONG, W::DWORD], W::VOIDP, W::STDCALL
      )
    end

    def safe_array_destroy
      @safe_array_destroy ||= Fiddle::Function.new(oleaut32['SafeArrayDestroy'], [W::VOIDP], W::LONG, W::STDCALL)
    end

    def safe_array_lock
      @safe_array_lock ||= Fiddle::Function.new(oleaut32['SafeArrayLock'], [W::VOIDP], W::LONG, W::STDCALL)
    end

    def safe_array_unlock
      @safe_array_unlock ||= Fiddle::Function.new(oleaut32['SafeArrayUnlock'], [W::VOIDP], W::LONG, W::STDCALL)
    end

    def safe_array_get_dim
      @safe_array_get_dim ||= Fiddle::Function.new(oleaut32['SafeArrayGetDim'], [W::VOIDP], W::DWORD, W::STDCALL)
    end

    def safe_array_get_lbound
      @safe_array_get_lbound ||= Fiddle::Function.new(
        oleaut32['SafeArrayGetLBound'], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_get_ubound
      @safe_array_get_ubound ||= Fiddle::Function.new(
        oleaut32['SafeArrayGetUBound'], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_ptr_of_index
      @safe_array_ptr_of_index ||= Fiddle::Function.new(
        oleaut32['SafeArrayPtrOfIndex'], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_put_element
      @safe_array_put_element ||= Fiddle::Function.new(
        oleaut32['SafeArrayPutElement'], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_access_data
      @safe_array_access_data ||= Fiddle::Function.new(
        oleaut32['SafeArrayAccessData'], [W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_unaccess_data
      @safe_array_unaccess_data ||= Fiddle::Function.new(
        oleaut32['SafeArrayUnaccessData'], [W::VOIDP], W::LONG, W::STDCALL
      )
    end

    # Port of ext/win32ole/win32ole.c's dimension() (line 1148) -- the
    # nesting depth is the MAX depth found across every branch, not just
    # the first element, so a ragged/mixed input still gets a well-defined
    # depth.
    def dimension_count(val)
      return 0 unless val.is_a?(::Array)

      val.reduce(0) { |max, v| [max, dimension_count(v)].max } + 1
    end

    # Port of ary_len_of_dim() (line 1167) -- the size at nesting level
    # `dim` (0-indexed, 0 == outermost) is the MAX size found across every
    # branch at that depth.
    def dimension_size(val, dim)
      return 0 unless val.is_a?(::Array)
      return val.size if dim.zero?

      val.reduce(0) { |max, v| [max, dimension_size(v, dim - 1)].max }
    end

    def dimension_sizes(val)
      dims = dimension_count(val)
      Array.new(dims) { |d| dimension_size(val, d) }
    end

    # Port of ole_ary_m_entry() (line 963) -- pid[0] indexes the outermost
    # Array, pid[1] the next level in, etc. Returns nil past a shorter
    # (ragged) branch, matching Ruby's own Array#[] out-of-range behavior.
    def nested_entry(val, pid)
      obj = val
      pid.each { |i| obj = obj.is_a?(::Array) ? obj[i] : nil }
      obj
    end

    # Port of ole_set_safe_array()'s pid increment loop (line 1116): the
    # LAST index (the innermost Ruby nesting level) varies fastest,
    # carrying left on overflow. This is the order Array->SAFEARRAY fill
    # (Task 6) writes elements in.
    def each_fill_index(sizes)
      return enum_for(:each_fill_index, sizes) unless block_given?

      pid = Array.new(sizes.size, 0)
      loop do
        yield pid.dup
        i = sizes.size - 1
        loop do
          pid[i] += 1
          break if pid[i] < sizes[i]

          pid[i] = 0
          i -= 1
          return if i.negative?
        end
      end
    end

    # Port of ole_variant2val()'s pid increment loop (line 1476): the
    # FIRST index (SAFEARRAY's own dimension 0) varies fastest --
    # deliberately the opposite order of each_fill_index above (see this
    # task's own note on why). lbounds/ubounds are inclusive, one entry
    # per dimension, as SafeArrayGetLBound/GetUBound return them.
    def each_read_index(lbounds, ubounds)
      return enum_for(:each_read_index, lbounds, ubounds) unless block_given?

      pid = lbounds.dup
      loop do
        yield pid.dup
        i = 0
        loop do
          pid[i] += 1
          break if pid[i] <= ubounds[i]

          pid[i] = lbounds[i]
          i += 1
          return if i == pid.size
        end
      end
    end

    # Pointer-typed elements (VT_BSTR, VT_DISPATCH, VT_UNKNOWN) are
    # deliberately absent here: unlike a true scalar, a pointer element
    # needs real marshaling (BSTR allocation/decoding, IDispatch/IUnknown
    # refcounting/wrapping) to become a meaningful Ruby value or back --
    # that marshaling only exists today inside the VT_VARIANT branch (via
    # WIN32OLE.ruby_value_to_variant_bytes/.variant_bytes_to_ruby_value).
    # A directly-BSTR/Dispatch/Unknown-typed SAFEARRAY element raises
    # NotImplementedError (via the .fetch fallback below) until that
    # marshaling is built for the non-VARIANT scalar path, rather than
    # packing/unpacking the wrong thing.
    ELEMENT_PACK_FORMAT = {
      W::VT_I1 => 'c', W::VT_UI1 => 'C', W::VT_I2 => 's', W::VT_UI2 => 'S',
      W::VT_I4 => 'l', W::VT_UI4 => 'L', W::VT_INT => 'l', W::VT_UINT => 'L',
      W::VT_I8 => 'q', W::VT_UI8 => 'Q', W::VT_R4 => 'f', W::VT_R8 => 'd',
      W::VT_ERROR => 'l', W::VT_BOOL => 's'
    }.freeze

    def pack_scalar_element(vt, value)
      fmt = ELEMENT_PACK_FORMAT.fetch(vt) { raise NotImplementedError, "VARTYPE #{vt} is not a supported array element type yet" }
      [value].pack(fmt)
    end

    def unpack_scalar_element(vt, bytes)
      fmt = ELEMENT_PACK_FORMAT.fetch(vt) { raise NotImplementedError, "VARTYPE #{vt} is not a supported array element type yet" }
      bytes.unpack1(fmt)
    end

    def ruby_array_to_safearray(ary, elem_vt, bstrs_to_free = [])
      base_vt = elem_vt & W::VT_TYPEMASK
      return ui1_safearray_from_bytes(ary) if base_vt == W::VT_UI1 && ary.is_a?(::String)

      sizes = dimension_sizes(ary)
      dims = sizes.size
      bounds = sizes.flat_map { |n| [n, 0] }.pack('L2' * dims)
      psa = safe_array_create.call(base_vt, dims, bounds)
      raise ::RuntimeError, 'memory allocation error' if psa.nil? || psa.to_i.zero?

      hr = safe_array_lock.call(psa)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayLock: #{W.hr_hex(hr)}" if W.failed?(hr)

      begin
        each_fill_index(sizes) do |pid|
          val = nested_entry(ary, pid)
          leaf = base_vt == W::VT_VARIANT ? WIN32OLE.ruby_value_to_variant_bytes(val, bstrs_to_free)
                                           : pack_scalar_element(base_vt, val)
          index_buf = pid.pack('l' * dims)
          hr = safe_array_put_element.call(psa, index_buf, W.native_pointer_for(leaf))
          raise WIN32OLE::RuntimeError, "failed to SafeArrayPutElement: #{W.hr_hex(hr)}" if W.failed?(hr)
        end
      ensure
        safe_array_unlock.call(psa)
      end
      psa
    end

    def safearray_to_ruby_array(psa, elem_vt)
      base_vt = elem_vt & W::VT_TYPEMASK
      return ui1_safearray_to_bytes(psa) if base_vt == W::VT_UI1

      dim = safe_array_get_dim.call(psa)
      lbounds = Array.new(dim) { |d| out = ("\x00" * 4).b; safe_array_get_lbound.call(psa, d + 1, out); out.unpack1('l') }
      ubounds = Array.new(dim) { |d| out = ("\x00" * 4).b; safe_array_get_ubound.call(psa, d + 1, out); out.unpack1('l') }

      hr = safe_array_lock.call(psa)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayLock: #{W.hr_hex(hr)}" if W.failed?(hr)

      result = nested_array_skeleton(ubounds.zip(lbounds).map { |u, l| u - l + 1 })
      begin
        each_read_index(lbounds, ubounds) do |pid|
          index_buf = pid.pack('l' * dim)
          elem_ptr_out = ("\x00" * W::PTR_SIZE).b
          hr = safe_array_ptr_of_index.call(psa, index_buf, elem_ptr_out)
          raise WIN32OLE::RuntimeError, "failed to SafeArrayPtrOfIndex: #{W.hr_hex(hr)}" if W.failed?(hr)

          elem_addr = elem_ptr_out.unpack1(W::PACK_PTR)
          val =
            if base_vt == W::VT_VARIANT
              WIN32OLE.variant_bytes_to_ruby_value(Fiddle::Pointer.new(elem_addr)[0, W::VARIANT_SIZE])
            else
              size = ELEMENT_PACK_FORMAT.fetch(base_vt) { raise NotImplementedError, "VARTYPE #{base_vt} is not a supported array element type yet" }
              unpack_scalar_element(base_vt, Fiddle::Pointer.new(elem_addr)[0, [1].pack(size).bytesize])
            end
          zero_based_pid = pid.each_with_index.map { |v, d| v - lbounds[d] }
          store_nested(result, zero_based_pid, val)
        end
      ensure
        safe_array_unlock.call(psa)
      end
      result
    end

    def nested_array_skeleton(sizes)
      return Array.new(sizes.first) if sizes.size == 1

      Array.new(sizes.first) { nested_array_skeleton(sizes[1..]) }
    end

    def store_nested(ary, pid, val)
      obj = ary
      pid[0..-2].each { |i| obj = obj[i] }
      obj[pid.last] = val
    end

    # VT_UI1|VT_ARRAY <-> String fast path: bulk-copy via
    # SafeArrayAccessData instead of the generic per-element path above --
    # a real, separate MRI code path (ole_val2olevariantdata's first
    # branch / folevariant_value's dim==1 reverse path), not an
    # optimization detail this design can skip: without it, a binary blob
    # argument would round-trip through an Array of small Integers instead
    # of staying a String.
    def ui1_safearray_from_bytes(bytes)
      bytes = bytes.b
      bounds = [bytes.bytesize, 0].pack('L2')
      psa = safe_array_create.call(W::VT_UI1, 1, bounds)
      raise ::RuntimeError, 'memory allocation error' if psa.nil? || psa.to_i.zero?

      data_out = ("\x00" * W::PTR_SIZE).b
      hr = safe_array_access_data.call(psa, data_out)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayAccessData: #{W.hr_hex(hr)}" if W.failed?(hr)

      Fiddle::Pointer.new(data_out.unpack1(W::PACK_PTR))[0, bytes.bytesize] = bytes
      safe_array_unaccess_data.call(psa)
      psa
    end

    def ui1_safearray_to_bytes(psa)
      data_out = ("\x00" * W::PTR_SIZE).b
      hr = safe_array_access_data.call(psa, data_out)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayAccessData: #{W.hr_hex(hr)}" if W.failed?(hr)

      lb_out = ("\x00" * 4).b
      ub_out = ("\x00" * 4).b
      safe_array_get_lbound.call(psa, 1, lb_out)
      safe_array_get_ubound.call(psa, 1, ub_out)
      len = ub_out.unpack1('l') - lb_out.unpack1('l') + 1
      bytes = Fiddle::Pointer.new(data_out.unpack1(W::PACK_PTR))[0, len].dup.b
      safe_array_unaccess_data.call(psa)
      bytes
    end
  end
end
