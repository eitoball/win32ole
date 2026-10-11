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

      return if sizes.any?(&:zero?)

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

      return if lbounds.zip(ubounds).any? { |lb, ub| ub < lb }

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

    POINTER_ELEMENT_TYPES = [W::VT_BSTR, W::VT_DISPATCH, W::VT_UNKNOWN].freeze

    # One element's value, read the way ole_variant2val reads a SAFEARRAY:
    # through a VT_BYREF VARIANT pointing at it, without taking ownership
    # of anything the array still holds.
    def read_element(elem_addr, base_vt)
      byref = W.pack_variant(base_vt | W::VT_BYREF, W.pack_pointer(elem_addr))
      WIN32OLE.variant_bytes_to_ruby_value(byref, owned: false)
    end

    # ole_set_safe_array's val2variant_ptr + SafeArrayPutElement: the
    # value's natural VARIANT, coerced to the element type; a VT_VARIANT
    # element takes the VARIANT itself, a BSTR/IDispatch/IUnknown element
    # the pointer (SafeArrayPutElement copies or AddRefs it), any other the
    # address of the value.
    def write_element(psa, index_buf, base_vt, val, to_free)
      if base_vt == W::VT_VARIANT
        var = WIN32OLE.ruby_value_to_variant_bytes(val, to_free)
        return put_element(psa, index_buf, W.native_pointer_for(var))
      end

      var = val.nil? ? W.pack_variant(base_vt, W.pack_empty) : WIN32OLE.ruby_value_to_variant_bytes(val, to_free)
      natural_vt, = W.unpack_variant(var)
      if natural_vt != base_vt
        var = W.change_variant_type(var, base_vt)
        converted_bstr = base_vt == W::VT_BSTR
      end
      _vt, payload = W.unpack_variant(var)
      if POINTER_ELEMENT_TYPES.include?(base_vt)
        ptr = W.unpack_pointer(payload)
        if ptr.zero? && base_vt != W::VT_BSTR
          raise WIN32OLE::RuntimeError, 'element of array does not have IDispatch or IUnknown Interface'
        end

        put_element(psa, index_buf, ptr)
        W.sys_free_string.call(ptr) if converted_bstr && !ptr.zero?
      else
        put_element(psa, index_buf, W.native_pointer_for(payload))
      end
    end

    def put_element(psa, index_buf, pv)
      hr = safe_array_put_element.call(psa, index_buf, pv)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayPutElement: #{W.hr_hex(hr)}" if W.failed?(hr)
    end

    def ruby_array_to_safearray(ary, elem_vt, to_free = [])
      base_vt = elem_vt & W::VT_TYPEMASK
      base_vt = W::VT_VARIANT if base_vt.zero? # ole_val_ary2variant_ary: plain VT_ARRAY means VT_VARIANT elements
      return ui1_safearray_from_bytes(ary) if base_vt == W::VT_UI1 && ary.is_a?(::String)
      raise TypeError, "wrong argument type #{ary.class} (expected Array)" unless ary.is_a?(::Array)

      sizes = dimension_sizes(ary)
      dims = sizes.size
      bounds = sizes.flat_map { |n| [n, 0] }.pack('L2' * dims)
      psa = safe_array_create.call(base_vt, dims, bounds)
      raise ::RuntimeError, 'memory allocation error' if psa.nil? || psa.to_i.zero?

      hr = safe_array_lock.call(psa)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayLock: #{W.hr_hex(hr)}" if W.failed?(hr)

      begin
        each_fill_index(sizes) do |pid|
          write_element(psa, pid.pack('l' * dims), base_vt, nested_entry(ary, pid), to_free)
        end
      ensure
        safe_array_unlock.call(psa)
      end
      psa
    end

    def safearray_to_ruby_array(psa, elem_vt)
      base_vt = elem_vt & W::VT_TYPEMASK

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

          val = read_element(elem_ptr_out.unpack1(W::PACK_PTR), base_vt)
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
