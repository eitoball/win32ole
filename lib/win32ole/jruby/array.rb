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
  end
end
