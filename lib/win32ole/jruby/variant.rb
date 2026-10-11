require 'fiddle'
require 'win32ole/jruby/win32'
require 'win32ole/jruby/array'

class WIN32OLE
  module VariantType
    VT_EMPTY = 0
    VT_NULL = 1
    VT_I2 = 2
    VT_I4 = 3
    VT_R4 = 4
    VT_R8 = 5
    VT_CY = 6
    VT_DATE = 7
    VT_BSTR = 8
    VT_DISPATCH = 9
    VT_ERROR = 10
    VT_BOOL = 11
    VT_VARIANT = 12
    VT_UNKNOWN = 13
    VT_DECIMAL = 14
    VT_I1 = 16
    VT_UI1 = 17
    VT_UI2 = 18
    VT_UI4 = 19
    VT_I8 = 20
    VT_UI8 = 21
    VT_INT = 22
    VT_UINT = 23
    VT_VOID = 24
    VT_HRESULT = 25
    VT_PTR = 26
    VT_SAFEARRAY = 27
    VT_CARRAY = 28
    VT_USERDEFINED = 29
    VT_LPSTR = 30
    VT_LPWSTR = 31
    VT_RECORD = 36
    VT_INT_PTR = 37
    VT_UINT_PTR = 38
    VT_FILETIME = 64
    VT_BLOB = 65
    VT_STREAM = 66
    VT_STORAGE = 67
    VT_STREAMED_OBJECT = 68
    VT_STORED_OBJECT = 69
    VT_BLOB_OBJECT = 70
    VT_CF = 71
    VT_CLSID = 72
    VT_VERSIONED_STREAM = 73
    VT_BSTR_BLOB = 0xFFF
    VT_VECTOR = 0x1000
    VT_ARRAY = 0x2000
    VT_BYREF = 0x4000
    VT_RESERVED = 0x8000
    VT_ILLEGAL = 0xFFFF
    VT_ILLEGALMASKED = 0xFFF
    VT_TYPEMASK = 0xFFF
  end

  VARIANT = VariantType

  class Variant
    W = Win32
    SA = SafeArray
    VT = VariantType
    private_constant :W, :SA, :VT

    def initialize(val, vartype = nil, _reserved = nil)
      if vartype.nil?
        @var = WIN32OLE.ruby_value_to_variant_bytes(val, @to_free = [])
        return
      end

      raise ArgumentError, 'WIN32OLE::Variant does not support VT_RECORD; use WIN32OLE::Record instead' if (vartype & VT::VT_TYPEMASK) == VT::VT_RECORD

      # ole_val_ary2variant_ary: a plain VT_ARRAY is an array of VARIANTs.
      vartype |= VT::VT_VARIANT if (vartype & ~VT::VT_BYREF) == VT::VT_ARRAY
      base_vt = vartype & VT::VT_TYPEMASK
      byref = (vartype & VT::VT_BYREF) != 0

      realvar_bytes =
        if (vartype & VT::VT_ARRAY) != 0
          @to_free = []
          # nil: the type with no SAFEARRAY behind it, as C leaves it
          psa = val.nil? ? 0 : SA.ruby_array_to_safearray(val, base_vt, @to_free).to_i
          W.pack_variant(vartype & ~VT::VT_BYREF, W.pack_pointer(psa))
        else
          @to_free = []
          scalar_realvar(val, vartype)
        end

      @realvar = byref ? persist_realvar(realvar_bytes) : realvar_bytes
      @var = byref ? W.pack_byref(vartype & ~VT::VT_BYREF, @realvar) : @realvar
      commit_realvar(@realvar) if byref
    end

    def value
      vt, = W.unpack_variant(@var)
      base_vt = vt & VT::VT_TYPEMASK
      if (vt & VT::VT_ARRAY) != 0 && base_vt == VT::VT_UI1
        psa = current_array_ptr
        return SA.ui1_safearray_to_bytes(psa) if !psa.zero? && SA.safe_array_get_dim.call(psa) == 1
      end

      # The Variant keeps owning what @var holds (folevariant_value copies).
      WIN32OLE.variant_bytes_to_ruby_value(@var, owned: false)
    end

    def value=(val)
      vt, = W.unpack_variant(@var)
      base_vt = vt & VT::VT_TYPEMASK
      byref = (vt & VT::VT_BYREF) != 0
      realvar_bytes =
        if (vt & VT::VT_ARRAY) != 0
          unless val.is_a?(::String) && base_vt == VT::VT_UI1
            raise WIN32OLE::RuntimeError, 'array value can only be replaced with a String for a VT_UI1|VT_ARRAY Variant'
          end

          psa = SA.ruby_array_to_safearray(val, base_vt, @to_free ||= [])
          W.pack_variant(vt & ~VT::VT_BYREF, W.pack_pointer(psa.to_i))
        else
          scalar_realvar(val, vt)
        end
      @realvar = byref ? persist_realvar(realvar_bytes) : realvar_bytes
      @var = byref ? W.pack_byref(vt & ~VT::VT_BYREF, @realvar) : @realvar
      commit_realvar(@realvar) if byref
    end

    def vartype
      vt, = W.unpack_variant(@var)
      vt
    end

    def self.array(dims, vt)
      raise TypeError, "wrong argument type #{dims.class} (expected Array)" unless dims.is_a?(::Array)

      bounds = dims.flat_map { |n| [n, 0] }.pack('L2' * dims.size)
      psa = SA.safe_array_create.call(vt & VT::VT_TYPEMASK, dims.size, bounds)
      raise ::RuntimeError, 'memory allocation error' if psa.nil? || psa.to_i.zero?

      allocate.tap { |v| v.send(:set_array_var, psa, vt) }
    end

    def [](*indices)
      base_vt, psa = array_state
      dim = SA.safe_array_get_dim.call(psa)
      raise ArgumentError, 'unmatch number of indices' unless indices.size == dim

      hr = SA.safe_array_lock.call(psa)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayLock: #{W.hr_hex(hr)}" if W.failed?(hr)

      begin
        index_buf = indices.pack('l' * indices.size)
        elem_ptr_out = ("\x00" * W::PTR_SIZE).b
        hr = SA.safe_array_ptr_of_index.call(psa, index_buf, elem_ptr_out)
        raise WIN32OLE::RuntimeError, "failed to SafeArrayPtrOfIndex: #{W.hr_hex(hr)}" if W.failed?(hr)

        SA.read_element(elem_ptr_out.unpack1(W::PACK_PTR), base_vt)
      ensure
        SA.safe_array_unlock.call(psa)
      end
    end

    def []=(*args)
      val = args.pop
      indices = args
      base_vt, psa = array_state
      dim = SA.safe_array_get_dim.call(psa)
      raise ArgumentError, 'unmatch number of indices' unless indices.size == dim

      hr = SA.safe_array_lock.call(psa)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayLock: #{W.hr_hex(hr)}" if W.failed?(hr)

      begin
        SA.write_element(psa, indices.pack('l' * indices.size), base_vt, val, @to_free ||= [])
        val
      ensure
        SA.safe_array_unlock.call(psa)
      end
    end

    private

    def current_array_ptr
      vt, payload = W.unpack_variant(@var)
      addr = W.unpack_pointer(payload)
      return addr if (vt & VT::VT_BYREF).zero?

      Fiddle::Pointer.new(addr)[0, W::PTR_SIZE].unpack1(W::PACK_PTR)
    end

    # ext/win32ole/win32ole_variant.c's ole_val2olevariantdata, non-array
    # branch: nil is the zero of the requested type; VT_ERROR takes the
    # SCODE as is; anything else is converted to its natural VARIANT (I8/UI8
    # packed directly, as ole_val2variant_ex does) and, when that type
    # differs, coerced with VariantChangeTypeEx. VT_VARIANT keeps the
    # natural VARIANT.
    def scalar_realvar(val, vt)
      base_vt = vt & ~VT::VT_BYREF
      return error_realvar(val) if base_vt == VT::VT_ERROR
      if val.nil?
        return W.pack_variant(VT::VT_EMPTY, W.pack_empty) if vt == (VT::VT_VARIANT | VT::VT_BYREF)

        return W.pack_variant(base_vt, W.pack_empty)
      end

      natural =
        case base_vt
        when VT::VT_I8 then W.pack_variant(base_vt, W.pack_i8(val))
        when VT::VT_UI8 then W.pack_variant(base_vt, W.pack_ui8(val))
        else WIN32OLE.ruby_value_to_variant_bytes(val, @to_free ||= [])
        end
      return natural if base_vt == VT::VT_VARIANT

      natural_vt, = W.unpack_variant(natural)
      natural_vt == base_vt ? natural : W.change_variant_type(natural, base_vt)
    end

    # ole_val2variant_err: an Integer SCODE (or nil for 0), possibly held
    # in another Variant.
    def error_realvar(val)
      scode = val.is_a?(Variant) ? val.value : val
      unless scode.nil? || scode.is_a?(::Integer)
        raise WIN32OLE::RuntimeError, "failed to convert VT_ERROR VARIANT:`#{scode.inspect}'"
      end

      W.pack_variant(VT::VT_ERROR, W.pack_error(scode || 0))
    end

    def set_array_var(psa, vt)
      vt |= VT::VT_ARRAY
      realvar_bytes = W.pack_variant(vt & ~VT::VT_BYREF, W.pack_pointer(psa.to_i))
      byref = (vt & VT::VT_BYREF) != 0
      @realvar = byref ? persist_realvar(realvar_bytes) : realvar_bytes
      @var = byref ? W.pack_byref(vt & ~VT::VT_BYREF, @realvar) : @realvar
      commit_realvar(@realvar) if byref
    end

    # Allocates a persistent buffer for bytes and returns it -- does NOT
    # touch @realvar_finalizer_state yet. Call commit_realvar with the
    # result once @var/@realvar have been rebuilt to reference it, so the
    # old buffer (if any) is only freed after nothing references it
    # anymore -- freeing it first (the naive order) leaves a window where
    # @var's own embedded bytes point at already-freed memory, and if
    # allocation itself raises, freeing the old buffer up front would
    # leave the finalizer holding an already-freed pointer (double-free
    # at GC time).
    def persist_realvar(bytes)
      @realvar_finalizer_state ||= {}.tap do |state|
        ObjectSpace.define_finalizer(self, self.class.realvar_finalizer(state))
      end
      W.persistent_pointer_for(bytes)
    end

    # Stores ptr as the current realvar buffer and frees whatever was
    # there before -- call this only after @var/@realvar already
    # reference ptr, not before.
    def commit_realvar(ptr)
      old_ptr = @realvar_finalizer_state[:ptr]
      @realvar_finalizer_state[:ptr] = ptr
      Fiddle.free(old_ptr) if old_ptr
    end

    def self.realvar_finalizer(state)
      proc { Fiddle.free(state[:ptr]) if state[:ptr] }
    end

    def array_state
      vt, payload = W.unpack_variant(@var)
      addr = W.unpack_pointer(payload)
      addr = Fiddle::Pointer.new(addr)[0, W::PTR_SIZE].unpack1(W::PACK_PTR) if (vt & VT::VT_BYREF) != 0
      [vt & VT::VT_TYPEMASK, addr]
    end

    DISP_E_PARAMNOTFOUND = -2147352572 # 0x80020004

    # Empty/Null/Nothing/NoParam are lazy, not eager: NoParam's construction
    # calls VariantChangeTypeEx (a real Fiddle.dlopen('oleaut32') on first
    # use), and every other native call in this codebase is lazy/memoized so
    # that simply requiring the gem is harmless on any platform. const_missing
    # preserves the exact same WIN32OLE::Variant::Empty-style constant-access
    # API MRI provides, while deferring construction to first actual access.
    def self.const_missing(name)
      case name
      when :Empty then const_set(:Empty, new(nil, VariantType::VT_EMPTY))
      when :Null then const_set(:Null, new(nil, VariantType::VT_NULL))
      when :Nothing then const_set(:Nothing, new(nil, VariantType::VT_DISPATCH))
      when :NoParam then const_set(:NoParam, new(DISP_E_PARAMNOTFOUND, VariantType::VT_ERROR))
      else super
      end
    end
  end
end
