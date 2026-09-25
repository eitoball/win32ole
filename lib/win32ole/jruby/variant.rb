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
        @var = WIN32OLE.ruby_value_to_variant_bytes(val, @bstrs_to_free = [])
        return
      end

      raise ArgumentError, 'WIN32OLE::Variant does not support VT_RECORD; use WIN32OLE::Record instead' if (vartype & VT::VT_TYPEMASK) == VT::VT_RECORD

      base_vt = vartype & VT::VT_TYPEMASK
      byref = (vartype & VT::VT_BYREF) != 0

      @realvar =
        if (vartype & VT::VT_ARRAY) != 0
          @bstrs_to_free = []
          psa = SA.ruby_array_to_safearray(val, base_vt, @bstrs_to_free)
          W.pack_variant(vartype & ~VT::VT_BYREF, W.pack_pointer(psa.to_i))
        else
          @bstrs_to_free = []
          pack_scalar_explicit(base_vt, val)
        end

      @var = byref ? W.pack_byref(vartype & ~VT::VT_BYREF, @realvar) : @realvar
    end

    def value
      vt, = W.unpack_variant(@var)
      base_vt = vt & VT::VT_TYPEMASK
      return SA.safearray_to_ruby_array(current_array_ptr, base_vt) if (vt & VT::VT_ARRAY) != 0

      WIN32OLE.variant_bytes_to_ruby_value(current_scalar_bytes)
    end

    def value=(val)
      vt, = W.unpack_variant(@var)
      base_vt = vt & VT::VT_TYPEMASK
      if (vt & VT::VT_ARRAY) != 0
        unless val.is_a?(::String) && base_vt == VT::VT_UI1
          raise WIN32OLE::RuntimeError, 'array value can only be replaced with a String for a VT_UI1|VT_ARRAY Variant'
        end

        psa = SA.ruby_array_to_safearray(val, base_vt, @bstrs_to_free ||= [])
        @realvar = W.pack_variant(vt & ~VT::VT_BYREF, W.pack_pointer(psa.to_i))
      else
        @realvar = pack_scalar_explicit(base_vt, val)
      end
      @var = (vt & VT::VT_BYREF) != 0 ? W.pack_byref(vt & ~VT::VT_BYREF, @realvar) : @realvar
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
      hr = SA.safe_array_lock.call(psa)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayLock: #{W.hr_hex(hr)}" if W.failed?(hr)

      begin
        index_buf = indices.pack('l' * indices.size)
        elem_ptr_out = ("\x00" * W::PTR_SIZE).b
        hr = SA.safe_array_ptr_of_index.call(psa, index_buf, elem_ptr_out)
        raise WIN32OLE::RuntimeError, "failed to SafeArrayPtrOfIndex: #{W.hr_hex(hr)}" if W.failed?(hr)

        elem_addr = elem_ptr_out.unpack1(W::PACK_PTR)
        if base_vt == VT::VT_VARIANT
          WIN32OLE.variant_bytes_to_ruby_value(Fiddle::Pointer.new(elem_addr)[0, W::VARIANT_SIZE])
        else
          fmt = SA::ELEMENT_PACK_FORMAT.fetch(base_vt) { raise NotImplementedError, "VARTYPE #{base_vt} is not a supported array element type yet" }
          SA.unpack_scalar_element(base_vt, Fiddle::Pointer.new(elem_addr)[0, [1].pack(fmt).bytesize])
        end
      ensure
        SA.safe_array_unlock.call(psa)
      end
    end

    def []=(*args)
      val = args.pop
      indices = args
      base_vt, psa = array_state
      hr = SA.safe_array_lock.call(psa)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayLock: #{W.hr_hex(hr)}" if W.failed?(hr)

      begin
        leaf = base_vt == VT::VT_VARIANT ? WIN32OLE.ruby_value_to_variant_bytes(val, @bstrs_to_free ||= [])
                                          : SA.pack_scalar_element(base_vt, val)
        index_buf = indices.pack('l' * indices.size)
        hr = SA.safe_array_put_element.call(psa, index_buf, W.native_pointer_for(leaf))
        raise WIN32OLE::RuntimeError, "failed to SafeArrayPutElement: #{W.hr_hex(hr)}" if W.failed?(hr)

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

    def current_scalar_bytes
      vt, payload = W.unpack_variant(@var)
      return @var if (vt & VT::VT_BYREF).zero?

      base_vt = vt & VT::VT_TYPEMASK
      addr = W.unpack_pointer(payload)
      if base_vt == VT::VT_VARIANT
        Fiddle::Pointer.new(addr)[0, W::VARIANT_SIZE]
      else
        body = Fiddle::Pointer.new(addr)[0, 8]
        W.pack_variant(base_vt, body)
      end
    end

    SCALAR_PACK = {
      VT::VT_I1 => :pack_i1, VT::VT_UI1 => :pack_ui1, VT::VT_I2 => :pack_i2, VT::VT_UI2 => :pack_ui2,
      VT::VT_I4 => :pack_i4, VT::VT_UI4 => :pack_ui4, VT::VT_INT => :pack_int, VT::VT_UINT => :pack_uint,
      VT::VT_I8 => :pack_i8, VT::VT_UI8 => :pack_ui8, VT::VT_R4 => :pack_r4, VT::VT_R8 => :pack_r8,
      VT::VT_BOOL => :pack_bool, VT::VT_ERROR => :pack_error, VT::VT_EMPTY => :pack_empty
    }.freeze

    def pack_scalar_explicit(base_vt, val)
      if base_vt == VT::VT_EMPTY || base_vt == VT::VT_NULL
        return W.pack_variant(base_vt, W.pack_empty)
      end
      if base_vt == VT::VT_BSTR
        bstr = W.sys_alloc_string.call(W.wstr(val))
        (@bstrs_to_free ||= []) << bstr
        return W.pack_variant(base_vt, W.pack_pointer(bstr))
      end
      if base_vt == VT::VT_DISPATCH || base_vt == VT::VT_UNKNOWN
        ptr = val.nil? ? 0 : val.instance_variable_get(:@ptr)
        return W.pack_variant(base_vt, W.pack_pointer(ptr))
      end

      pack_method = SCALAR_PACK[base_vt]
      unless pack_method
        raise NotImplementedError, "VARTYPE #{base_vt} is not supported yet"
      end

      inferred_vt = W.ruby_to_variant_type(val) rescue nil
      needs_coercion = inferred_vt && W::VT_FOR_TYPE[inferred_vt] != base_vt
      if needs_coercion
        coerce_via_variant_change_type(val, base_vt)
      else
        W.pack_variant(base_vt, W.send(pack_method, val))
      end
    end

    # Mirrors MRI's ole_val2variant_ex + VariantChangeTypeEx fallback for
    # a mismatched explicit VARTYPE (spec §4.2) -- e.g.
    # Variant.new("2e3", VariantType::VT_R4). Native call, CI-only.
    def coerce_via_variant_change_type(val, base_vt)
      src_type = W.ruby_to_variant_type(val)
      src = W.pack_variant(W::VT_FOR_TYPE.fetch(src_type), WIN32OLE.ruby_value_to_variant_bytes(val, @bstrs_to_free ||= [])[8, 8])
      dest = ("\x00" * W::VARIANT_SIZE).b
      hr = W.variant_change_type.call(dest, src, W::LOCALE_SYSTEM_DEFAULT, 0, base_vt)
      raise WIN32OLE::RuntimeError, "failed to change variant type: #{W.hr_hex(hr)}" if W.failed?(hr)

      dest
    end

    def set_array_var(psa, vt)
      vt |= VT::VT_ARRAY
      @realvar = W.pack_variant(vt & ~VT::VT_BYREF, W.pack_pointer(psa.to_i))
      @var = (vt & VT::VT_BYREF) != 0 ? W.pack_byref(vt & ~VT::VT_BYREF, @realvar) : @realvar
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
