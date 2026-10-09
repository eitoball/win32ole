# lib/win32ole/jruby/param.rb
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'

class WIN32OLE
  class Param
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    PARAMFLAG_FIN = 0x1
    PARAMFLAG_FOUT = 0x2
    PARAMFLAG_FRETVAL = 0x8
    PARAMFLAG_FOPT = 0x10
    PARAMFLAG_FHASDEFAULT = 0x20
    PARAMDESCEX_DEFAULT_VALUE_OFFSET = 8

    def initialize(itypeinfo_ptr, elemdesc_ptr, name)
      @name = name
      elemdesc = TI::ELEMDESC.new(elemdesc_ptr)
      @type_detail = TI.typedesc_details(
        itypeinfo_ptr, elemdesc_ptr.to_i + TI::ELEMDESC.offsetof('tdesc_union_ptr')
      )
      @param_flags = elemdesc.wParamFlags
      # Read the default value now, not lazily: PARAMDESCEX lives inside the
      # owning Method's FUNCDESC, which the Method releases (ReleaseFuncDesc)
      # right after all of its Params are constructed. Holding the raw
      # pointer past this point and dereferencing it later (e.g. from a
      # public #default call) would be a use-after-free.
      @default = read_default(elemdesc.paramdescex_ptr)
    end

    def name
      @name
    end
    alias to_s name

    def ole_type
      @type_detail.last
    end

    def ole_type_detail
      @type_detail.dup
    end

    def input?
      (@param_flags & PARAMFLAG_FIN) != 0
    end

    def output?
      (@param_flags & PARAMFLAG_FOUT) != 0
    end

    def optional?
      (@param_flags & PARAMFLAG_FOPT) != 0
    end

    def retval?
      (@param_flags & PARAMFLAG_FRETVAL) != 0
    end

    def default
      @default
    end

    # ext/win32ole/win32ole_param.c's foleparam_inspect: the default value,
    # not the type, follows the name -- and only when there is one.
    def inspect
      detail = @default.nil? ? name : "#{name}=#{@default.inspect}"
      "#<WIN32OLE::Param:#{detail}>"
    end

    private

    def read_default(paramdescex_ptr)
      return nil unless (@param_flags & PARAMFLAG_FHASDEFAULT) != 0
      return nil if paramdescex_ptr.to_i.zero?

      # PARAMDESCEX is { ULONG cBytes; VARIANTARG varDefaultValue; } -- a
      # VARIANT holds a DOUBLE/LONGLONG, so it is 8-byte aligned and starts
      # at offset 8 (cBytes plus 4 bytes of padding), not right after cBytes.
      variant_bytes = Fiddle::Pointer.new(paramdescex_ptr)[PARAMDESCEX_DEFAULT_VALUE_OFFSET, W::VARIANT_SIZE]
      vt, payload = W.unpack_variant(variant_bytes)
      case W.variant_ruby_type(vt)
      when :i4 then W.unpack_i4(payload)
      when :i8 then W.unpack_i8(payload)
      when :r8 then W.unpack_r8(payload)
      when :bool then W.unpack_bool(payload)
      when :bstr then W.bstr_to_s(W.unpack_pointer(payload))
      else nil
      end
    rescue NotImplementedError
      nil
    end
  end
end
