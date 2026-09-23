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

    def initialize(elemdesc_ptr, name)
      @name = name
      elemdesc = TI::ELEMDESC.new(elemdesc_ptr)
      @vt = elemdesc.tdesc_vt
      @param_flags = elemdesc.wParamFlags
      @paramdescex_ptr = elemdesc.paramdescex_ptr
    end

    def name
      @name
    end

    def ole_type
      W.variant_ruby_type(@vt).to_s.upcase
    rescue NotImplementedError
      "VT_#{@vt}"
    end

    def ole_type_detail
      [ole_type]
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
      return nil unless (@param_flags & PARAMFLAG_FHASDEFAULT) != 0
      return nil if @paramdescex_ptr.nil? || @paramdescex_ptr.zero?

      # PARAMDESCEX is { ULONG cBytes; VARIANTARG varDefaultValue; } — the
      # VARIANTARG starts 4 bytes into the struct, right after cBytes.
      variant_bytes = Fiddle::Pointer.new(@paramdescex_ptr)[4, W::VARIANT_SIZE]
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

    def inspect
      "#<WIN32OLE::Param:#{name}=#{ole_type}>"
    end
  end
end
