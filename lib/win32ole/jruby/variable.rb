# lib/win32ole/jruby/variable.rb
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'

class WIN32OLE
  class Variable
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    VARFLAG_FRESTRICTED = 0x80
    VARFLAG_FHIDDEN = 0x40
    VARFLAG_FNONBROWSABLE = 0x400
    VAR_CONST = 2

    def initialize(itypeinfo_ptr, index)
      vardesc_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.var_desc_fn(itypeinfo_ptr).call(itypeinfo_ptr, index, vardesc_out)
      if W.failed?(hr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('GetVarDesc', W.hr_hex(hr))
      end
      vardesc_ptr = vardesc_out.unpack1(W::PACK_PTR)
      vardesc = TI::VARDESC.new(vardesc_ptr)

      @memid = vardesc.memid
      @vt = vardesc.tdesc_vt
      @var_flags = vardesc.wVarFlags
      @varkind = vardesc.varkind
      @value = read_const_value(vardesc, vardesc_ptr) if @varkind == VAR_CONST
      @name = read_name(itypeinfo_ptr, @memid)

      TI.release_var_desc_fn(itypeinfo_ptr).call(itypeinfo_ptr, vardesc_ptr)
    end

    def name
      @name
    end

    def ole_type
      TI.vartype_name(@vt)
    end

    def ole_type_detail
      [ole_type]
    end

    def value
      @value
    end

    def visible?
      (@var_flags & (VARFLAG_FRESTRICTED | VARFLAG_FHIDDEN | VARFLAG_FNONBROWSABLE)) == 0
    end

    def variable_kind
      TI::VARKIND_NAMES.fetch(@varkind, 'UNKNOWN')
    end

    def varkind
      @varkind
    end

    def inspect
      "#<WIN32OLE::Variable:#{name}=#{value.inspect}>"
    end

    private

    def read_const_value(vardesc, vardesc_ptr)
      # VARDESC.union_oInst_or_lpvarValue holds a VARIANT* when varkind == VAR_CONST.
      variant_ptr = vardesc.union_oInst_or_lpvarValue
      return nil if variant_ptr.to_i.zero?

      variant_bytes = Fiddle::Pointer.new(variant_ptr)[0, W::VARIANT_SIZE]
      vt, payload = W.unpack_variant(variant_bytes)
      type = W.variant_ruby_type(vt)
      case type
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

    def read_name(itypeinfo_ptr, memid)
      names_out = ("\x00" * W::PTR_SIZE).b
      count_out = ("\x00" * 4).b
      get_names_fn(itypeinfo_ptr).call(itypeinfo_ptr, memid, names_out, 1, count_out)
      bstr = names_out.unpack1(W::PACK_PTR)
      name = W.bstr_to_s(bstr)
      W.sys_free_string.call(bstr) unless bstr.zero?
      name
    end

    def get_names_fn(itypeinfo_ptr)
      @@get_names_fns ||= {}
      @@get_names_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, TI::ITYPEINFO_VTBL[:GetNames],
        [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end
  end
end
