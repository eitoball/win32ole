# lib/win32ole/jruby/method.rb
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'
require 'win32ole/jruby/param'

class WIN32OLE
  class Method
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    FUNCFLAG_FRESTRICTED = 0x1
    FUNCFLAG_FHIDDEN = 0x40 # per MEMBERID/FUNCFLAGS, distinct from TYPEFLAG's own 0x10
    FUNCFLAG_FNONBROWSABLE = 0x400

    def initialize(itypeinfo_ptr, index)
      funcdesc_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.func_desc_fn(itypeinfo_ptr).call(itypeinfo_ptr, index, funcdesc_out)
      if W.failed?(hr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('GetFuncDesc', W.hr_hex(hr))
      end
      funcdesc_ptr = funcdesc_out.unpack1(W::PACK_PTR)
      funcdesc = TI::FUNCDESC.new(funcdesc_ptr)

      @memid = funcdesc.memid
      @invkind = funcdesc.invkind
      @dispid = funcdesc.memid
      @offset_vtbl = funcdesc.oVft
      @size_params = funcdesc.cParams
      @size_opt_params = funcdesc.cParamsOpt
      @func_flags = funcdesc.wFuncFlags
      @return_vt = funcdesc.ret_tdesc_vt

      @name, param_names = read_names(itypeinfo_ptr, @memid, funcdesc.cParams)
      @helpstring, @help_context, @helpfile = read_documentation(itypeinfo_ptr, @memid)

      # lprgelemdescParam is a POINTER FIELD inside FUNCDESC — its value is
      # the address of a contiguous array of cParams ELEMDESC structs, not
      # an offset into FUNCDESC itself.
      elemdesc_array_ptr = funcdesc.lprgelemdescParam
      @params = Array.new(funcdesc.cParams) do |i|
        elemdesc_ptr = elemdesc_array_ptr + i * TI::ELEMDESC.size
        WIN32OLE::Param.new(elemdesc_ptr, param_names[i])
      end

      TI.release_func_desc_fn(itypeinfo_ptr).call(itypeinfo_ptr, funcdesc_ptr)
    end

    def name
      @name
    end

    def return_type
      W.variant_ruby_type(@return_vt).to_s.upcase
    rescue NotImplementedError
      "VT_#{@return_vt}"
    end

    def return_vtype
      @return_vt
    end

    def return_type_detail
      [return_type]
    end

    def invoke_kind
      TI.invoke_kind_name(@invkind)
    end

    def invkind
      @invkind
    end

    def visible?
      (@func_flags & (FUNCFLAG_FRESTRICTED | FUNCFLAG_FHIDDEN | FUNCFLAG_FNONBROWSABLE)) == 0
    end

    def dispid
      @dispid
    end

    def helpstring
      @helpstring
    end

    def helpfile
      @helpfile
    end

    def helpcontext
      @help_context
    end

    def offset_vtbl
      @offset_vtbl
    end

    def size_params
      @size_params
    end

    def size_opt_params
      @size_opt_params
    end

    def params
      @params
    end

    def inspect
      "#<WIN32OLE::Method:#{name}>"
    end

    def event?
      raise NotImplementedError, 'ImplType traversal is not implemented yet (Phase 2 non-goal)'
    end

    def event_interface
      raise NotImplementedError, 'ImplType traversal is not implemented yet (Phase 2 non-goal)'
    end

    private

    def read_names(itypeinfo_ptr, memid, cparams)
      # GetNames(MEMBERID memid, BSTR *rgBstrNames, UINT cMaxNames, UINT *pcNames)
      # rgBstrNames[0] is the member's own name; rgBstrNames[1..] are param names.
      max_names = cparams + 1
      names_out = ("\x00" * (max_names * W::PTR_SIZE)).b
      count_out = ("\x00" * 4).b
      get_names_fn(itypeinfo_ptr).call(itypeinfo_ptr, memid, names_out, max_names, count_out)
      count = count_out.unpack1('L')
      bstrs = names_out.unpack(W::PACK_PTR * count)
      strings = bstrs.map { |b| W.bstr_to_s(b) }
      bstrs.each { |b| W.sys_free_string.call(b) unless b.zero? }
      [strings[0], strings[1..] || []]
    end

    def get_names_fn(itypeinfo_ptr)
      @@get_names_fns ||= {}
      @@get_names_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, TI::ITYPEINFO_VTBL[:GetNames],
        [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def read_documentation(itypeinfo_ptr, memid)
      name_out = ("\x00" * W::PTR_SIZE).b
      docstring_out = ("\x00" * W::PTR_SIZE).b
      helpcontext_out = ("\x00" * 4).b
      helpfile_out = ("\x00" * W::PTR_SIZE).b
      TI.documentation_fn_for_typeinfo(itypeinfo_ptr).call(
        itypeinfo_ptr, memid, name_out, docstring_out, helpcontext_out, helpfile_out
      )
      name_bstr = name_out.unpack1(W::PACK_PTR)
      docstring_bstr = docstring_out.unpack1(W::PACK_PTR)
      helpfile_bstr = helpfile_out.unpack1(W::PACK_PTR)
      helpstring = W.bstr_to_s(docstring_bstr)
      helpfile = W.bstr_to_s(helpfile_bstr)
      [name_bstr, docstring_bstr, helpfile_bstr].each { |b| W.sys_free_string.call(b) unless b.zero? }
      [helpstring, helpcontext_out.unpack1('L'), helpfile]
    end
  end
end
