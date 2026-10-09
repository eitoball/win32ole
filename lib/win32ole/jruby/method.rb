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

    # Mirrors ext/win32ole/win32ole_method.c's folemethod_initialize:
    # WIN32OLE::Method.new(oletype, method_name) searches oletype's own
    # funcs first, then (since a coclass like "Shell" has none of its own
    # -- its methods live on the interfaces it implements) one level of
    # its implemented types, case-insensitively by name (COM method names
    # are case-insensitive).
    def self.new(oletype, method)
      raise TypeError, '1st argument should be WIN32OLE::Type object' unless oletype.is_a?(WIN32OLE::Type)

      method = check_string!(method)
      owner_ptr = oletype.itypeinfo_ptr
      index = func_index_in(owner_ptr, method)
      return from_typeinfo_ptr(owner_ptr, index) if index

      # ole_method_sub(self, pTypeInfo, pRefTypeInfo, ...): a match on an
      # implemented type records oletype as the owner (for #event?).
      # initialize only reads ref_ptr synchronously, so it is released here
      # either way.
      TI.each_impl_type(owner_ptr) do |_flags, ref_ptr|
        begin
          ref_index = func_index_in(ref_ptr, method)
          return from_typeinfo_ptr(ref_ptr, ref_index, owner_ptr) if ref_index
        ensure
          TI.release(ref_ptr)
        end
      end
      raise WIN32OLE::RuntimeError, "not found #{method}"
    end

    def self.check_string!(val)
      return val if val.is_a?(::String)
      return val.to_str if val.respond_to?(:to_str)

      raise TypeError, "no implicit conversion of #{val.class} into String"
    end
    private_class_method :check_string!

    # Index of the FUNCDESC in itypeinfo_ptr whose name matches
    # case-insensitively, or nil.
    def self.func_index_in(itypeinfo_ptr, name)
      (TI.type_attr_value(itypeinfo_ptr, :cFuncs) || 0).times do |i|
        funcdesc_out = ("\x00" * W::PTR_SIZE).b
        next if W.failed?(TI.func_desc_fn(itypeinfo_ptr).call(itypeinfo_ptr, i, funcdesc_out))

        funcdesc_ptr = funcdesc_out.unpack1(W::PACK_PTR)
        memid = TI::FUNCDESC.new(funcdesc_ptr).memid
        TI.release_func_desc_fn(itypeinfo_ptr).call(itypeinfo_ptr, funcdesc_ptr)

        fname = TI.member_name(itypeinfo_ptr, memid)
        return i if fname && fname.casecmp?(name)
      end
      nil
    end
    private_class_method :func_index_in

    # owner_ptr is the coclass whose implemented type itypeinfo_ptr is
    # (olemethoddata's pOwnerTypeInfo), or nil for a type's own function.
    def self.from_typeinfo_ptr(itypeinfo_ptr, index, owner_ptr = nil)
      allocate.tap { |method| method.send(:initialize, itypeinfo_ptr, index, owner_ptr) }
    end

    def initialize(itypeinfo_ptr, index, owner_ptr = nil)
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
      @return_type_detail = TI.typedesc_details(
        itypeinfo_ptr, funcdesc_ptr + TI::FUNCDESC.offsetof('ret_tdesc_union_ptr')
      )

      @name, param_names = read_names(itypeinfo_ptr, @memid, funcdesc.cParams)
      # folemethod_event / folemethod_event_interface, computed now because
      # neither ITypeInfo* is retained past construction.
      @event = owner_ptr ? TI.method_event?(owner_ptr, index, @name) : false
      @event_interface = @event ? TI.type_name(itypeinfo_ptr) : nil
      @helpstring, @help_context, @helpfile = read_documentation(itypeinfo_ptr, @memid)

      # lprgelemdescParam is a POINTER FIELD inside FUNCDESC — its value is
      # the address of a contiguous array of cParams ELEMDESC structs, not
      # an offset into FUNCDESC itself.
      elemdesc_array_ptr = funcdesc.lprgelemdescParam
      # MRI builds one Param per name GetNames actually returned beyond the
      # member's own name (param_names.length), not per FUNCDESC.cParams —
      # for a property-put method, GetNames returns only the property name
      # (cParams=1 but zero param names), so cParams would otherwise build a
      # spurious nameless Param.
      @params = Array.new(param_names.length) do |i|
        elemdesc_ptr = elemdesc_array_ptr + i * TI::ELEMDESC.size
        WIN32OLE::Param.from_elemdesc(itypeinfo_ptr, elemdesc_ptr, param_names[i])
      end

      TI.release_func_desc_fn(itypeinfo_ptr).call(itypeinfo_ptr, funcdesc_ptr)
    end

    def name
      @name
    end
    alias to_s name

    def return_type
      @return_type_detail.last
    end

    def return_vtype
      @return_vt
    end

    def return_type_detail
      @return_type_detail.dup
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
      @event
    end

    def event_interface
      @event_interface
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
      @@get_names_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
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
      helpstring = W.bstr_to_s(docstring_bstr) || '' # C's WC2VSTR(NULL) is '', not nil
      helpfile = W.bstr_to_s(helpfile_bstr) || ''
      [name_bstr, docstring_bstr, helpfile_bstr].each { |b| W.sys_free_string.call(b) unless b.zero? }
      [helpstring, helpcontext_out.unpack1('L'), helpfile]
    end
  end
end
