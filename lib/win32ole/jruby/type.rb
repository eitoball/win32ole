# lib/win32ole/jruby/type.rb
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'
require 'win32ole/jruby/method'
require 'win32ole/jruby/variable'

class WIN32OLE
  class Type
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    TYPEFLAG_FHIDDEN = 0x10
    TYPEFLAG_FRESTRICTED = 0x200
    TKIND_ALIAS = 6
    VT_USERDEFINED = 29

    def initialize(itypeinfo_ptr)
      @ptr = itypeinfo_ptr

      attr_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.type_attr_fn(@ptr).call(@ptr, attr_out)
      if W.failed?(hr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('GetTypeAttr', W.hr_hex(hr))
      end
      attr_ptr = attr_out.unpack1(W::PACK_PTR)
      typeattr = TI::TYPEATTR.new(attr_ptr)

      @guid_bytes = [typeattr.guid_Data1, typeattr.guid_Data2, typeattr.guid_Data3].pack('LSS') +
                    typeattr.guid_Data4.pack('C8')
      @typekind = typeattr.typekind
      @major = typeattr.wMajorVerNum
      @minor = typeattr.wMinorVerNum
      @type_flags = typeattr.wTypeFlags
      @alias_vt = typeattr.tdescAlias_vt
      # tdescAlias_union_ptr is a `void *` struct field, so Fiddle hands back
      # a Fiddle::Pointer here (never a plain integer). GetRefTypeInfo's
      # hRefType parameter is a DWORD, not a pointer, so it must be
      # converted to an integer before being passed to that call (see
      # src_type below) — a Fiddle::Pointer object is not a valid argument
      # for a DWORD-typed native parameter.
      @alias_union_ptr = typeattr.tdescAlias_union_ptr.to_i
      TI.release_type_attr_fn(@ptr).call(@ptr, attr_ptr)

      @name, @helpstring, @help_context, @helpfile = read_documentation(@ptr, -1)

      install_finalizer
    end

    def name
      @name
    end

    def ole_type
      TI::TYPEKIND_NAMES[@typekind]
    end

    def guid
      d1, d2, d3 = @guid_bytes.unpack('LSS')
      d4 = @guid_bytes[8, 8].unpack('C8')
      format('{%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}', d1, d2, d3, *d4)
    end

    def progid
      W.prog_id_from_clsid(@guid_bytes)
    end

    def visible?
      (@type_flags & (TYPEFLAG_FHIDDEN | TYPEFLAG_FRESTRICTED)) == 0
    end

    def major_version
      @major
    end

    def minor_version
      @minor
    end

    def typekind
      @typekind
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

    def src_type
      return nil unless @typekind == TKIND_ALIAS && @alias_vt == VT_USERDEFINED

      ref_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.ref_type_info_fn(@ptr).call(@ptr, @alias_union_ptr, ref_out)
      return nil if W.failed?(hr)

      Type.new(ref_out.unpack1(W::PACK_PTR)).name
    end

    def variables
      count = type_attr_var_count
      Array.new(count) { |i| WIN32OLE::Variable.new(@ptr, i) }
    end

    def ole_methods
      count = type_attr_func_count
      Array.new(count) { |i| WIN32OLE::Method.new(@ptr, i) }
    end

    def ole_typelib
      tlib_out = ("\x00" * W::PTR_SIZE).b
      index_out = ("\x00" * 4).b
      hr = TI.containing_typelib_fn(@ptr).call(@ptr, tlib_out, index_out)
      if W.failed?(hr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('GetContainingTypeLib', W.hr_hex(hr))
      end
      WIN32OLE::TypeLib.new(tlib_out.unpack1(W::PACK_PTR))
    end

    def inspect
      "#<WIN32OLE::Type:#{name}>"
    end

    private

    def type_attr_func_count
      attr_out = ("\x00" * W::PTR_SIZE).b
      TI.type_attr_fn(@ptr).call(@ptr, attr_out)
      attr_ptr = attr_out.unpack1(W::PACK_PTR)
      count = TI::TYPEATTR.new(attr_ptr).cFuncs
      TI.release_type_attr_fn(@ptr).call(@ptr, attr_ptr)
      count
    end

    def type_attr_var_count
      attr_out = ("\x00" * W::PTR_SIZE).b
      TI.type_attr_fn(@ptr).call(@ptr, attr_out)
      attr_ptr = attr_out.unpack1(W::PACK_PTR)
      count = TI::TYPEATTR.new(attr_ptr).cVars
      TI.release_type_attr_fn(@ptr).call(@ptr, attr_ptr)
      count
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
      name = W.bstr_to_s(name_bstr)
      helpstring = W.bstr_to_s(docstring_bstr)
      helpfile = W.bstr_to_s(helpfile_bstr)
      [name_bstr, docstring_bstr, helpfile_bstr].each { |b| W.sys_free_string.call(b) unless b.zero? }
      [name, helpstring, helpcontext_out.unpack1('L'), helpfile]
    end

    def install_finalizer
      ptr = @ptr
      release_fn = W.vtable_function(ptr, 2, [W::VOIDP], W::DWORD)
      ObjectSpace.define_finalizer(self, self.class.finalizer(ptr, release_fn))
    end

    def self.finalizer(ptr, release_fn)
      proc { release_fn.call(ptr) unless ptr.zero? }
    end
  end
end
