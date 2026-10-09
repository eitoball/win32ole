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

    # Mirrors ext/win32ole/win32ole_type.c's foletype_initialize:
    # WIN32OLE::Type.new(typelib, ole_class) resolves +typelib+ (a
    # registered typelib name or a CLSID string) to a file via
    # TypeLib.file_for, LoadTypeLibEx's it, then finds the ole_class
    # member by name.
    def self.new(typelib, ole_class)
      typelib = check_string!(typelib)
      ole_class = check_string!(ole_class)

      itypelib_ptr = load_typelib_for(typelib)
      begin
        found = find_type_by_name(itypelib_ptr, ole_class)
        raise WIN32OLE::RuntimeError, "not found `#{ole_class}` in `#{typelib}`" if found.nil?

        found
      ensure
        W.vtable_function(itypelib_ptr, 2, [W::VOIDP], W::DWORD).call(itypelib_ptr)
      end
    end

    def self.check_string!(val)
      return val if val.is_a?(::String)
      return val.to_str if val.respond_to?(:to_str)

      raise TypeError, "no implicit conversion of #{val.class} into String"
    end
    private_class_method :check_string!

    def self.load_typelib_for(typelib)
      file = WIN32OLE::TypeLib.file_for(typelib) || typelib
      out = ("\x00" * W::PTR_SIZE).b
      hr = TI.load_type_lib_ex.call(W.wstr(file), TI::REGKIND_NONE, out)
      raise WIN32OLE::RuntimeError, 'failed to LoadTypeLibEx' if W.failed?(hr)

      out.unpack1(W::PACK_PTR)
    end
    private_class_method :load_typelib_for

    # Mirrors oleclass_from_typelib: scan the typelib's members by
    # GetDocumentation name first (cheap), and only GetTypeInfo (which
    # AddRefs the result) once a name actually matches.
    def self.find_type_by_name(itypelib_ptr, ole_class)
      count = TI.type_info_count_fn(itypelib_ptr).call(itypelib_ptr)
      count.times do |i|
        name_out = ("\x00" * W::PTR_SIZE).b
        hr = TI.documentation_fn_for_typelib(itypelib_ptr).call(itypelib_ptr, i, name_out, nil, nil, nil)
        next if W.failed?(hr)

        name_bstr = name_out.unpack1(W::PACK_PTR)
        name = W.bstr_to_s(name_bstr)
        W.sys_free_string.call(name_bstr) unless name_bstr.zero?
        next unless name == ole_class

        ti_out = ("\x00" * W::PTR_SIZE).b
        hr = TI.type_info_fn(itypelib_ptr).call(itypelib_ptr, i, ti_out)
        return nil if W.failed?(hr)

        return from_typeinfo_ptr(ti_out.unpack1(W::PACK_PTR))
      end
      nil
    end
    private_class_method :find_type_by_name

    def self.from_typeinfo_ptr(itypeinfo_ptr)
      allocate.tap { |type| type.send(:initialize, itypeinfo_ptr) }
    end

    def initialize(itypeinfo_ptr)
      @ptr = itypeinfo_ptr
      # Install the finalizer before any call that could raise: the caller
      # has already AddRef'd this pointer, so if GetTypeAttr/read_documentation
      # (or the vtable lookups they perform) raise, this Release must still
      # happen -- otherwise the reference leaks permanently.
      install_finalizer

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
      # ole_type_src_type: the aliased type, resolved now (while TYPEATTR is
      # alive) through the raw TYPEDESC -- reading its union through the
      # struct's `void *` accessor overflows on the garbage upper bytes it
      # holds for anything but a real pointer.
      @src_type =
        if @typekind == TKIND_ALIAS
          TI.typedesc_details(@ptr, attr_ptr + TI::TYPEATTR.offsetof('tdescAlias_union_ptr')).last
        end
      TI.release_type_attr_fn(@ptr).call(@ptr, attr_ptr)

      @name, @helpstring, @help_context, @helpfile = read_documentation(@ptr, -1)
    end

    # WIN32OLE::Method.new(oletype, name) needs this raw ITypeInfo* to
    # search oletype's own funcs and its implemented interfaces -- same
    # deliberate, documented crack in encapsulation as WIN32OLE#dispatch_ptr.
    def itypeinfo_ptr
      @ptr
    end

    def name
      @name
    end
    alias to_s name

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
      @src_type
    end

    def variables
      count = type_attr_var_count
      Array.new(count) { |i| WIN32OLE::Variable.from_typeinfo_ptr(@ptr, i) }
    end

    # ole_methods_from_typeinfo: the type's own functions, then (one level)
    # those of every type it implements -- a coclass such as "Shell" has
    # none of its own.
    def ole_methods
      methods = Array.new(type_attr_func_count) { |i| WIN32OLE::Method.from_typeinfo_ptr(@ptr, i) }
      TI.each_impl_type(@ptr) do |_flags, ref_ptr|
        begin
          (TI.type_attr_value(ref_ptr, :cFuncs) || 0).times do |i|
            methods << WIN32OLE::Method.from_typeinfo_ptr(ref_ptr, i, @ptr)
          end
        ensure
          TI.release(ref_ptr)
        end
      end
      methods
    end

    def ole_typelib
      tlib_out = ("\x00" * W::PTR_SIZE).b
      index_out = ("\x00" * 4).b
      hr = TI.containing_typelib_fn(@ptr).call(@ptr, tlib_out, index_out)
      if W.failed?(hr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('GetContainingTypeLib', W.hr_hex(hr))
      end
      WIN32OLE::TypeLib.from_itypelib_ptr(tlib_out.unpack1(W::PACK_PTR))
    end

    def inspect
      "#<WIN32OLE::Type:#{name}>"
    end

    def implemented_ole_types
      impl_ole_types(0)
    end

    def source_ole_types
      impl_ole_types(TI::IMPLTYPEFLAG_FSOURCE)
    end

    def default_event_sources
      impl_ole_types(TI::IMPLTYPEFLAG_FSOURCE | TI::IMPLTYPEFLAG_FDEFAULT)
    end

    def default_ole_types
      impl_ole_types(TI::IMPLTYPEFLAG_FDEFAULT)
    end

    # foletype_s_ole_classes (obsolete in C too)
    def self.ole_classes(typelib)
      WIN32OLE::TypeLib.new(typelib).ole_types
    end

    # foletype_s_typelibs (obsolete in C too)
    def self.typelibs
      WIN32OLE::TypeLib.typelibs.map(&:name)
    end

    # foletype_s_progids: the ProgID and VersionIndependentProgID of every
    # HKCR\CLSID entry that has them.
    def self.progids
      tl = WIN32OLE::TypeLib
      progids = []
      tl.with_reg_subkey(TI::HKEY_CLASSES_ROOT, 'CLSID') do |hclsids|
        tl.each_reg_subkey_name(hclsids) do |clsid|
          tl.with_reg_subkey(hclsids, clsid) do |hclsid|
            %w[ProgID VersionIndependentProgID].each do |key|
              progid = tl.with_reg_subkey(hclsid, key) { |hkey| tl.reg_default_string_value(hkey) }
              progids << progid if progid
            end
          end
        end
      end
      progids
    end

    private

    # ole_type_impl_ole_types: implemented types whose IMPLTYPEFLAGs include
    # all of +implflags+. A matching ref ITypeInfo* is handed to the new
    # Type (whose finalizer releases it); the rest are released here.
    def impl_ole_types(implflags)
      types = []
      TI.each_impl_type(@ptr) do |flags, ref_ptr|
        if (flags & implflags) == implflags
          begin
            types << Type.from_typeinfo_ptr(ref_ptr)
          rescue WIN32OLE::QueryInterfaceError
            # like ole_type_from_itypeinfo returning nil: skip it (the
            # Type's finalizer, installed first, still releases ref_ptr)
          end
        else
          TI.release(ref_ptr)
        end
      end
      types
    end

    def type_attr_func_count
      attr_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.type_attr_fn(@ptr).call(@ptr, attr_out)
      if W.failed?(hr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('GetTypeAttr', W.hr_hex(hr))
      end
      attr_ptr = attr_out.unpack1(W::PACK_PTR)
      count = TI::TYPEATTR.new(attr_ptr).cFuncs
      TI.release_type_attr_fn(@ptr).call(@ptr, attr_ptr)
      count
    end

    def type_attr_var_count
      attr_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.type_attr_fn(@ptr).call(@ptr, attr_out)
      if W.failed?(hr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('GetTypeAttr', W.hr_hex(hr))
      end
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
      helpstring = W.bstr_to_s(docstring_bstr) || '' # C's WC2VSTR(NULL) is '', not nil
      helpfile = W.bstr_to_s(helpfile_bstr) || ''
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
