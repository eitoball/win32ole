# lib/win32ole/jruby/typeinfo.rb
require 'fiddle'
require 'fiddle/import'
require 'win32ole/jruby/win32'

class WIN32OLE
  module TypeInfo
    extend Fiddle::Importer

    GUID = struct([
      'unsigned int Data1', 'unsigned short Data2', 'unsigned short Data3',
      'unsigned char Data4[8]'
    ])

    TYPEDESC = struct([
      'void *union_ptr',  # lptdesc / lpadesc / hreftype share this slot;
                          # caller's context determines which interpretation applies
      'unsigned short vt'
    ])

    ELEMDESC = struct([
      'void *tdesc_union_ptr', 'unsigned short tdesc_vt',   # TYPEDESC tdesc, inlined
      'void *paramdescex_ptr', 'unsigned short wParamFlags' # PARAMDESC/IDLDESC union
    ])

    FUNCDESC = struct([
      'int memid', 'char _pad1[4]',  # memid + padding to align lprgscode
      'void *lprgscode', 'void *lprgelemdescParam',
      'int funckind', 'int invkind', 'int callconv',
      'short cParams', 'short cParamsOpt', 'short oVft', 'short cScodes',
      'char _pad2[4]',  # padding to align elemdescFunc
      # elemdescFunc (ELEMDESC, inlined) — same 4-field shape as ELEMDESC above
      'void *ret_tdesc_union_ptr', 'unsigned short ret_tdesc_vt', 'char _pad3[6]',
      'void *ret_paramdescex_ptr', 'unsigned short ret_wParamFlags', 'char _pad4[6]',
      'unsigned short wFuncFlags', 'char _pad5[6]'
    ])

    VARDESC = struct([
      'int memid', 'char _pad1[4]',  # memid + padding to align lpstrSchema
      'void *lpstrSchema', 'void *union_oInst_or_lpvarValue',
      # elemdescVar (ELEMDESC, inlined) — full 32-byte block: TYPEDESC (16) + PARAMDESC/IDLDESC (16)
      'void *tdesc_union_ptr', 'unsigned short tdesc_vt', 'char _pad2[6]',  # TYPEDESC: 8+2+6=16
      'void *paramdescex_ptr', 'unsigned short wParamFlags', 'char _pad3[6]',  # PARAMDESC/IDLDESC: 8+2+6=16
      'unsigned short wVarFlags', 'int varkind'
    ])

    TYPEATTR = struct([
      'unsigned int guid_Data1', 'unsigned short guid_Data2',
      'unsigned short guid_Data3', 'unsigned char guid_Data4[8]',
      'unsigned int lcid', 'unsigned int dwReserved',
      'int memidConstructor', 'int memidDestructor',
      'void *lpstrSchema', 'unsigned int cbSizeInstance',
      'int typekind', 'unsigned short cFuncs', 'unsigned short cVars',
      'unsigned short cImplTypes', 'unsigned short cbSizeVft',
      'unsigned short cbAlignment', 'unsigned short wTypeFlags',
      'unsigned short wMajorVerNum', 'unsigned short wMinorVerNum',
      # tdescAlias (TYPEDESC, inlined) — complete 16-byte block with trailing padding
      'void *tdescAlias_union_ptr', 'unsigned short tdescAlias_vt', 'char _pad1[6]',
      # idldescType (IDLDESC: ULONG_PTR dwReserved + USHORT wIDLFlags)
      'void *idldescType_dwReserved', 'unsigned short idldescType_wIDLFlags'
    ])

    ITYPEINFO_VTBL = {
      GetTypeAttr: 3, GetFuncDesc: 5, GetVarDesc: 6, GetNames: 7,
      GetRefTypeOfImplType: 8, GetImplTypeFlags: 9, GetIDsOfNames: 10,
      Invoke: 11, GetDocumentation: 12, GetDllEntry: 13, GetRefTypeInfo: 14,
      AddressOfMember: 15, CreateInstance: 16, GetMops: 17,
      GetContainingTypeLib: 18, ReleaseTypeAttr: 19, ReleaseFuncDesc: 20,
      ReleaseVarDesc: 21
    }.freeze

    ITYPELIB_VTBL = {
      GetTypeInfoCount: 3, GetTypeInfo: 4, GetTypeInfoType: 5,
      GetTypeInfoOfGuid: 6, GetLibAttr: 7, GetTypeComp: 8,
      GetDocumentation: 9, IsName: 10, FindName: 11, ReleaseTLibAttr: 12
    }.freeze

    TYPEKIND_NAMES = {
      0 => 'Enum', 1 => 'Record', 2 => 'Module', 3 => 'Interface',
      4 => 'Dispatch', 5 => 'Class', 6 => 'Alias', 7 => 'Union', 8 => 'Max'
    }.freeze

    VARKIND_NAMES = {
      0 => 'PERINSTANCE', 1 => 'STATIC', 2 => 'CONSTANT', 3 => 'DISPATCH'
    }.freeze

    # Transcribed from MRI's ole_typedesc2val (ext/win32ole/win32ole.c) —
    # introspection type names, distinct from Win32.variant_ruby_type's much
    # smaller table (which maps the few VARTYPEs Phase 1's marshaling code
    # actually needs to send/receive, not the full set a real type library
    # can describe). VT_DISPATCH(9) and VT_UNKNOWN(13) are genuinely
    # different values with genuinely different names — do not conflate them.
    VARTYPE_NAMES = {
      0 => 'EMPTY', 1 => 'NULL', 2 => 'I2', 3 => 'I4', 4 => 'R4', 5 => 'R8',
      6 => 'CY', 7 => 'DATE', 8 => 'BSTR', 9 => 'DISPATCH', 10 => 'ERROR',
      11 => 'BOOL', 12 => 'VARIANT', 13 => 'UNKNOWN', 14 => 'DECIMAL',
      16 => 'I1', 17 => 'UI1', 18 => 'UI2', 19 => 'UI4', 20 => 'I8', 21 => 'UI8',
      22 => 'INT', 23 => 'UINT', 24 => 'VOID', 25 => 'HRESULT', 26 => 'PTR',
      27 => 'SAFEARRAY', 28 => 'CARRAY', 29 => 'USERDEFINED', 30 => 'LPSTR',
      31 => 'LPWSTR', 36 => 'RECORD'
    }.freeze

    INVOKE_FUNC = 0x1
    INVOKE_PROPERTYGET = 0x2
    INVOKE_PROPERTYPUT = 0x4
    INVOKE_PROPERTYPUTREF = 0x8

    module_function

    def vartype_name(vt)
      VARTYPE_NAMES.fetch(vt) { "Unknown Type #{vt}" }
    end

    def invoke_kind_name(invkind)
      if (invkind & INVOKE_PROPERTYGET != 0) && (invkind & INVOKE_PROPERTYPUT != 0)
        'PROPERTY'
      elsif invkind & INVOKE_PROPERTYGET != 0
        'PROPERTYGET'
      elsif invkind & INVOKE_PROPERTYPUT != 0
        'PROPERTYPUT'
      elsif invkind & INVOKE_PROPERTYPUTREF != 0
        'PROPERTYPUTREF'
      elsif invkind & INVOKE_FUNC != 0
        'FUNC'
      else
        'UNKNOWN'
      end
    end

    W = Win32
    private_constant :W

    # Every *_fn method below memoizes a resolved Fiddle::Function keyed by
    # the object's VTABLE address (W.vtable_address), not the object's own
    # address: after a COM object is released, its heap address can be
    # reused by an unrelated object with a different vtable, so keying by
    # object address risks resolving a function pointer from the wrong
    # vtable. Keying by vtable address is both correct (it's the actual
    # implementation identity) and bounded (one entry per distinct
    # implementation, not per live object).

    def get_type_info_fn(idispatch_ptr)
      @get_type_info_fns ||= {}
      @get_type_info_fns[W.vtable_address(idispatch_ptr)] ||= W.vtable_function(
        idispatch_ptr, 4, [W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def type_attr_fn(itypeinfo_ptr)
      @type_attr_fns ||= {}
      @type_attr_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetTypeAttr], [W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def func_desc_fn(itypeinfo_ptr)
      @func_desc_fns ||= {}
      @func_desc_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetFuncDesc], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def var_desc_fn(itypeinfo_ptr)
      @var_desc_fns ||= {}
      @var_desc_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetVarDesc], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def documentation_fn_for_typeinfo(itypeinfo_ptr)
      @documentation_fns ||= {}
      @documentation_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetDocumentation],
        [W::VOIDP, W::LONG, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def containing_typelib_fn(itypeinfo_ptr)
      @containing_typelib_fns ||= {}
      @containing_typelib_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetContainingTypeLib], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def release_type_attr_fn(itypeinfo_ptr)
      @release_type_attr_fns ||= {}
      @release_type_attr_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:ReleaseTypeAttr], [W::VOIDP, W::VOIDP], W::VOID
      )
    end

    def release_func_desc_fn(itypeinfo_ptr)
      @release_func_desc_fns ||= {}
      @release_func_desc_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:ReleaseFuncDesc], [W::VOIDP, W::VOIDP], W::VOID
      )
    end

    def release_var_desc_fn(itypeinfo_ptr)
      @release_var_desc_fns ||= {}
      @release_var_desc_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:ReleaseVarDesc], [W::VOIDP, W::VOIDP], W::VOID
      )
    end

    def ref_type_info_fn(itypeinfo_ptr)
      @ref_type_info_fns ||= {}
      @ref_type_info_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetRefTypeInfo], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def type_info_count_fn(itypelib_ptr)
      @type_info_count_fns ||= {}
      @type_info_count_fns[W.vtable_address(itypelib_ptr)] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:GetTypeInfoCount], [W::VOIDP], W::DWORD
      )
    end

    def type_info_fn(itypelib_ptr)
      @type_info_fns ||= {}
      @type_info_fns[W.vtable_address(itypelib_ptr)] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:GetTypeInfo], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def lib_attr_fn(itypelib_ptr)
      @lib_attr_fns ||= {}
      @lib_attr_fns[W.vtable_address(itypelib_ptr)] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:GetLibAttr], [W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def documentation_fn_for_typelib(itypelib_ptr)
      @typelib_documentation_fns ||= {}
      @typelib_documentation_fns[W.vtable_address(itypelib_ptr)] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:GetDocumentation],
        [W::VOIDP, W::LONG, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def release_tlib_attr_fn(itypelib_ptr)
      @release_tlib_attr_fns ||= {}
      @release_tlib_attr_fns[W.vtable_address(itypelib_ptr)] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:ReleaseTLibAttr], [W::VOIDP, W::VOIDP], W::VOID
      )
    end

    HKEY_CLASSES_ROOT = 0x80000000
    KEY_READ = 0x20019
    REG_SZ = 1

    def advapi32
      @advapi32 ||= Fiddle.dlopen('advapi32')
    end

    def reg_open_key_ex
      @reg_open_key_ex ||= Fiddle::Function.new(
        advapi32['RegOpenKeyExW'], [W::VOIDP, W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def reg_query_value_ex
      @reg_query_value_ex ||= Fiddle::Function.new(
        advapi32['RegQueryValueExW'],
        [W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def reg_close_key
      @reg_close_key ||= Fiddle::Function.new(advapi32['RegCloseKey'], [W::VOIDP], W::LONG, W::STDCALL)
    end
  end
end
