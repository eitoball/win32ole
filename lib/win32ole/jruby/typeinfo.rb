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

    INVOKE_FUNC = 0x1
    INVOKE_PROPERTYGET = 0x2
    INVOKE_PROPERTYPUT = 0x4
    INVOKE_PROPERTYPUTREF = 0x8

    module_function

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
  end
end
