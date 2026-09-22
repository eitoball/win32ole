# JRuby support for win32ole — Phase 2 (type library introspection) design

Date: 2026-09-22
Status: proposed

## 1. Background

Phase 1 (`docs/superpowers/specs/2026-09-22-jruby-win32ole-support-design.md`)
shipped a pure-Ruby + Fiddle `WIN32OLE` core for JRuby: construction via
ProgID/CLSID, dynamic dispatch through `method_missing`, basic
Ruby↔VARIANT type marshaling, and error translation matching the MRI C
extension's message format. It landed, was verified against real
Windows+JRuby CI across six iterations (several real bugs were found and
fixed only by execution, not by review — a missing `CoInitialize` call, a
wrong `DISPATCH_*` flag combination, a JRuby-specific dangling-pointer bug
in the keep-alive discipline), and Phase 1's own roadmap (§5) explicitly
deferred Phase 2's design to "once Phase 1 has landed and the patterns it
establishes (file layout, keep-alive discipline, error translation) are
proven against the real test suite." That condition is now met.

Phase 2 is `WIN32OLE::TYPE`, `TYPELIB`, `METHOD`, `PARAM`, `VARIABLE` —
type-library introspection via `ITypeInfo`/`ITypeLib`, the two "custom"
(non-`IDispatch`) COM interfaces with fixed vtables that MRI's
`ext/win32ole/win32ole_type.c` (923 lines), `win32ole_typelib.c` (849
lines), `win32ole_method.c` (955 lines), `win32ole_param.c` (441 lines),
and `win32ole_variable.c` (386 lines) implement.

### 1.1 What the core mechanism actually is (verified against MRI source)

Reading the real C extension resolved what could otherwise have been
design guesswork:

- A `WIN32OLE` instance's own type (`excel.ole_type`) is obtained via
  `IDispatch::GetTypeInfo(0, lcid, &pTypeInfo)` — vtable slot **4** on the
  same `IDispatch*` Phase 1 already holds in `@ptr` (slots 2/5/6 are
  already wired for `Release`/`GetIDsOfNames`/`Invoke`; this is one more
  slot on the identical interface, not a new one). `WIN32OLE::Type` then
  wraps that `ITypeInfo*` **directly** (`ext/win32ole/win32ole.c:3457`,
  `fole_type`) — no name lookup, no registry.
- `WIN32OLE::Type#ole_typelib` is `ITypeInfo::GetContainingTypeLib` on
  that same `ITypeInfo*`, wrapping the resulting `ITypeLib*` directly
  (`ext/win32ole/win32ole_type.c:92`).
- `WIN32OLE#ole_methods` (and `ole_get_methods`/`ole_put_methods`/
  `ole_func_methods`) go through a slightly different path
  (`typeinfo_from_ole`, `win32ole.c:3322`): `GetTypeInfo` →
  `GetDocumentation` (get this type's own name) → `GetContainingTypeLib`
  → iterate `GetTypeInfoCount`/`GetDocumentation` on the typelib to find
  the entry whose name matches → `GetTypeInfo(i)` again. This name
  round-trip (instead of just reusing the first `ITypeInfo*` the way
  `fole_type` does) looks redundant on paper. It is carried into this
  design unexamined — §8 flags it as an implementation-time question
  rather than assumed away, since silently "simplifying" it without
  understanding why MRI does it risks losing behavior for `IDispatch`
  implementations where `GetTypeInfo(0, ...)` returns a coclass's type
  info rather than the dispinterface's.
- None of this requires the Windows registry (`HKEY_CLASSES_ROOT\TypeLib`)
  or `LoadTypeLib`/`LoadRegTypeLib`. Registry-based lookup is only needed
  for API surface this design explicitly excludes (§3).

### 1.2 What makes this phase harder than Phase 1

`ITypeInfo::GetTypeAttr`/`GetFuncDesc`/`GetVarDesc` return pointers to
`TYPEATTR`/`FUNCDESC`/`VARDESC` — C structs far more complex than
Phase 1's `VARIANT`/`DISPPARAMS`:

- `TYPEATTR`: 16 fields including a 16-byte `GUID`, multiple `WORD`
  counts, and a nested `TYPEDESC`.
- `FUNCDESC`: 12 fields including two enums (`FUNCKIND`, `INVOKEKIND`,
  `CALLCONV` — three, actually), a pointer to an array of `ELEMDESC`
  (length `cParams`, itself a struct containing a union), and a nested
  `ELEMDESC` for the return type.
- `VARDESC`: 6 fields including a union and a nested `ELEMDESC`.

Phase 1 already hit two real, execution-only-discovered layout bugs on
much simpler structs (`sizeof(VARIANT)` being 24 not 16 bytes on x64;
`DISPPARAMS`'s pointer fields needing native width, not a hardcoded 8
bytes). Hand-deriving offset tables (Phase 1's `EXCEPINFO_OFFSETS` style)
for structs this much more nested multiplies that error surface. §2
addresses this directly.

## 2. Goals

- Implement `WIN32OLE::TYPE`, `TYPELIB`, `METHOD`, `PARAM`, `VARIABLE`
  for the core, already-proven-necessary use case: given a live
  `WIN32OLE` instance (or a `WIN32OLE::Type` obtained from one),
  introspect its members — `excel.ole_type`, `excel.ole_methods`,
  `method.params`, `type.variables`, and their attributes (VARTYPE,
  parameter direction, help strings, DISPIDs, etc.).
- Reuse Phase 1's established patterns without re-litigating them: the
  `win32.rb` (pure logic + native substrate) / `dispatch.rb` (mixin)
  split, the per-instance finalizer pattern for COM interface pointers,
  `WIN32OLE::RuntimeError`/`QueryInterfaceError` for error translation,
  and the "new dedicated behavior, reviewed against real Windows+JRuby
  CI" verification loop.
- One deliberate, explicit deviation from a Phase 1 global constraint:
  use `Fiddle::CStructBuilder`/`Fiddle::Importer.struct` for the four
  complex structs (§2.1), rather than hand-computed offset tables.

### 2.1 Why this phase uses Fiddle's own struct-building facilities

Phase 1's global constraint ("no `FFI::Struct`/`Fiddle::Importer`
struct classes — byte-pack/unpack with `Array#pack`/`String#unpack`,
io-console style") was right for Phase 1's structs (`VARIANT`,
`DISPPARAMS`, `EXCEPINFO`): flat-enough, and few enough of them, that
hand-computed offsets were tractable and got caught by CI when wrong.
`TYPEATTR`/`FUNCDESC`/`VARDESC`/`ELEMDESC` are qualitatively more
complex (§1.2), and Phase 1 already proved by direct experience — twice
— that hand-computed offsets on real Win32 structs are an easy, silent
way to get x86/x64 alignment wrong. For this phase, `Fiddle::CStructBuilder`
(`Fiddle::Importer.struct([...])`, part of the `fiddle` standard library
on both MRI and JRuby) computes correct member offsets and total size
from a field-type declaration list, the same way a C compiler would —
removing an entire class of bug this codebase has already hit twice, at
the cost of one line of stated inconsistency with Phase 1's own
byte-packing style.

**This needs empirical confirmation, exactly the way Phase 1 confirmed
`STDCALL`/`VARIANT_SIZE` differences before relying on them**: does
JRuby's Fiddle backend (`fiddle/ffi_backend.rb`) actually implement
`Fiddle::CStructBuilder`/`Fiddle::Importer.struct`, and does the struct
it builds report the same `.size` and per-member offsets as MRI's Fiddle
for the exact same field-type list on both x86 and x64? This is an
implementation-time verification step (§8), not assumed true because it
"should" work — Phase 1's own experience is the reason to check rather
than assume.

## 3. Non-goals (for this document / Phase 2)

- **Registry *enumeration***: `WIN32OLE::TypeLib.typelibs`,
  `WIN32OLE::TYPE.typelibs`/`.ole_classes`/`.progids`, and
  `WIN32OLE::TypeLib.new(name, version)`'s registry-search constructor.
  These walk the entire `HKEY_CLASSES_ROOT\TypeLib` subtree
  (`RegEnumKeyEx` over an unknown number of children) and need a
  distinct traversal this design doesn't build. Lower usage frequency
  than instance-based introspection, and independent enough to be its
  own follow-up design. (This excludes *walking* the registry, not
  every registry read — §5's `TypeLib#path` still needs one targeted,
  single-key lookup, which is a different, much smaller thing; see
  there for why that one stays in scope.)
- **`ImplType` traversal**: `WIN32OLE::Type#implemented_ole_types`,
  `#source_ole_types`, `#default_event_sources`, `#default_ole_types`,
  and `WIN32OLE::Method#event?`/`#event_interface`. These walk a type's
  secondary/implemented interfaces via `GetRefTypeOfImplType`/
  `GetRefTypeInfo`/`GetImplTypeFlags` and are specifically how Phase 4
  (`WIN32OLE::Event`) will need to find a coclass's default outgoing
  (event) interface — deferred to be designed alongside Phase 4, not
  half-built here.
- `WIN32OLE#ole_method_help`/`#ole_obj_help` (launch a `.chm`/`.hlp`
  help file — a Windows-desktop-interaction feature, not a data-access
  one) and `#ole_query_interface` (QI for an arbitrary caller-supplied
  IID — a distinct, general-purpose capability rather than typelib
  introspection).
- `WIN32OLE::Type.new(typelib, ole_class)`'s public, name-based
  constructor. The only construction path this design implements is the
  internal one (wrapping an already-obtained `ITypeInfo*`/`ITypeLib*`),
  since that's what every in-scope entry point in §4 actually needs;
  name-based lookup requires either registry access or enumerating an
  already-open `WIN32OLE::TypeLib`'s `ole_types` (in scope) to find a
  match by name, which a caller can already do in plain Ruby with
  `tlib.ole_types.find { |t| t.name == 'Workbook' }` without a dedicated
  constructor.
- Encoding/locale fidelity beyond what Phase 1 already established
  (inherits Phase 1's §8 risk #5, not resolved further here).
- x86 (32-bit) verification (inherits Phase 1's §8 risk #3 — the struct
  layouts below are written to be correct on paper for both via
  `Fiddle::CStructBuilder`'s own platform-aware offset computation, but
  only x64 CI actually exercises this).

## 4. Architecture

### 4.1 File layout

```
lib/win32ole/jruby/
  typeinfo.rb      # NEW: Fiddle::CStructBuilder struct declarations for
                   # TYPEATTR/FUNCDESC/VARDESC/ELEMDESC/TYPEDESC, plus
                   # ITypeInfo/ITypeLib vtable-slot constants and the
                   # enum→string lookup tables (§4.4)
  type.rb          # WIN32OLE::Type
  typelib.rb       # WIN32OLE::TypeLib
  method.rb        # WIN32OLE::Method
  param.rb         # WIN32OLE::Param
  variable.rb      # WIN32OLE::Variable
```

This mirrors the C extension's own file boundaries (`win32ole_type.c`,
`win32ole_typelib.c`, ...), continuing Phase 1's stated rationale for
that choice: cross-referencing behavior against the reference
implementation during a 1:1-port task stays straightforward.
`typeinfo.rb` is this phase's counterpart to Phase 1's `win32.rb` — the
shared low-level substrate every other new file depends on — kept
separate from `win32.rb` itself since it's a distinct concern (custom
COM interfaces and their structs, not `IDispatch` and `VARIANT`) with no
reason to grow the existing file.

### 4.2 `typeinfo.rb`: the struct/vtable substrate

- **Vtable slot constants**, reusing Phase 1's `Win32.vtable_function`
  substrate unchanged (it already takes an arbitrary object address,
  slot index, and signature — nothing about it is `IDispatch`-specific):
  - `ITypeInfo` (standard, fixed layout after the 3 `IUnknown` slots):
    `GetTypeAttr`=3, `GetFuncDesc`=5, `GetVarDesc`=6, `GetNames`=7,
    `GetRefTypeOfImplType`=8, `GetImplTypeFlags`=9, `GetIDsOfNames`=10,
    `Invoke`=11, `GetDocumentation`=12, `GetDllEntry`=13,
    `GetRefTypeInfo`=14, `AddressOfMember`=15, `CreateInstance`=16,
    `GetMops`=17, `GetContainingTypeLib`=18, `ReleaseTypeAttr`=19,
    `ReleaseFuncDesc`=20, `ReleaseVarDesc`=21. Phase 2 only needs
    3/5/6/12/18/19/20/21 (marked in bold intent, not literally bolded
    here) — the rest exist in the table for completeness/documentation
    but aren't bound to `Fiddle::Function`s until something needs them.
  - `ITypeLib`: `GetTypeInfoCount`=3, `GetTypeInfo`=4, `GetTypeInfoType`=5,
    `GetTypeInfoOfGuid`=6, `GetLibAttr`=7, `GetTypeComp`=8,
    `GetDocumentation`=9, `IsName`=10, `FindName`=11, `ReleaseTLibAttr`=12.
    Phase 2 needs 3/4/7/9/12 (`GetTypeInfoCount`/`GetTypeInfo` for
    `TypeLib#ole_types`, `GetLibAttr`/`ReleaseTLibAttr` for `guid`/
    `version`/`major_version`/`minor_version`, `GetDocumentation` for
    `name`/`helpstring`/`helpfile`/`helpcontext`... on `WIN32OLE::TypeLib`
    itself, not just on member lookups).
- **Struct declarations** via `Fiddle::Importer.struct`, one per C
  struct, each guarded the same way Phase 1's `VARIANT_SIZE`/`PTR_SIZE`
  branch is (`Fiddle::SIZEOF_VOIDP`-conditional where a field is
  pointer-sized), for example:

  ```ruby
  GUID = Fiddle::Importer.struct([
    'unsigned long Data1', 'unsigned short Data2', 'unsigned short Data3',
    'unsigned char Data4[8]'
  ])

  TYPEDESC = Fiddle::Importer.struct([
    'void *union_ptr',  # lptdesc / lpadesc / hreftype share this slot;
                         # read as whichever the caller's context expects
    'unsigned short vt'
  ])

  ELEMDESC = Fiddle::Importer.struct([
    'void *tdesc_union_ptr', 'unsigned short tdesc_vt', # TYPEDESC tdesc, inlined
    'void *paramdescex_ptr', 'unsigned short wParamFlags' # PARAMDESC/IDLDESC union
  ])

  FUNCDESC = Fiddle::Importer.struct([
    'long memid', 'void *lprgscode', 'void *lprgelemdescParam',
    'int funckind', 'int invkind', 'int callconv',
    'short cParams', 'short cParamsOpt', 'short oVft', 'short cScodes',
    # elemdescFunc (ELEMDESC, inlined) — 4 fields, same shape as ELEMDESC above
    'void *ret_tdesc_union_ptr', 'unsigned short ret_tdesc_vt',
    'void *ret_paramdescex_ptr', 'unsigned short ret_wParamFlags',
    'unsigned short wFuncFlags'
  ])

  VARDESC = Fiddle::Importer.struct([
    'long memid', 'void *lpstrSchema', 'void *union_oInst_or_lpvarValue',
    # elemdescVar (ELEMDESC, inlined)
    'void *tdesc_union_ptr', 'unsigned short tdesc_vt',
    'void *paramdescex_ptr', 'unsigned short wParamFlags',
    'unsigned short wVarFlags', 'int varkind'
  ])

  TYPEATTR = Fiddle::Importer.struct([
    'unsigned long guid_Data1', 'unsigned short guid_Data2',
    'unsigned short guid_Data3', 'unsigned char guid_Data4[8]',
    'unsigned long lcid', 'unsigned long dwReserved',
    'long memidConstructor', 'long memidDestructor',
    'void *lpstrSchema', 'unsigned long cbSizeInstance',
    'int typekind', 'unsigned short cFuncs', 'unsigned short cVars',
    'unsigned short cImplTypes', 'unsigned short cbSizeVft',
    'unsigned short cbAlignment', 'unsigned short wTypeFlags',
    'unsigned short wMajorVerNum', 'unsigned short wMinorVerNum',
    # tdescAlias (TYPEDESC, inlined) + idldescType (IDLDESC: void* + unsigned short)
    'void *tdescAlias_union_ptr', 'unsigned short tdescAlias_vt',
    'unsigned long idldescType_dwReserved', 'unsigned short idldescType_wIDLFlags'
  ])
  ```

  These field lists are the design's best-effort transcription of the
  real `oaidl.h` layouts (unions flattened into their widest member,
  nested structs inlined field-by-field, exactly the way Phase 1 already
  treats `VARIANT`'s value union). **They are a starting point for
  implementation, not verified byte-for-byte against a live Windows
  build** — the implementation plan's first task must build each struct
  standalone and assert `.size`/member `.offset` against known reference
  values (Microsoft's documented struct layouts, or by reading the real
  values off a live `ITypeInfo` in CI) before anything else depends on
  them, exactly the discipline Phase 1 applied to `VARIANT_SIZE`.
- **Enum → human-readable-string lookup tables**, transcribed verbatim
  from the switch statements MRI's C extension actually uses (§4.4)
  rather than re-derived from first principles.
- **`native_pointer_for`/keep-alive discipline**: reused from Phase 1
  unchanged (`Win32.native_pointer_for`, holding the returned
  `Fiddle::Pointer` — not just its address — in a local variable through
  any native call that dereferences it, per the dangling-pointer lesson
  Phase 1 learned the hard way).

### 4.3 Class responsibilities

- **`WIN32OLE::Type`**: wraps an `ITypeInfo*`. `Type#new`-equivalent is
  a private/internal constructor (`Type.send(:from_typeinfo_ptr, ptr)`
  or similar — exact spelling is an implementation detail) called from
  `WIN32OLE#ole_type`/`#ole_methods` and `WIN32OLE::TypeLib#ole_types`,
  never a public name-based lookup (§3). Holds `@type_attr` — the
  `TYPEATTR` fields read once at construction and released immediately
  (§4.5) — plus the live `ITypeInfo*` (finalized like Phase 1's
  `WIN32OLE`).
- **`WIN32OLE::TypeLib`**: wraps an `ITypeLib*`, obtained only via
  `ITypeInfo::GetContainingTypeLib` (§1.1) in this phase. Same
  finalizer pattern as `Type`.
- **`WIN32OLE::Method`**: wraps a `(ITypeInfo*, function_index)` pair —
  it needs the owning `ITypeInfo*` alive for `params` to later call
  `GetFuncDesc` again (once per `Method` instance, at construction,
  producing all of the `Param` objects up front — no lazy re-fetching),
  but does not itself hold a separate finalizable interface pointer; it
  keeps the owning `Type`'s `ITypeInfo*` reachable via a plain instance
  variable reference to the `Type` object itself, which is enough to
  keep it (and therefore its `ITypeInfo*`) alive for as long as any
  `Method` built from it exists — standard Ruby object-graph reachability,
  no new mechanism.
- **`WIN32OLE::Param`**: pure data, built from one `ELEMDESC` (a
  `FUNCDESC`'s `lprgelemdescParam[i]`) at the same time as its owning
  `Method`. No COM reference of its own at all.
- **`WIN32OLE::Variable`**: wraps a `(ITypeInfo*, variable_index)` pair,
  same reachability-via-owning-Type pattern as `Method`.

### 4.4 Enum → string mappings (transcribed from MRI, §1.2/§4.2)

```ruby
TYPEKIND_NAMES = {
  0 => 'Enum', 1 => 'Record', 2 => 'Module', 3 => 'Interface',
  4 => 'Dispatch', 5 => 'Class', 6 => 'Alias', 7 => 'Union', 8 => 'Max'
}.freeze # TKIND_ENUM..TKIND_MAX, used by Type#ole_type (Type#typekind
         # returns the raw integer, matching MRI's own two-method split)

VARKIND_NAMES = {
  0 => 'PERINSTANCE', 1 => 'STATIC', 2 => 'CONSTANT', 3 => 'DISPATCH'
}.freeze # VAR_PERINSTANCE..VAR_DISPATCH, used by Variable#variable_kind
         # (Variable#varkind returns the raw integer); "UNKNOWN" fallback
         # for anything outside 0..3, matching MRI's default case

# Method#invoke_kind: derived from the INVOKEKIND *bitmask* (not a plain
# enum switch) exactly as MRI computes it — check PROPERTYGET+PROPERTYPUT
# both set (=> "PROPERTY") before checking either alone:
#   INVOKE_FUNC = 0x1, INVOKE_PROPERTYGET = 0x2, INVOKE_PROPERTYPUT = 0x4,
#   INVOKE_PROPERTYPUTREF = 0x8
# get & put both set    => "PROPERTY"
# get only               => "PROPERTYGET"
# put only               => "PROPERTYPUT"
# putref only            => "PROPERTYPUTREF"
# func only               => "FUNC"
# none of the above       => "UNKNOWN"
# (Method#invkind returns the raw bitmask integer.)
```

### 4.5 Resource lifetime (§2 of the plan this feeds will need this precisely)

- `Type`/`TypeLib`: live `ITypeInfo*`/`ITypeLib*`, released via the
  Phase-1-identical finalizer pattern (`install_finalizer`, captures
  only the raw pointer + a memoized `Release` `Fiddle::Function` —
  never `self`).
- `Type`'s `TYPEATTR`, `Method`'s `FUNCDESC`, `Variable`'s `VARDESC`:
  read-then-release. Call `GetTypeAttr`/`GetFuncDesc`/`GetVarDesc`,
  copy every field this design's API surface (§ per-class tables below)
  needs into plain Ruby ivars, then immediately call
  `ReleaseTypeAttr`/`ReleaseFuncDesc`/`ReleaseVarDesc` before the
  constructor returns. No finalizer needed for these three classes —
  simpler than Phase 1's `WIN32OLE`, since typelib metadata is static
  and there is nothing to re-invoke later the way a live dispatch call
  is.
- `Param`'s `ELEMDESC`: never independently released — it's a slice of
  its owning `FUNCDESC`'s `lprgelemdescParam` array, freed as part of
  that single `ReleaseFuncDesc` call. `Param` objects only read fields
  out of it during their owning `Method`'s construction; they hold no
  pointer afterward.

## 5. Per-class API (Phase 2 scope)

Public method surface this phase implements; anything not listed here
raises `NotImplementedError` (matching Phase 1's own convention for
out-of-scope members, §6.3 of the Phase 1 design).

| Class | Implemented | Notes |
|---|---|---|
| `WIN32OLE::TypeLib` | `guid`, `name`, `version`, `major_version`, `minor_version`, `path`, `visible?`, `library_name`, `ole_types`, `inspect` | `path` needs the typelib's on-disk file path — MRI reads this from the registry (`reg_get_typelib_file_path`, keyed by the typelib's GUID+version); **this is the one place §3's registry exclusion bites even a core method**. Resolved here: implement `path` via the same GUID+version-keyed single registry read MRI does (`RegOpenKeyEx`/`RegQueryValueEx` on one known key, not an enumeration) — this is a single targeted lookup, not the `.typelibs`-style full-tree walk §3 excludes, so it needs exactly one new pair of registry API bindings (`RegOpenKeyEx`, `RegQueryValueEx`, `RegCloseKey`) in `typeinfo.rb`. If this turns out more involved during implementation than this paragraph assumes, `path` degrades to `NotImplementedError` rather than blocking the rest of the class — flagged in §8. |
| `WIN32OLE::Type` | `name`, `ole_type`, `guid`, `progid`, `visible?`, `major_version`, `minor_version`, `typekind`, `helpstring`, `src_type`, `helpfile`, `helpcontext`, `variables`, `ole_methods`, `ole_typelib`, `inspect` | `progid` (`ProgIDFromCLSID`) is a single targeted OS call (`oleaut32`/`ole32`, not a registry walk) — in scope. `src_type` (for `TKIND_ALIAS` — the aliased type's name) needs `TYPEDESC` union interpretation when `vt == VT_USERDEFINED`, calling `GetRefTypeInfo` on the embedded `HREFTYPE` — in scope, uses vtable slot 14. |
| `WIN32OLE::Method` | `name`, `return_type`, `return_vtype`, `return_type_detail`, `invoke_kind`, `invkind`, `visible?`, `helpstring`, `helpfile`, `helpcontext`, `dispid`, `offset_vtbl`, `size_params`, `size_opt_params`, `params`, `inspect` | `return_type_detail` returns an array describing a possibly-nested `TYPEDESC` (e.g. `["PTR", "VT_UI1"]` for a `BYTE*` return) — the one place `TYPEDESC`'s self-referential union (`lptdesc` pointing at another `TYPEDESC`, for pointer/array/SAFEARRAY-of-X types) must actually be walked recursively rather than read once. |
| `WIN32OLE::Param` | `name`, `ole_type`, `ole_type_detail`, `input?`, `output?`, `optional?`, `retval?`, `default`, `inspect` | `default` (a parameter's default value, from `PARAMDESCEX.varDefaultValue` when `PARAMFLAG_FHASDEFAULT` is set) reuses Phase 1's `Win32.variant_ruby_type`/unpack helpers directly — a `VARIANTARG` is a `VARIANTARG` regardless of which COM API produced it. |
| `WIN32OLE::Variable` | `name`, `ole_type`, `ole_type_detail`, `value`, `visible?`, `variable_kind`, `varkind`, `inspect` | `value` (for `VAR_CONST` members) also reuses Phase 1's VARIANT unpack helpers, reading `VARDESC`'s `lpvarValue` union member. |
| `WIN32OLE` (additions) | `ole_type`, `ole_methods`, `ole_get_methods`, `ole_put_methods`, `ole_func_methods`, `ole_typelib`, `ole_respond_to?` | `ole_respond_to?(name)` is `dispid_for(name)` returning non-nil — already-existing Phase 1 machinery, exposed as public API; no new native calls. |

## 6. Error translation

`GetTypeInfo`/`GetContainingTypeLib`/`GetFuncDesc`/`GetVarDesc`/
`GetTypeAttr` failures raise `WIN32OLE::QueryInterfaceError` (defined in
Phase 1, never raised there — this phase is its first real caller),
matching `ext/win32ole/win32ole.c`'s own choice of exception class for
exactly these failure sites (`eWIN32OLEQueryInterfaceError`, not
`eWIN32OLERuntimeError`). Message format follows Phase 1's established
`Win32.method_error_message`-style helpers (a new
`Win32.query_interface_error_message(operation, detail)` alongside them,
same shape).

## 7. Testing / CI strategy

Unlike Phase 1 (where almost none of the pre-existing legacy suite
applied, since it exercises exactly the features Phase 1 didn't build),
Phase 2 directly implements what `test/win32ole/test_win32ole_type.rb`,
`test_win32ole_typelib.rb`, `test_win32ole_method.rb`,
`test_win32ole_param.rb`, and `test_win32ole_variable.rb` (614 lines
total) already test. Run them unmodified against the new backend
(they already run under the existing `test-jruby` CI job, per Phase 1's
own `windows.yml` addition) and triage failures the same disciplined way
Phase 1 did after each CI run: confirm each failure is either (a) exactly
the registry/`ImplType` scope this document excludes (§3) — expected,
not actionable — or (b) a real Phase 2 bug — fix and re-verify. Add a
`GC.stress`-enabled test for the `Type`/`TypeLib` finalizer path (the
class of bug that "passes every normal run and then crashes
intermittently," per Phase 1's own §7 rationale for the same kind of
test there) — no dedicated test needed for the read-then-release
classes (`Method`/`Param`/`Variable`), since they hold no COM reference
whose finalization timing could matter.

Struct layout verification (§4.2) is its own test, independent of any
live COM object: assert `Fiddle::Importer.struct([...]).size` and each
member's byte offset against known values before anything else in this
phase is built on top — this can and should run locally (no Windows/COM
needed, exactly like Phase 1's win32.rb pure-logic layer), and then
again in CI to confirm JRuby's Fiddle backend agrees with MRI's (§2.1's
open verification question).

## 8. Risks / open questions carried forward

1. **`Fiddle::CStructBuilder`/`Fiddle::Importer.struct` support on
   JRuby is unverified** (§2.1) — the central technical bet of this
   phase. Needs empirical confirmation (matching struct `.size`/offsets
   between MRI and JRuby for the same field-type list) before
   implementation proceeds past the first task.
2. **The struct field-type transcriptions in §4.2 are unverified against
   a live Windows build** — written from `oaidl.h` knowledge, not
   confirmed byte-for-byte. The implementation plan's first task must
   validate each one (§4.2, §7) before any class depends on it.
3. **`ole_methods`'s name-round-trip-via-typelib (§1.1) is not
   understood, only observed.** Implementing it as a literal port
   (round-trip through `GetDocumentation`+name-match, matching MRI
   exactly) is the safe default; understanding *why* MRI doesn't just
   reuse the first `ITypeInfo*` the way `fole_type` does (coclass vs.
   dispinterface type info returned by `GetTypeInfo(0, ...)`, most
   likely) is an implementation-time investigation, not a blocking
   question — the port works either way, but the "why" affects whether
   a future simplification would be safe.
4. **`TypeLib#path`'s single-key registry read (§5) may prove more
   involved than assumed** — if so, `path` alone degrades to
   `NotImplementedError` without blocking the rest of `WIN32OLE::TypeLib`.
5. **x86 (32-bit) unverified** — inherited from Phase 1 §8 risk #3,
   unchanged. `Fiddle::CStructBuilder`'s own platform-aware offset
   computation (§2.1) should make this phase's structs *more* portable
   than Phase 1's hand-computed ones, but "should" is exactly the word
   Phase 1's own history (§1.2) warns against trusting without running it.
6. **Performance**: one `Method`/`Variable`/`Param` object per member,
   eagerly constructed from `ole_methods`/`variables`, means a
   large-interface COM object (Excel's `Application`, hundreds of
   members) does hundreds of `GetFuncDesc`/`ReleaseFuncDesc` round-trips
   up front. Matches MRI's own eager behavior (`ole_methods_from_typeinfo`
   builds the whole array in one call, not lazily) — not a regression
   Phase 2 introduces, but worth naming as a known cost, consistent with
   Phase 1's own accepted-performance-tradeoff stance (§8 risk #4 there).
