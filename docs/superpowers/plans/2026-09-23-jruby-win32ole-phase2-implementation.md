# JRuby win32ole Phase 2 (type library introspection) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement `WIN32OLE::TYPE`, `TYPELIB`, `METHOD`, `PARAM`, `VARIABLE` for JRuby — instance-based type-library introspection (`excel.ole_type`, `excel.ole_methods`, `method.params`, `type.variables`) — on top of the already-shipped Phase 1 `WIN32OLE` core.

**Architecture:** A new `lib/win32ole/jruby/typeinfo.rb` substrate (vtable slot constants for `ITypeInfo`/`ITypeLib`, `Fiddle::CStructBuilder`-declared structs for `TYPEATTR`/`FUNCDESC`/`VARDESC`/`ELEMDESC`/`TYPEDESC`, and enum→string lookup tables) underlies five new class files that mirror the MRI C extension's own file boundaries. `WIN32OLE::Type`/`TypeLib` wrap live, finalized COM interface pointers exactly like Phase 1's `WIN32OLE`; `WIN32OLE::Method`/`Param`/`Variable` are read-then-release data snapshots with no finalizer.

**Tech Stack:** Ruby stdlib `fiddle`, specifically `Fiddle::Importer.struct` (the one deliberate departure from Phase 1's hand-computed-offset style — justified in the spec §2.1 by Phase 1 having already hit two real layout bugs on much simpler structs).

**Spec:** `docs/superpowers/specs/2026-09-22-jruby-win32ole-phase2-typelib-design.md` — this plan implements §4-§7 in full; §3's exclusions (registry enumeration, `ImplType` traversal) are out of scope here, and §8's six open risks are addressed by name in the relevant tasks below, not assumed away.

## Global Constraints

- **Reuse Phase 1's substrate unchanged**: `Win32.vtable_function(object_addr, index, arg_types, ret_type)`, `Win32.native_pointer_for(buffer)` (hold the returned `Fiddle::Pointer` — not just its address — in a local variable/array through any native call that dereferences it; this is the exact dangling-pointer lesson Phase 1 learned from real CI failures), `Win32.sys_free_string`, `Win32.bstr_to_s`, `Win32.variant_ruby_type`/`ruby_to_variant_type`/pack/unpack helpers, `WIN32OLE::RuntimeError`/`QueryInterfaceError`.
- **`GetTypeInfo`/`GetContainingTypeLib`/`GetFuncDesc`/`GetVarDesc`/`GetTypeAttr` failures raise `WIN32OLE::QueryInterfaceError`** (spec §6) — not `WIN32OLE::RuntimeError`, matching `ext/win32ole/win32ole.c`'s own choice of exception class for exactly these failure sites.
- **`Type`/`TypeLib` get the Phase-1-identical finalizer pattern** (`install_finalizer`-equivalent: `ObjectSpace.define_finalizer` on a proc that captures only the raw pointer + a memoized `Release` `Fiddle::Function`, never `self`). **`Method`/`Param`/`Variable` get no finalizer at all** — `TYPEATTR`/`FUNCDESC`/`VARDESC` are read-then-released synchronously at construction (spec §4.5).
- **No public name-based constructors.** `WIN32OLE::Type`/`TypeLib` are only ever constructed by wrapping an already-obtained `ITypeInfo*`/`ITypeLib*` pointer (spec §3, §4.3).
- **Registry *enumeration* is out of scope** (`.typelibs`, `.ole_classes`, `.progids`) — but `TypeLib#path`'s single-key registry read is in scope (spec §5, §3).
- **`ImplType` traversal is out of scope**: `implemented_ole_types`, `source_ole_types`, `default_event_sources`, `default_ole_types`, `Method#event?`/`#event_interface` all raise `NotImplementedError`.
- **Struct field lists and vtable slot numbers below are transcribed from the spec (§4.2), which itself flags them as "a starting point for implementation, not verified byte-for-byte against a live Windows build."** Task 1 exists specifically to validate them — do not treat the numbers in this plan as ground truth without that validation passing.
- **Verification reality**: identical to Phase 1 — pure-logic/struct-declaration work (no native calls) is locally testable on any OS/engine right now; anything calling `ITypeInfo`/`ITypeLib`/registry APIs is Windows+COM-only and can only be verified by pushing to the `test-jruby` CI job already wired up in `.github/workflows/windows.yml`.

---

## Task 1: `typeinfo.rb` — struct declarations + layout verification

**Files:**
- Create: `lib/win32ole/jruby/typeinfo.rb`
- Test: `test/win32ole/jruby/test_typeinfo.rb`

**Interfaces:**
- Produces: `WIN32OLE::TypeInfo::GUID`, `::TYPEDESC`, `::ELEMDESC`, `::FUNCDESC`, `::VARDESC`, `::TYPEATTR` — `Fiddle::CStructBuilder`-built struct classes (each has `.size`, `.malloc`, and instance `[]`/`[]=` member access once wrapped around a `Fiddle::Pointer`).

This is the file's *only* content for this task — no vtable calls, no `Fiddle.dlopen`, nothing Windows-specific. Every test below runs on this machine right now, on any engine, exactly like Phase 1's `win32.rb` pure-logic layer.

**Why these particular field lists:** transcribed directly from `oaidl.h`'s real layouts, with C unions flattened to their widest member (the same treatment Phase 1 already gives `VARIANT`'s value union) and nested structs inlined field-by-field. The byte-size assertions in Step 1 below are this plan author's own hand-derivation (alignment rules: each field aligns to its own size up to the platform pointer width, and the whole struct's size rounds up to its largest member's alignment) — **if `Fiddle::Importer.struct`'s computed `.size` disagrees with the assertion here, do not "fix" the test to match it. First cross-check the disagreement against Microsoft's own documented struct layouts (search "TYPEATTR structure win32", "FUNCDESC structure win32", etc.) to determine which one is actually wrong** — a struct-layout bug here would silently corrupt every `Method`/`Param`/`Variable`/`Type` read built on top of it in later tasks.

- [ ] **Step 1: Write the failing test**

```ruby
# test/win32ole/jruby/test_typeinfo.rb
require 'test/unit'
require 'win32ole/jruby/typeinfo'

class TestTypeInfo < Test::Unit::TestCase
  TI = WIN32OLE::TypeInfo
  PTR64 = Fiddle::SIZEOF_VOIDP == 8

  def test_guid_size_is_16_bytes_on_every_platform
    assert_equal(16, TI::GUID.size)
  end

  def test_typedesc_size
    assert_equal(PTR64 ? 16 : 8, TI::TYPEDESC.size)
  end

  def test_elemdesc_size
    assert_equal(PTR64 ? 32 : 16, TI::ELEMDESC.size)
  end

  def test_funcdesc_size
    assert_equal(PTR64 ? 88 : 52, TI::FUNCDESC.size)
  end

  def test_vardesc_size
    assert_equal(PTR64 ? 64 : 36, TI::VARDESC.size)
  end

  def test_typeattr_size
    assert_equal(PTR64 ? 88 : 76, TI::TYPEATTR.size)
  end

  def test_typeattr_guid_is_at_offset_zero
    buf = TI::TYPEATTR.malloc
    buf.guid_Data1 = 0x00020400
    bytes = buf.to_ptr[0, 4].unpack1('L')
    assert_equal(0x00020400, bytes)
  ensure
    Fiddle.free(buf.to_ptr) if buf.respond_to?(:to_ptr)
  end

  def test_funcdesc_memid_is_at_offset_zero
    buf = TI::FUNCDESC.malloc
    buf.memid = 42
    assert_equal(42, buf.to_ptr[0, 4].unpack1('l'))
  ensure
    Fiddle.free(buf.to_ptr) if buf.respond_to?(:to_ptr)
  end

  def test_funcdesc_cparams_field_roundtrips
    buf = TI::FUNCDESC.malloc
    buf.cParams = 3
    assert_equal(3, buf.cParams)
  ensure
    Fiddle.free(buf.to_ptr) if buf.respond_to?(:to_ptr)
  end

  def test_vardesc_varkind_field_roundtrips
    buf = TI::VARDESC.malloc
    buf.varkind = 2
    assert_equal(2, buf.varkind)
  ensure
    Fiddle.free(buf.to_ptr) if buf.respond_to?(:to_ptr)
  end
end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_typeinfo.rb`
Expected: FAIL/ERROR — `win32ole/jruby/typeinfo` doesn't exist yet.

- [ ] **Step 3: Implement `typeinfo.rb`**

```ruby
# lib/win32ole/jruby/typeinfo.rb
require 'fiddle'
require 'fiddle/import'
require 'win32ole/jruby/win32'

class WIN32OLE
  module TypeInfo
    extend Fiddle::Importer

    GUID = struct([
      'unsigned long Data1', 'unsigned short Data2', 'unsigned short Data3',
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
      'long memid', 'void *lprgscode', 'void *lprgelemdescParam',
      'int funckind', 'int invkind', 'int callconv',
      'short cParams', 'short cParamsOpt', 'short oVft', 'short cScodes',
      # elemdescFunc (ELEMDESC, inlined) — same 4-field shape as ELEMDESC above
      'void *ret_tdesc_union_ptr', 'unsigned short ret_tdesc_vt',
      'void *ret_paramdescex_ptr', 'unsigned short ret_wParamFlags',
      'unsigned short wFuncFlags'
    ])

    VARDESC = struct([
      'long memid', 'void *lpstrSchema', 'void *union_oInst_or_lpvarValue',
      # elemdescVar (ELEMDESC, inlined)
      'void *tdesc_union_ptr', 'unsigned short tdesc_vt',
      'void *paramdescex_ptr', 'unsigned short wParamFlags',
      'unsigned short wVarFlags', 'int varkind'
    ])

    TYPEATTR = struct([
      'unsigned long guid_Data1', 'unsigned short guid_Data2',
      'unsigned short guid_Data3', 'unsigned char guid_Data4[8]',
      'unsigned long lcid', 'unsigned long dwReserved',
      'long memidConstructor', 'long memidDestructor',
      'void *lpstrSchema', 'unsigned long cbSizeInstance',
      'int typekind', 'unsigned short cFuncs', 'unsigned short cVars',
      'unsigned short cImplTypes', 'unsigned short cbSizeVft',
      'unsigned short cbAlignment', 'unsigned short wTypeFlags',
      'unsigned short wMajorVerNum', 'unsigned short wMinorVerNum',
      # tdescAlias (TYPEDESC, inlined) + idldescType (IDLDESC: void* placeholder + DWORD + WORD)
      'void *tdescAlias_union_ptr', 'unsigned short tdescAlias_vt',
      'unsigned long idldescType_dwReserved', 'unsigned short idldescType_wIDLFlags'
    ])
  end
end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_typeinfo.rb`
Expected: all tests PASS. If any size assertion fails, follow the cross-check instruction above (Microsoft's documented layouts) before changing anything — do not adjust the assertion to match an unverified number.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/typeinfo.rb test/win32ole/jruby/test_typeinfo.rb
git commit -m "jruby: Fiddle::CStructBuilder declarations for TYPEATTR/FUNCDESC/VARDESC/ELEMDESC/TYPEDESC"
```

---

## Task 2: `typeinfo.rb` — vtable slots, enum tables, error message helper

**Files:**
- Modify: `lib/win32ole/jruby/typeinfo.rb`
- Modify: `test/win32ole/jruby/test_typeinfo.rb`

**Interfaces:**
- Consumes: nothing new.
- Produces: `WIN32OLE::TypeInfo::ITYPEINFO_VTBL`, `::ITYPELIB_VTBL` (slot-name→index hashes), `TYPEKIND_NAMES`, `VARKIND_NAMES`, `.invoke_kind_name(bitmask)`, `Win32.query_interface_error_message(operation, detail)`.

Still pure logic — no native calls, fully testable on this machine right now.

- [ ] **Step 1: Write the failing tests**

Add to `test/win32ole/jruby/test_typeinfo.rb` (inside the existing `TestTypeInfo` class):

```ruby
  def test_itypeinfo_vtbl_slots
    assert_equal(3, TI::ITYPEINFO_VTBL[:GetTypeAttr])
    assert_equal(5, TI::ITYPEINFO_VTBL[:GetFuncDesc])
    assert_equal(6, TI::ITYPEINFO_VTBL[:GetVarDesc])
    assert_equal(12, TI::ITYPEINFO_VTBL[:GetDocumentation])
    assert_equal(14, TI::ITYPEINFO_VTBL[:GetRefTypeInfo])
    assert_equal(18, TI::ITYPEINFO_VTBL[:GetContainingTypeLib])
    assert_equal(19, TI::ITYPEINFO_VTBL[:ReleaseTypeAttr])
    assert_equal(20, TI::ITYPEINFO_VTBL[:ReleaseFuncDesc])
    assert_equal(21, TI::ITYPEINFO_VTBL[:ReleaseVarDesc])
  end

  def test_itypelib_vtbl_slots
    assert_equal(3, TI::ITYPELIB_VTBL[:GetTypeInfoCount])
    assert_equal(4, TI::ITYPELIB_VTBL[:GetTypeInfo])
    assert_equal(7, TI::ITYPELIB_VTBL[:GetLibAttr])
    assert_equal(9, TI::ITYPELIB_VTBL[:GetDocumentation])
    assert_equal(12, TI::ITYPELIB_VTBL[:ReleaseTLibAttr])
  end

  def test_typekind_names
    assert_equal('Enum', TI::TYPEKIND_NAMES[0])
    assert_equal('Dispatch', TI::TYPEKIND_NAMES[4])
    assert_equal('Max', TI::TYPEKIND_NAMES[8])
    assert_nil(TI::TYPEKIND_NAMES[99])
  end

  def test_varkind_names
    assert_equal('PERINSTANCE', TI::VARKIND_NAMES[0])
    assert_equal('CONSTANT', TI::VARKIND_NAMES[2])
    assert_nil(TI::VARKIND_NAMES[99])
  end

  def test_invoke_kind_name_property_when_get_and_put_both_set
    assert_equal('PROPERTY', TI.invoke_kind_name(0x2 | 0x4))
  end

  def test_invoke_kind_name_propertyget_only
    assert_equal('PROPERTYGET', TI.invoke_kind_name(0x2))
  end

  def test_invoke_kind_name_propertyput_only
    assert_equal('PROPERTYPUT', TI.invoke_kind_name(0x4))
  end

  def test_invoke_kind_name_propertyputref_only
    assert_equal('PROPERTYPUTREF', TI.invoke_kind_name(0x8))
  end

  def test_invoke_kind_name_func_only
    assert_equal('FUNC', TI.invoke_kind_name(0x1))
  end

  def test_invoke_kind_name_unknown_bitmask
    assert_equal('UNKNOWN', TI.invoke_kind_name(0))
  end

  def test_query_interface_error_message_matches_method_error_message_shape
    msg = WIN32OLE::Win32.query_interface_error_message('GetTypeInfo', 'boom')
    assert_match(/\Afailed to GetTypeInfo: boom\z/, msg)
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_typeinfo.rb`
Expected: the 10 new tests FAIL with `NameError`/`NoMethodError` (constants/method not defined yet); the Task 1 tests still PASS.

- [ ] **Step 3: Implement**

Append inside `module TypeInfo` in `lib/win32ole/jruby/typeinfo.rb` (after the struct declarations), and add one method to `WIN32OLE::Win32` (in `lib/win32ole/jruby/win32.rb`, alongside the existing `method_error_message`/`property_put_error_message`):

```ruby
    # lib/win32ole/jruby/typeinfo.rb, still inside module TypeInfo

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
```

```ruby
    # lib/win32ole/jruby/win32.rb — add alongside method_error_message/property_put_error_message
    def query_interface_error_message(operation, detail)
      "failed to #{operation}: #{detail}"
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_typeinfo.rb`
Expected: all tests PASS (Task 1's 10 + this task's 10 = 20 total).

Also re-run Phase 1's pure-logic suite to confirm the `win32.rb` addition didn't break anything: `ruby -Ilib -Itest test/win32ole/jruby/test_win32.rb` — expect the existing 24/24 still passing.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/typeinfo.rb lib/win32ole/jruby/win32.rb test/win32ole/jruby/test_typeinfo.rb
git commit -m "jruby: ITypeInfo/ITypeLib vtable slots, enum-name tables, query-interface error message"
```

---

## Task 3: `typeinfo.rb` — native vtable function bindings

**Files:**
- Modify: `lib/win32ole/jruby/typeinfo.rb`

**Interfaces:**
- Consumes: `Win32.vtable_function`, `Win32.native_pointer_for` (Phase 1), `ITYPEINFO_VTBL`/`ITYPELIB_VTBL` (Task 2).
- Produces: `WIN32OLE::TypeInfo.get_type_info_fn(idispatch_ptr)`, `.type_attr_fn(itypeinfo_ptr)`, `.func_desc_fn(itypeinfo_ptr)`, `.var_desc_fn(itypeinfo_ptr)`, `.documentation_fn_for_typeinfo(itypeinfo_ptr)`, `.containing_typelib_fn(itypeinfo_ptr)`, `.release_type_attr_fn(itypeinfo_ptr)`, `.release_func_desc_fn(itypeinfo_ptr)`, `.release_var_desc_fn(itypeinfo_ptr)`, `.ref_type_info_fn(itypeinfo_ptr)`, `.type_info_count_fn(itypelib_ptr)`, `.type_info_fn(itypelib_ptr)`, `.lib_attr_fn(itypelib_ptr)`, `.documentation_fn_for_typelib(itypelib_ptr)`, `.release_tlib_attr_fn(itypelib_ptr)` — every one a memoized-per-pointer `Fiddle::Function` (memoized on the pointer value, since unlike Phase 1's per-instance `dispatch.rb` mixin, these are plain module functions operating on a raw address passed in each time, not on an object's own `@ptr`).

This is where the file gains real Windows dependencies. Like Phase 1's Task 3, nothing here is runnable on this machine — verification is CI-only (Task 11), but written now because Tasks 4-10 need it.

- [ ] **Step 1: Implement**

Append inside `module TypeInfo`:

```ruby
    W = Win32
    private_constant :W

    module_function

    def get_type_info_fn(idispatch_ptr)
      @get_type_info_fns ||= {}
      @get_type_info_fns[idispatch_ptr] ||= W.vtable_function(
        idispatch_ptr, 4, [W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def type_attr_fn(itypeinfo_ptr)
      @type_attr_fns ||= {}
      @type_attr_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetTypeAttr], [W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def func_desc_fn(itypeinfo_ptr)
      @func_desc_fns ||= {}
      @func_desc_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetFuncDesc], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def var_desc_fn(itypeinfo_ptr)
      @var_desc_fns ||= {}
      @var_desc_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetVarDesc], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def documentation_fn_for_typeinfo(itypeinfo_ptr)
      @documentation_fns ||= {}
      @documentation_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetDocumentation],
        [W::VOIDP, W::LONG, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def containing_typelib_fn(itypeinfo_ptr)
      @containing_typelib_fns ||= {}
      @containing_typelib_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetContainingTypeLib], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def release_type_attr_fn(itypeinfo_ptr)
      @release_type_attr_fns ||= {}
      @release_type_attr_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:ReleaseTypeAttr], [W::VOIDP, W::VOIDP], W::VOID
      )
    end

    def release_func_desc_fn(itypeinfo_ptr)
      @release_func_desc_fns ||= {}
      @release_func_desc_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:ReleaseFuncDesc], [W::VOIDP, W::VOIDP], W::VOID
      )
    end

    def release_var_desc_fn(itypeinfo_ptr)
      @release_var_desc_fns ||= {}
      @release_var_desc_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:ReleaseVarDesc], [W::VOIDP, W::VOIDP], W::VOID
      )
    end

    def ref_type_info_fn(itypeinfo_ptr)
      @ref_type_info_fns ||= {}
      @ref_type_info_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetRefTypeInfo], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def type_info_count_fn(itypelib_ptr)
      @type_info_count_fns ||= {}
      @type_info_count_fns[itypelib_ptr] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:GetTypeInfoCount], [W::VOIDP], W::DWORD
      )
    end

    def type_info_fn(itypelib_ptr)
      @type_info_fns ||= {}
      @type_info_fns[itypelib_ptr] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:GetTypeInfo], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def lib_attr_fn(itypelib_ptr)
      @lib_attr_fns ||= {}
      @lib_attr_fns[itypelib_ptr] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:GetLibAttr], [W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def documentation_fn_for_typelib(itypelib_ptr)
      @typelib_documentation_fns ||= {}
      @typelib_documentation_fns[itypelib_ptr] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:GetDocumentation],
        [W::VOIDP, W::LONG, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def release_tlib_attr_fn(itypelib_ptr)
      @release_tlib_attr_fns ||= {}
      @release_tlib_attr_fns[itypelib_ptr] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:ReleaseTLibAttr], [W::VOIDP, W::VOIDP], W::VOID
      )
    end
```

`GetTypeInfoCount` genuinely returns `UINT` directly (not an `HRESULT` out-param) per the real COM ABI — `W::DWORD` as the return type here is correct, not a copy-paste slip from the `HRESULT`-returning neighbors.

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/typeinfo.rb
git commit -m "jruby: ITypeInfo/ITypeLib native vtable function bindings"
```

Not independently runnable — no Windows on this machine. Re-run `ruby -Ilib -Itest test/win32ole/jruby/test_typeinfo.rb` as a smoke check that Tasks 1-2's 20 tests still pass unaffected (confirms the file is still safely requireable without touching Windows, same discipline as Phase 1).

---

## Task 4: `WIN32OLE::Param`

**Files:**
- Create: `lib/win32ole/jruby/param.rb`

**Interfaces:**
- Consumes: `WIN32OLE::TypeInfo::ELEMDESC` (Task 1), `Win32.variant_ruby_type`/unpack helpers (Phase 1).
- Produces: `WIN32OLE::Param.new(elemdesc_ptr, name)` (internal constructor — an `ELEMDESC*` pointing into a `FUNCDESC`'s `lprgelemdescParam` array, plus the parameter's name from `GetNames`), instance methods `name`, `ole_type`, `ole_type_detail`, `input?`, `output?`, `optional?`, `retval?`, `default`, `inspect`.

`Param` is pure data (spec §4.5) — built once from an already-read `ELEMDESC`, holds no COM reference of its own, needs no finalizer. This task is written to be independently readable, but like Task 3, it calls into `WIN32OLE::TypeInfo` structures whose real memory only exists behind a live `ITypeInfo*` — not independently testable without Windows; verified together with `Method` in Task 11.

The `PARAMFLAG_*` bit values (`wParamFlags` in `ELEMDESC`'s `PARAMDESC` union member) are fixed COM constants: `FIN=0x1`, `FOUT=0x2`, `FLCID=0x4`, `FRETVAL=0x8`, `FOPT=0x10`, `FHASDEFAULT=0x20`.

- [ ] **Step 1: Implement**

```ruby
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
```

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/param.rb
git commit -m "jruby: WIN32OLE::Param wrapping an ELEMDESC"
```

---

## Task 5: `WIN32OLE::Method`

**Files:**
- Create: `lib/win32ole/jruby/method.rb`

**Interfaces:**
- Consumes: `WIN32OLE::TypeInfo::FUNCDESC` (Task 1), `TypeInfo.func_desc_fn`/`release_func_desc_fn`/`documentation_fn_for_typeinfo` (Task 3), `TypeInfo.invoke_kind_name` (Task 2), `WIN32OLE::Param.new(elemdesc_ptr, name)` (Task 4), `Win32.bstr_to_s`, `Win32.native_pointer_for`.
- Produces: `WIN32OLE::Method.new(itypeinfo_ptr, index)` (internal constructor, called from `Type#ole_methods`/`#variables`-equivalent in Task 8), instance methods `name`, `return_type`, `return_vtype`, `return_type_detail`, `invoke_kind`, `invkind`, `visible?`, `helpstring`, `helpfile`, `helpcontext`, `dispid`, `offset_vtbl`, `size_params`, `size_opt_params`, `params`, `inspect`.

Read-then-release (spec §4.5): `GetFuncDesc` is called exactly once, at construction; every field this task's API needs is copied into a plain ivar before `ReleaseFuncDesc` runs. `params` is built eagerly at construction time too (one `Param` per `lprgelemdescParam[i]`), not lazily on first access — matching MRI's own eager `ole_methods_from_typeinfo` behavior (spec §8 risk #6).

`TYPEFLAG_FHIDDEN = 0x10` is the bit `visible?` checks (inverted — hidden means NOT visible), read from `FUNCDESC.wFuncFlags`.

- [ ] **Step 1: Implement**

```ruby
# lib/win32ole/jruby/method.rb
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'
require 'win32ole/jruby/param'

class WIN32OLE
  class Method
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    FUNCFLAG_FHIDDEN = 0x40 # per MEMBERID/FUNCFLAGS, distinct from TYPEFLAG's own 0x10

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
      (@func_flags & FUNCFLAG_FHIDDEN) == 0
    end

    def dispid
      @dispid
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
  end
end
```

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/method.rb
git commit -m "jruby: WIN32OLE::Method wrapping a FUNCDESC"
```

Not independently testable (needs a live `ITypeInfo*`) — verified in Task 11.

---

## Task 6: `WIN32OLE::Variable`

**Files:**
- Create: `lib/win32ole/jruby/variable.rb`

**Interfaces:**
- Consumes: `WIN32OLE::TypeInfo::VARDESC` (Task 1), `TypeInfo.var_desc_fn`/`release_var_desc_fn`/`documentation_fn_for_typeinfo`/`VARKIND_NAMES` (Tasks 2-3), `Win32.variant_ruby_type`/unpack helpers, `Win32.native_pointer_for`.
- Produces: `WIN32OLE::Variable.new(itypeinfo_ptr, index)` (internal constructor), instance methods `name`, `ole_type`, `ole_type_detail`, `value`, `visible?`, `variable_kind`, `varkind`, `inspect`.

Read-then-release, same discipline as `Method` (Task 5) — one `GetVarDesc`/`ReleaseVarDesc` pair per instance, at construction. `visible?` mirrors `Method`'s own fixed pattern (Task 5's review found and fixed a real bug there): real MRI (`win32ole_variable.c`) checks three `VARFLAGS` bits, not one — `VARFLAG_FRESTRICTED = 0x80`, `VARFLAG_FHIDDEN = 0x40`, `VARFLAG_FNONBROWSABLE = 0x400`. Note these are DIFFERENT bit values than `Method`'s `FUNCFLAG_FRESTRICTED`/`FUNCFLAG_FNONBROWSABLE` (0x1/0x400 there) even though `FHIDDEN` happens to be `0x40` in both enums — `VARFLAGS` and `FUNCFLAGS` are parallel but not identical enums; do not assume a bit value carries over from one to the other without checking.

- [ ] **Step 1: Implement**

```ruby
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
      W.variant_ruby_type(@vt).to_s.upcase
    rescue NotImplementedError
      "VT_#{@vt}"
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
      return nil if variant_ptr.zero?

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
      @@get_names_fns[itypeinfo_ptr] ||= W.vtable_function(
        itypeinfo_ptr, TI::ITYPEINFO_VTBL[:GetNames],
        [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end
  end
end
```

**Reviewer note:** `get_names_fn` is duplicated between `Method` (Task 5) and `Variable` (this task) with the same body. Leave the duplication for this task — do not extract a shared helper mid-task; if the task reviewer flags it, the fix belongs in `typeinfo.rb` (e.g. `TypeInfo.get_names_fn(itypeinfo_ptr)`, memoized there once instead of once per class) as a small follow-up, not a blocker for this task's own review.

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/variable.rb
git commit -m "jruby: WIN32OLE::Variable wrapping a VARDESC"
```

---

## Task 7: `WIN32OLE::TypeLib`

**Files:**
- Create: `lib/win32ole/jruby/typelib.rb`

**Interfaces:**
- Consumes: `TypeInfo.lib_attr_fn`/`release_tlib_attr_fn`/`documentation_fn_for_typelib`/`type_info_count_fn`/`type_info_fn` (Task 3), `WIN32OLE::TypeInfo::GUID` (Task 1), `Win32.native_pointer_for`, the Phase-1 `install_finalizer`-equivalent pattern.
- Produces: `WIN32OLE::TypeLib.new(itypelib_ptr)` (internal constructor — called from `Type#ole_typelib` in Task 8), instance methods `guid`, `name`, `version`, `major_version`, `minor_version`, `path`, `visible?`, `library_name`, `ole_types`, `inspect`.

`TypeLib` holds a live, refcounted `ITypeLib*` — same finalizer discipline as Phase 1's `WIN32OLE` (capture only the raw pointer + a memoized `Release` `Fiddle::Function`, never `self`). `TLIBATTR` (from `GetLibAttr`) is smaller and flatter than `TYPEATTR`/`FUNCDESC`/`VARDESC` — `{ GUID guid; LCID lcid; SYSKIND syskind; WORD wMajorVerNum; WORD wMinorVerNum; WORD wLibFlags; }` — declared inline in this file rather than added to `typeinfo.rb`, since nothing else needs it.

`LIBFLAG_FHIDDEN = 0x1` gates `visible?`. `path` implements the spec's §5/§8-risk-#4 single-key registry read: `HKEY_CLASSES_ROOT\TypeLib\{guid}\{major}.{minor}\{lcid}\win32` (falling back to `win64`/`win16` per the real MRI lookup order in `reg_get_typelib_file_path`) via `RegOpenKeyExW`/`RegQueryValueExW`/`RegCloseKey` — three new, narrowly-scoped registry bindings in `typeinfo.rb`, added by this task since nothing else needs them. If this proves more involved during implementation than this paragraph assumes, `path` degrades to raising `NotImplementedError` rather than blocking the rest of the class (spec §8 risk #4) — the task reviewer should accept either a working `path` or a clearly-`NotImplementedError`'d one, but not a silently-wrong one.

- [ ] **Step 1: Add the three registry bindings to `typeinfo.rb`**

```ruby
    # lib/win32ole/jruby/typeinfo.rb — inside module TypeInfo, alongside the other *_fn methods
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
```

(`W` here is `Win32`, already aliased earlier in this file per Task 3.)

- [ ] **Step 2: Implement `typelib.rb`**

```ruby
# lib/win32ole/jruby/typelib.rb
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'
require 'win32ole/jruby/type'

class WIN32OLE
  class TypeLib
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    LIBFLAG_FHIDDEN = 0x1

    def initialize(itypelib_ptr)
      @ptr = itypelib_ptr

      lib_attr_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.lib_attr_fn(@ptr).call(@ptr, lib_attr_out)
      if W.failed?(hr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('GetLibAttr', W.hr_hex(hr))
      end
      attr_ptr = lib_attr_out.unpack1(W::PACK_PTR)
      p = Fiddle::Pointer.new(attr_ptr)
      @guid_bytes = p[0, 16]
      @lcid = p[16, 4].unpack1('L')
      @major = p[16 + 4 + 4, 2].unpack1('S') # skip lcid(4) + syskind(4, enum-sized)
      @minor = p[16 + 4 + 4 + 2, 2].unpack1('S')
      @lib_flags = p[16 + 4 + 4 + 2 + 2, 2].unpack1('S')
      TI.release_tlib_attr_fn(@ptr).call(@ptr, attr_ptr)

      @name, @helpstring, @help_context, @helpfile = read_documentation(@ptr, -1)

      install_finalizer
    end

    def guid
      d1, d2, d3 = @guid_bytes.unpack('LSS')
      d4 = @guid_bytes[8, 8].unpack('C8')
      format('{%08X-%04X-%04X-%02X%02X-%02X%02X%02X%02X%02X%02X}', d1, d2, d3, *d4)
    end

    def name
      @name
    end

    def version
      "#{@major}.#{@minor}"
    end

    def major_version
      @major
    end

    def minor_version
      @minor
    end

    def visible?
      (@lib_flags & LIBFLAG_FHIDDEN) == 0
    end

    def library_name
      @name
    end

    def path
      reg_path_lookup(guid, version, @lcid)
    end

    def ole_types
      count = TI.type_info_count_fn(@ptr).call(@ptr)
      Array.new(count) do |i|
        ti_out = ("\x00" * W::PTR_SIZE).b
        hr = TI.type_info_fn(@ptr).call(@ptr, i, ti_out)
        next nil if W.failed?(hr)

        WIN32OLE::Type.new(ti_out.unpack1(W::PACK_PTR))
      end.compact
    end

    def inspect
      "#<WIN32OLE::TypeLib:#{name}>"
    end

    private

    def read_documentation(itypelib_ptr, index)
      name_out = ("\x00" * W::PTR_SIZE).b
      docstring_out = ("\x00" * W::PTR_SIZE).b
      helpcontext_out = ("\x00" * 4).b
      helpfile_out = ("\x00" * W::PTR_SIZE).b
      TI.documentation_fn_for_typelib(itypelib_ptr).call(
        itypelib_ptr, index, name_out, docstring_out, helpcontext_out, helpfile_out
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

    def reg_path_lookup(guid_str, version_str, lcid)
      # HKEY_CLASSES_ROOT is itself the merged Classes root, so the subkey
      # path must NOT be prefixed with "SOFTWARE\Classes\" (that prefix is
      # only needed when opening under HKEY_LOCAL_MACHINE/HKEY_CURRENT_USER
      # directly, as MRI's own clsid_from_remote does for the DCOM path).
      key_path = W.wstr("TypeLib\\#{guid_str}\\#{version_str}\\#{lcid}\\win32")
      hkey_out = ("\x00" * W::PTR_SIZE).b
      err = TI.reg_open_key_ex.call(TI::HKEY_CLASSES_ROOT, key_path, 0, TI::KEY_READ, hkey_out)
      return nil unless err.zero?

      hkey = hkey_out.unpack1(W::PACK_PTR)
      begin
        empty_name = W.wstr('')
        type_out = ("\x00" * 4).b
        size_out = [520].pack('L') # 260 WCHARs, generous for a MAX_PATH-style value
        data_out = ("\x00" * 520).b
        err = TI.reg_query_value_ex.call(hkey, empty_name, nil, type_out, data_out, size_out)
        return nil unless err.zero? && type_out.unpack1('L') == TI::REG_SZ

        data_out[0, size_out.unpack1('L')].force_encoding('UTF-16LE').encode('UTF-8').delete("\x00")
      ensure
        TI.reg_close_key.call(hkey)
      end
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
```

`TLIBATTR`'s field layout used above (`guid`(16) + `lcid`(4) + `syskind`(4, it's an enum/`int`) + `wMajorVerNum`(2) + `wMinorVerNum`(2) + `wLibFlags`(2)) follows the same alignment reasoning as Task 1's structs — cross-check against Microsoft's `TLIBATTR` documentation during this task's own review, the same discipline Task 1 established, since this is a new struct not covered by Task 1's own tests.

- [ ] **Step 3: Commit**

```bash
git add lib/win32ole/jruby/typeinfo.rb lib/win32ole/jruby/typelib.rb
git commit -m "jruby: WIN32OLE::TypeLib wrapping an ITypeLib*, including TypeLib#path registry lookup"
```

---

## Task 8: `WIN32OLE::Type`

**Files:**
- Create: `lib/win32ole/jruby/type.rb`

**Interfaces:**
- Consumes: `WIN32OLE::TypeInfo::TYPEATTR` (Task 1), `TypeInfo.type_attr_fn`/`release_type_attr_fn`/`documentation_fn_for_typeinfo`/`containing_typelib_fn`/`ref_type_info_fn`/`TYPEKIND_NAMES` (Tasks 2-3), `WIN32OLE::Method.new(itypeinfo_ptr, index)` (Task 5), `WIN32OLE::Variable.new(itypeinfo_ptr, index)` (Task 6), `WIN32OLE::TypeLib.new(itypelib_ptr)` (Task 7, circular require — see note below).
- Produces: `WIN32OLE::Type.new(itypeinfo_ptr)` (internal constructor, called from `WIN32OLE#ole_type`/`#ole_methods` in Task 9 and `TypeLib#ole_types` in Task 7), instance methods `name`, `ole_type`, `guid`, `progid`, `visible?`, `major_version`, `minor_version`, `typekind`, `helpstring`, `src_type`, `helpfile`, `helpcontext`, `variables`, `ole_methods`, `ole_typelib`, `inspect`.

`Type` holds a live `ITypeInfo*` — same finalizer discipline as `TypeLib`. `visible?` is gated by TWO bits, not one: `TYPEFLAG_FHIDDEN = 0x10` AND `TYPEFLAG_FRESTRICTED = 0x200` (confirmed directly against `ext/win32ole/win32ole_type.c`'s `ole_type_visible`, which checks `wTypeFlags & (TYPEFLAG_FHIDDEN | TYPEFLAG_FRESTRICTED)`) — this is the same recurring pattern seen in every other Phase 2 class so far (`Method`, `Variable`, `TypeLib` all needed more than one bit checked; do not assume a single-bit check is complete here either). `TYPEFLAG_FHIDDEN`'s bit value (0x10) happens to numerically coincide with `Method`'s unrelated `FUNCFLAG_FHIDDEN` bit value — a coincidence of bit position, not a shared constant or shared meaning.

**Circular require note:** `type.rb` and `typelib.rb` reference each other's classes (`Type#ole_typelib` builds a `TypeLib`; `TypeLib#ole_types` builds `Type`s). Ruby's `require` handles this the same way Phase 1's `win32ole.rb`/`dispatch.rb` mutual references were handled: both files `require` each other, and since method bodies (not load-time code) are what reference the other class, whichever file loads first only needs the other class to exist by the time its methods are actually *called*, not by the time it's *defined*. `lib/win32ole/jruby.rb` (Phase 1's entry point) must `require` both, in either order — add both requires there in this task.

- [ ] **Step 1: Implement**

```ruby
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
      @alias_union_ptr = typeattr.tdescAlias_union_ptr
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
```

**`type_attr_func_count`/`type_attr_var_count` call `GetTypeAttr`/`ReleaseTypeAttr` a second and third time** (the constructor already called it once for `@typekind`/`guid`/etc.). This is a deliberate simplicity-over-micro-optimization choice consistent with read-then-release's own philosophy (§4.5) — `cFuncs`/`cVars` aren't read eagerly in the constructor because `ole_methods`/`variables` might never be called on a given `Type` instance, and a second `GetTypeAttr` round-trip is cheap relative to the `GetFuncDesc`/`GetVarDesc` calls that follow it. If the task reviewer considers this wasteful enough to fix, storing `@cfuncs`/`@cvars` from the constructor's own `typeattr` read (before it's released) is the straightforward fix — flag it as a finding rather than silently changing scope.

**`W.prog_id_from_clsid`** is a new Phase-1-substrate addition this task needs (`ProgIDFromCLSID`, `oleaut32.dll`) — add it to `lib/win32ole/jruby/win32.rb` alongside the other lazy native bindings:

```ruby
    # lib/win32ole/jruby/win32.rb — add alongside co_create_instance etc.
    def prog_id_from_clsid_fn
      @prog_id_from_clsid_fn ||= Fiddle::Function.new(oleaut32['ProgIDFromCLSID'], [VOIDP, VOIDP], LONG, STDCALL)
    end

    def prog_id_from_clsid(clsid_bytes)
      out = ("\x00" * PTR_SIZE).b
      hr = prog_id_from_clsid_fn.call(clsid_bytes, out)
      return nil if failed?(hr)

      bstr = out.unpack1(PACK_PTR)
      str = bstr_to_s(bstr)
      sys_free_string.call(bstr) unless bstr.zero?
      str
    end
```

- [ ] **Step 2: Wire both new files into the entry point**

```ruby
# lib/win32ole/jruby.rb — add after the existing win32ole require
require 'win32ole/jruby/type'
require 'win32ole/jruby/typelib'
```

- [ ] **Step 3: Commit**

```bash
git add lib/win32ole/jruby/type.rb lib/win32ole/jruby/win32.rb lib/win32ole/jruby.rb
git commit -m "jruby: WIN32OLE::Type wrapping an ITypeInfo*, ProgIDFromCLSID binding"
```

---

## Task 9: `WIN32OLE` instance additions

**Files:**
- Modify: `lib/win32ole/jruby/win32ole.rb`

**Interfaces:**
- Consumes: `WIN32OLE::Type.new(itypeinfo_ptr)` (Task 8), `TypeInfo.get_type_info_fn` (Task 3), `dispid_for` (Phase 1's `Dispatch` mixin).
- Produces: `WIN32OLE#ole_type`, `#ole_methods`, `#ole_get_methods`, `#ole_put_methods`, `#ole_func_methods`, `#ole_typelib`, `#ole_respond_to?`.

`ole_type`/`ole_typelib` use the direct `GetTypeInfo`→wrap path (spec §1.1's `fole_type` pattern — no name round-trip). `ole_methods`/`ole_get_methods`/`ole_put_methods`/`ole_func_methods` use the name-round-trip-via-typelib path (spec §1.1's `typeinfo_from_ole`/§8 risk #3 — ported as-is, not simplified, per that risk's own stated resolution) and filter `Type#ole_methods`'s full member list by `invkind` bitmask.

- [ ] **Step 1: Implement**

Append inside `class WIN32OLE` (after Task 6's dispatch/marshaling methods from Phase 1):

```ruby
  def ole_type
    type_info_ptr = get_type_info_ptr
    raise WIN32OLE::QueryInterfaceError, 'failed to GetTypeInfo' if type_info_ptr.nil?

    Type.new(type_info_ptr)
  end

  def ole_typelib
    ole_type.ole_typelib
  end

  def ole_methods
    ole_methods_by_invkind(nil)
  end

  def ole_get_methods
    ole_methods_by_invkind(TypeInfo::INVOKE_PROPERTYGET)
  end

  def ole_put_methods
    ole_methods_by_invkind(TypeInfo::INVOKE_PROPERTYPUT | TypeInfo::INVOKE_PROPERTYPUTREF)
  end

  def ole_func_methods
    ole_methods_by_invkind(TypeInfo::INVOKE_FUNC)
  end

  def ole_respond_to?(name)
    !dispid_for(name.to_s).nil?
  end

  private

  def get_type_info_ptr
    out = ("\x00" * W::PTR_SIZE).b
    hr = TypeInfo.get_type_info_fn(@ptr).call(@ptr, 0, W::LOCALE_SYSTEM_DEFAULT, out)
    return nil if W.failed?(hr)

    out.unpack1(W::PACK_PTR)
  end

  def ole_methods_by_invkind(mask)
    type_via_containing_typelib.ole_methods.select do |m|
      mask.nil? || (m.invkind & mask) != 0
    end
  end

  # Ports ext/win32ole/win32ole.c's typeinfo_from_ole: GetTypeInfo →
  # GetDocumentation (this type's own name) → GetContainingTypeLib →
  # scan the typelib for the entry with a matching name → GetTypeInfo(i)
  # again. See design spec §1.1/§8 risk #3 for why this round-trip exists
  # instead of just reusing the first ITypeInfo* the way #ole_type does —
  # ported as-is rather than "simplified" without understanding it.
  def type_via_containing_typelib
    first_type = ole_type
    target_name = first_type.name
    tlib = first_type.ole_typelib
    match = tlib.ole_types.find { |t| t.name == target_name }
    match || first_type
  end
```

Add `LOCALE_SYSTEM_DEFAULT = 0x0800` to `lib/win32ole/jruby/win32.rb`'s constants (alongside `CLSCTX_INPROC_SERVER` etc.) — this is the real Win32 constant value (`MAKELCID(MAKELANGID(LANG_NEUTRAL, SUBLANG_SYS_DEFAULT), SORT_DEFAULT)`).

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/win32ole.rb lib/win32ole/jruby/win32.rb
git commit -m "jruby: WIN32OLE#ole_type/ole_methods/ole_typelib/ole_respond_to? entry points"
```

---

## Task 10: Explicit `NotImplementedError`s for out-of-scope members

**Files:**
- Modify: `lib/win32ole/jruby/type.rb`
- Modify: `lib/win32ole/jruby/method.rb`
- Modify: `lib/win32ole/jruby/typelib.rb`
- Modify: `lib/win32ole/jruby/win32ole.rb`

**Interfaces:** none new — this task only adds raising stubs, per Phase 1's own established convention (design spec §3, matching Phase 1's own "raise NotImplementedError/TypeError for out-of-scope members, don't silently omit them or let them fall through to NoMethodError").

- [ ] **Step 1: Add to `Type`**

```ruby
  # lib/win32ole/jruby/type.rb — inside class Type
  def implemented_ole_types
    raise NotImplementedError, 'ImplType traversal is not implemented yet (Phase 2 non-goal)'
  end

  def source_ole_types
    raise NotImplementedError, 'ImplType traversal is not implemented yet (Phase 2 non-goal)'
  end

  def default_event_sources
    raise NotImplementedError, 'ImplType traversal is not implemented yet (Phase 2 non-goal)'
  end

  def default_ole_types
    raise NotImplementedError, 'ImplType traversal is not implemented yet (Phase 2 non-goal)'
  end

  def self.ole_classes(typelib)
    raise NotImplementedError, 'registry enumeration is not implemented yet (Phase 2 non-goal)'
  end

  def self.typelibs
    raise NotImplementedError, 'registry enumeration is not implemented yet (Phase 2 non-goal)'
  end

  def self.progids
    raise NotImplementedError, 'registry enumeration is not implemented yet (Phase 2 non-goal)'
  end
```

- [ ] **Step 2: Add to `Method`**

```ruby
  # lib/win32ole/jruby/method.rb — inside class Method
  def event?
    raise NotImplementedError, 'ImplType traversal is not implemented yet (Phase 2 non-goal)'
  end

  def event_interface
    raise NotImplementedError, 'ImplType traversal is not implemented yet (Phase 2 non-goal)'
  end
```

- [ ] **Step 3: Add to `TypeLib`**

```ruby
  # lib/win32ole/jruby/typelib.rb — inside class TypeLib
  def self.typelibs
    raise NotImplementedError, 'registry enumeration is not implemented yet (Phase 2 non-goal)'
  end
```

Do NOT add a `TypeLib.new(name, version)` public class-level constructor stub — there is no such method to stub, since this design never defines one (spec §3); only the internal `TypeLib.new(itypelib_ptr)` from Task 7 exists, and it is not part of the public contract this task needs to guard.

- [ ] **Step 4: Add to `WIN32OLE`**

```ruby
  # lib/win32ole/jruby/win32ole.rb — inside class WIN32OLE
  def ole_method_help(*)
    raise NotImplementedError, 'launching help files is not implemented (Phase 2 non-goal)'
  end

  def ole_obj_help
    raise NotImplementedError, 'launching help files is not implemented (Phase 2 non-goal)'
  end

  def ole_query_interface(*)
    raise NotImplementedError, 'arbitrary QueryInterface is not implemented (Phase 2 non-goal)'
  end
```

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/type.rb lib/win32ole/jruby/method.rb lib/win32ole/jruby/typelib.rb lib/win32ole/jruby/win32ole.rb
git commit -m "jruby: explicit NotImplementedError for Phase 2 non-goals (registry enumeration, ImplType traversal, help/QI)"
```

---

## Task 11: CI verification against the existing legacy test suite

**Files:**
- None created — this task pushes Tasks 1-10's work to CI and fixes whatever real bugs the existing test suite finds, the same disciplined loop Phase 1 used repeatedly.
- Create (only if the `GC.stress` gap below needs it): `test/win32ole/jruby/test_typelib_gc_stress.rb`

**Interfaces:** none new.

Unlike Phase 1, this phase's target test files already exist and already test almost exactly this API surface: `test/win32ole/test_win32ole_type.rb` (199 lines), `test_win32ole_typelib.rb` (117 lines), `test_win32ole_method.rb` (134 lines), `test_win32ole_param.rb` (98 lines), `test_win32ole_variable.rb` (66 lines) — 614 lines total, already wired into the `test-jruby` CI job via `bundle exec rake` (Phase 1 Task 7). No new test file is required to exercise the core API surface; this task's job is to run them, triage failures against the spec's own scope boundaries (§3), and fix real bugs.

- [ ] **Step 1: Add the `GC.stress` finalizer test**

```ruby
# test/win32ole/jruby/test_typelib_gc_stress.rb
begin
  require 'win32ole'
rescue LoadError
end
require 'test/unit'

if defined?(WIN32OLE) && RUBY_ENGINE == 'jruby'
  class TestTypeLibGCStress < Test::Unit::TestCase
    def test_gc_stress_survives_repeated_type_and_typelib_construction
      dict = WIN32OLE.new('Scripting.Dictionary')
      GC.stress = true
      50.times do
        type = dict.ole_type
        tlib = type.ole_typelib
        assert_kind_of(WIN32OLE::Type, type)
        assert_kind_of(WIN32OLE::TypeLib, tlib)
      end
    ensure
      GC.stress = false
    end
  end
end
```

- [ ] **Step 2: Commit and push**

```bash
git add test/win32ole/jruby/test_typelib_gc_stress.rb
git commit -m "test: GC.stress probe for Type/TypeLib finalizer path"
git push origin jruby-support
```

(Confirm with whoever is driving this plan's execution before pushing — per this repo's own established practice this session, pushing to `origin/jruby-support` triggers real GitHub Actions billing and should be a deliberate, acknowledged action, not an unannounced default.)

- [ ] **Step 3: Read the CI log, triage every failure**

Run: `gh run list --branch jruby-support --limit 1`, then `gh run view <run-id> --job=<jruby-job-id> --log-failed`.

For each failure in `test_win32ole_type.rb`/`test_win32ole_typelib.rb`/`test_win32ole_method.rb`/`test_win32ole_param.rb`/`test_win32ole_variable.rb`, classify it:

- **Expected (§3 non-goal)**: the test exercises `.ole_classes`/`.typelibs`/`.progids`, `implemented_ole_types`/`source_ole_types`/`default_event_sources`/`default_ole_types`, `Method#event?`/`#event_interface`, or `TypeLib.new(name, version)`'s public registry-search constructor. No action — these correctly raise `NotImplementedError` per Task 10.
- **Real bug**: anything else. Fix it, re-run the single failing test file locally is not possible (Windows-only), so fix based on re-reading the relevant vtable signature/struct field against Microsoft's documentation and the MRI C source, commit, push, and re-check CI.

- [ ] **Step 4: Confirm zero regressions in the pre-existing Phase 1 test set**

Diff the new failure list against Phase 1's last known-good baseline (63 failures, all confirmed Phase 2+/out-of-scope) the same way Phase 1's own CI-driven fix rounds did — any Phase-1-scoped test that was passing before this phase's changes and now fails is a real regression this phase introduced (most likely via a shared-file edit, e.g. `win32.rb`'s new `LOCALE_SYSTEM_DEFAULT` constant or `WIN32OLE::QueryInterfaceError` now actually being raised somewhere Phase 1 code didn't expect).

- [ ] **Step 5: Final commit once CI is clean**

```bash
git add -A
git commit -m "jruby: Phase 2 fixes from CI triage" # only if Step 3 produced fixes beyond Step 2's own commit
```

---

## Self-Review

**Spec coverage:** §4.1 (file layout) → Tasks 1,4-9. §4.2 (struct/vtable substrate) → Tasks 1-3. §4.3 (class responsibilities) → Tasks 4-8. §4.4 (enum tables) → Task 2. §4.5 (resource lifetime) → Tasks 4-8 (each task's own construction/release code). §5 (per-class API) → Tasks 4-9's method lists, cross-checked one-by-one against the spec's table. §6 (error translation) → `WIN32OLE::QueryInterfaceError` raised at every `GetTypeInfo`/`GetContainingTypeLib`/`GetFuncDesc`/`GetVarDesc`/`GetTypeAttr` failure site across Tasks 5,6,7,8,9. §7 (testing/CI) → Task 11. §8 risks: #1 (CStructBuilder support) → Task 1's own tests are the check; #2 (struct layout) → Task 1's cross-check instruction; #3 (name round-trip) → Task 9's `type_via_containing_typelib`, ported not simplified; #4 (`TypeLib#path`) → Task 7's degrade-to-`NotImplementedError` fallback; #5 (x86) → inherited, not re-litigated; #6 (performance) → Task 5/8's eager-construction note.

**Placeholder scan:** no TBD/TODO/"implement later" in any task. An earlier draft of Task 5's `Method` constructor contained dead code and an invalid `.offsetof` call for computing `Param` array addresses — both were corrected in place (the constructor now reads `lprgelemdescParam`'s pointer *value* directly and indexes `elemdesc_array_ptr + i * TI::ELEMDESC.size`) rather than being left as flagged-but-broken code, consistent with Phase 1's own lesson (its plan initially showed a similarly "wrong then corrected" `FormatMessage` snippet and was fixed to show only the final correct code directly).

**Type consistency:** `Param.new(elemdesc_ptr, name)` (Task 4) matches its call site in `Method` (Task 5: `WIN32OLE::Param.new(elemdesc_ptr, param_names[i])`). `Method.new(itypeinfo_ptr, index)`/`Variable.new(itypeinfo_ptr, index)` match their call sites in `Type#ole_methods`/`#variables` (Task 8). `TypeLib.new(itypelib_ptr)` matches its call sites in `Type#ole_typelib` (Task 8) and `WIN32OLE#ole_typelib` (via `Type`, Task 9). `Type.new(itypeinfo_ptr)` matches its call sites in `WIN32OLE#ole_type` (Task 9), `TypeLib#ole_types` (Task 7), and `Type#src_type`'s own recursive call (Task 8).
