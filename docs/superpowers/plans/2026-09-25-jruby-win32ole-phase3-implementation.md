# JRuby win32ole Phase 3 (Record / Variant) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement `WIN32OLE::Record`, `WIN32OLE::Variant`/`WIN32OLE::VariantType`, and `VT_ARRAY` (`SAFEARRAY`)/`VT_RECORD` marshaling for ordinary dispatch, on top of the already-shipped Phase 1 (`WIN32OLE` core dispatch) and Phase 2 (`WIN32OLE::Type`/`TypeLib`/`Method`/`Param`/`Variable` introspection).

**Architecture:** Three new substrate files (`array.rb`: SAFEARRAY Fiddle bindings + N-dimensional index-walk logic; `record.rb`: `IRecordInfo` vtable + `WIN32OLE::Record`; `variant.rb`: `WIN32OLE::Variant`/`VariantType`) plus two extended files (`win32.rb`: wider VARIANT substrate — generalized pack/unpack, full scalar family, `VT_BYREF` construction; `win32ole.rb`: dispatch hook gains `Array`/`Record`/`Variant` argument and `VT_ARRAY`/`VT_RECORD` result branches). A small Tidy-First refactor (Task 3) promotes Phase 1's private per-instance marshal methods to `WIN32OLE` class methods first, since `array.rb`/`record.rb`/`variant.rb` all need to recursively marshal arbitrary Ruby values (a `SAFEARRAY` of `VT_VARIANT` elements, a `Record`'s field values) without being a live `WIN32OLE` instance themselves.

**Tech Stack:** Ruby stdlib `fiddle`, plain `Array#pack`/`String#unpack1` (no `Fiddle::Importer.struct` — `SAFEARRAY`'s own layout is never hand-constructed, only read via `oleaut32` calls; see Task 4).

**Spec:** `docs/superpowers/specs/2026-09-24-jruby-win32ole-phase3-record-variant-design.md` — this plan implements §2/§4/§5/§6 in full; §3's non-goals (`VT_CY`/`VT_DATE`, registry enumeration, unbound `IRecordInfo` members) are out of scope here, and §8's six open risks are addressed by name in the relevant tasks, not assumed away.

**A materially important finding not in the spec, discovered while writing this plan:** the design's own §8 risk #1 calls for validating the hand-transcribed `IRecordInfo` vtable slots against a live Windows build "before anything else depends on it." This project's actual CI (`.github/workflows/windows.yml`) runs on a bare `windows-latest` runner with no step that builds or registers `RbComTest.ComSrvTest` — the custom VB.NET COM server the legacy `test_win32ole_record.rb`/most of `test_win32ole_variant.rb`'s struct-passing tests require (those tests already self-`omit` when it's absent — confirmed by reading `test/win32ole/test_win32ole_record.rb:79-86`). There is no standard, always-available Windows Automation object exposing a `VT_RECORD`-typed member the way `Scripting.Dictionary` exposes ordinary scalars/`SAFEARRAY`s. **Decision (confirmed with the person driving this plan): proceed with `Record` implemented but only structurally/unit-tested — §8 risk #1 is explicitly *not* resolved by this phase and is carried forward, not silently assumed fixed.** `WIN32OLE::Variant`'s `SAFEARRAY` work does not have this problem: `Scripting.Dictionary#Keys`/`#Items` return a real 1-D `SAFEARRAY` of `VARIANT`s and are already proven available in this CI (used throughout Phase 1/2), so Task 6/17 verify array marshaling against that live fixture. The same reasoning applies to design §7's live "`VT_BYREF` round-trip... through a live OLE out-parameter call" — no always-available stock COM object exposes a `ByRef`/out-parameter method either (ADO/Office, which do, are not guaranteed present — see `test/win32ole/available_ole.rb`'s own `ado_available?` gate). Task 16's `VT_BYREF` test is therefore structural (pointer-aliasing math) too, consistent with the `Record` decision above, not a live call.

## Global Constraints

- **Reuse Phase 1/2's substrate unchanged**: `Win32.vtable_function`, `Win32.native_pointer_for` (hold the returned `Fiddle::Pointer` in a local variable/array through any native call that dereferences it), `Win32.sys_free_string`, `Win32.bstr_to_s`, `WIN32OLE::RuntimeError`/`QueryInterfaceError`, `TypeInfo`'s vtable-address-keyed memoization pattern, `TypeLib#@ptr`/`Type#@ptr` (raw `ITypeLib*`/`ITypeInfo*`).
- **`Record` construction failures use `WIN32OLE::RuntimeError`**, not `QueryInterfaceError` — a real, easy-to-miss distinction from Phase 2's `ITypeInfo`-family failures (spec §4.4, §6; confirmed against `ext/win32ole/win32ole_record.c`'s own `eWIN32OLERuntimeError` use throughout).
- **`Variant.new(val, VT_RECORD)` raises `ArgumentError`** (spec §4.5/§6) — the two classes are deliberately disjoint for record types.
- **Unknown/mismatched `Record` field name → Ruby's own `KeyError`** (via `Hash#fetch`), not a `WIN32OLE`-namespaced exception (spec §6; confirmed against `ext/win32ole/win32ole_record.c:401-422`'s `olerecord_ivar_get`/`_set`).
- **`VT_CY`/`VT_DATE` are out of scope** (spec §3) — `Win32.variant_ruby_type`'s existing generic `else` branch already raises `NotImplementedError` for any VARTYPE this plan doesn't add a `case` for, so the correct action for these two is to **add no `pack_cy`/`pack_date`/`VT_CY`/`VT_DATE` case anywhere** — not a positive thing to implement, an omission to preserve. Do not port MRI's own string-fallback `default:` case for unhandled VARTYPEs (`ext/win32ole/win32ole.c`'s `ole_variant2val`) — that would silently contradict this policy.
- **`ext/win32ole/win32ole.c`, `win32ole_record.c`, `win32ole_variant.c`, `win32ole_variant_m.c` are present in this repo** (real MRI C extension source, not the JRuby port) — several tasks below cite exact line numbers from files read while writing this plan; cross-check against them directly for any detail not spelled out in a task's own code block, the same discipline Phase 2's plan established for `TYPEATTR`/`FUNCDESC`/`VARDESC`.
- **Verification reality**: pure-logic/struct-declaration work (no native calls) is locally testable on any OS/engine right now; anything calling `oleaut32`/COM APIs is Windows+COM-only and can only be verified by pushing to the `test-jruby` CI job. Per the finding above, `Record`'s live COM calls (Tasks 7-10) and `Variant`'s explicit `VT_BYREF` live round-trip (Task 16) cannot be empirically confirmed even by CI in this project's current environment — those tasks say so explicitly rather than implying a CI push will settle them.
- **New test files follow the established `RUBY_ENGINE == 'jruby'` guard convention** (`test/win32ole/jruby/test_typelib_gc_stress.rb`, `test_win32ole_phase1.rb`; `test_typeinfo.rb`/`test_win32.rb` were just fixed to match). Any file requiring a `lib/win32ole/jruby/*` file at top level must guard that require, or the Rakefile's global `test/**/test_*.rb` glob breaks every non-JRuby CI matrix job again.

---

## Task 1: `win32.rb` — wider VARIANT substrate (generalized pack/unpack + full scalar family)

**Files:**
- Modify: `lib/win32ole/jruby/win32.rb`
- Test: `test/win32ole/jruby/test_win32.rb` (inside the existing `if RUBY_ENGINE == 'jruby'` guard)

**Interfaces:**
- Consumes: nothing new — this is the foundation everything else in Phase 3 builds on.
- Produces: `Win32.pack_variant(vt, body)`/`Win32.unpack_variant(bytes, body_size: 8)` (generalized, backward compatible), `Win32.pack_i1`/`pack_ui1`/`pack_i2`/`pack_ui2`/`pack_ui4`/`pack_int`/`pack_uint`/`pack_ui8`/`pack_r4`/`pack_error` and their `unpack_*` counterparts, new constants `VT_NULL`/`VT_I2`/`VT_R4`/`VT_ERROR`/`VT_VARIANT`/`VT_I1`/`VT_UI1`/`VT_UI2`/`VT_UI4`/`VT_UI8`/`VT_INT`/`VT_UINT`/`VT_RECORD`/`VT_ARRAY`/`VT_BYREF`/`VT_TYPEMASK`.

Pure logic, no native calls — runs on this machine right now, any engine, exactly like Phase 1/2's own pure-logic layers.

- [ ] **Step 1: Write the failing test**

Add inside `test/win32ole/jruby/test_win32.rb`'s existing `class TestWin32 < Test::Unit::TestCase` body (which is itself inside the file's `if RUBY_ENGINE == 'jruby'` guard):

```ruby
  def test_pack_variant_accepts_a_variable_length_body
    bytes = W.pack_variant(W::VT_RECORD, "\x01\x02")
    assert_equal(W::VARIANT_SIZE, bytes.bytesize)
  end

  def test_pack_variant_zero_pads_a_short_body
    bytes = W.pack_variant(W::VT_I4, W.pack_i4(7))
    assert_equal("\x00".b * (W::VARIANT_SIZE - 12), bytes[12, W::VARIANT_SIZE - 12])
  end

  def test_pack_variant_rejects_a_body_longer_than_the_slot
    assert_raise(ArgumentError) { W.pack_variant(W::VT_I4, "\x00".b * (W::VARIANT_SIZE - 7)) }
  end

  def test_unpack_variant_reads_the_requested_body_size
    body = [1, 2].pack(W::PACK_PTR * 2)
    bytes = W.pack_variant(W::VT_RECORD, body)
    vt, read_body = W.unpack_variant(bytes, body_size: body.bytesize)
    assert_equal(W::VT_RECORD, vt)
    assert_equal(body, read_body)
  end

  def test_unpack_variant_default_body_size_is_unchanged
    bytes = W.pack_variant(W::VT_I4, W.pack_i4(42))
    vt, payload = W.unpack_variant(bytes)
    assert_equal(W::VT_I4, vt)
    assert_equal(42, W.unpack_i4(payload))
  end

  def test_scalar_family_round_trips
    assert_equal(-5, W.unpack_i1(W.pack_i1(-5)))
    assert_equal(200, W.unpack_ui1(W.pack_ui1(200)))
    assert_equal(-1000, W.unpack_i2(W.pack_i2(-1000)))
    assert_equal(40_000, W.unpack_ui2(W.pack_ui2(40_000)))
    assert_equal(4_000_000_000, W.unpack_ui4(W.pack_ui4(4_000_000_000)))
    assert_equal(-7, W.unpack_int(W.pack_int(-7)))
    assert_equal(7, W.unpack_uint(W.pack_uint(7)))
    assert_equal(18_000_000_000, W.unpack_ui8(W.pack_ui8(18_000_000_000)))
    assert_in_delta(1.5, W.unpack_r4(W.pack_r4(1.5)), 0.0001)
    assert_equal(-2_147_024_809, W.unpack_error(W.pack_error(-2_147_024_809)))
  end

  def test_scalar_family_payloads_are_eight_bytes_zero_padded
    assert_equal(8, W.pack_i1(1).bytesize)
    assert_equal(8, W.pack_ui8(1).bytesize) # the one already-8-byte-wide case: no padding needed
    assert_equal("\x00".b * 7, W.pack_i1(1)[1, 7])
  end

  def test_new_vt_constants_do_not_collide_with_existing_ones
    existing = [W::VT_EMPTY, W::VT_I4, W::VT_R8, W::VT_BSTR, W::VT_DISPATCH, W::VT_BOOL, W::VT_UNKNOWN, W::VT_I8]
    new_scalars = [W::VT_NULL, W::VT_I2, W::VT_R4, W::VT_ERROR, W::VT_VARIANT,
                   W::VT_I1, W::VT_UI1, W::VT_UI2, W::VT_UI4, W::VT_UI8, W::VT_INT, W::VT_UINT, W::VT_RECORD]
    assert_empty(existing & new_scalars)
    assert_equal(0x2000, W::VT_ARRAY)
    assert_equal(0x4000, W::VT_BYREF)
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_win32.rb`
Expected: FAIL/ERROR — `W::VT_RECORD`, `W.pack_i1`, etc. don't exist yet; `pack_variant`/`unpack_variant` don't accept the new arguments yet.

- [ ] **Step 3: Implement**

Replace the existing `VT_*` constant block in `lib/win32ole/jruby/win32.rb`:

```ruby
    VT_EMPTY    = 0
    VT_NULL     = 1
    VT_I2       = 2
    VT_I4       = 3
    VT_R4       = 4
    VT_R8       = 5
    VT_BSTR     = 8
    VT_DISPATCH = 9
    VT_ERROR    = 10
    VT_BOOL     = 11
    VT_VARIANT  = 12
    VT_UNKNOWN  = 13
    VT_I1       = 16
    VT_UI1      = 17
    VT_UI2      = 18
    VT_UI4      = 19
    VT_I8       = 20
    VT_UI8      = 21
    VT_INT      = 22
    VT_UINT     = 23
    VT_RECORD   = 36
    VT_TYPEMASK = 0x0FFF
    VT_ARRAY    = 0x2000
    VT_BYREF    = 0x4000
```

Replace `pack_variant`/`unpack_variant`:

```ruby
    def pack_variant(vt, body)
      body = body.b
      max = VARIANT_SIZE - 8
      if body.bytesize > max
        raise ArgumentError, "body must be <= #{max} bytes, got #{body.bytesize}"
      end

      [vt, 0, 0, 0].pack('S4') + body + ("\x00".b * (max - body.bytesize))
    end

    def unpack_variant(bytes, body_size: 8)
      vt, = bytes.unpack1('S')
      [vt, bytes[8, body_size]]
    end
```

Add the scalar family next to the existing `pack_i4`/`pack_i8`/etc. block:

```ruby
    def pack_i1(value)    = [value].pack('c') + ("\x00".b * 7)
    def pack_ui1(value)   = [value].pack('C') + ("\x00".b * 7)
    def pack_i2(value)    = [value].pack('s') + ("\x00".b * 6)
    def pack_ui2(value)   = [value].pack('S') + ("\x00".b * 6)
    def pack_ui4(value)   = [value].pack('L') + ("\x00".b * 4)
    def pack_int(value)   = pack_i4(value)  # VT_INT is a plain 32-bit int, same shape as VT_I4
    def pack_uint(value)  = pack_ui4(value)
    def pack_ui8(value)   = [value].pack('Q')
    def pack_r4(value)    = [value].pack('f') + ("\x00".b * 4)
    def pack_error(value) = pack_i4(value)  # VT_ERROR is a plain LONG (SCODE)

    def unpack_i1(payload)    = payload.unpack1('c')
    def unpack_ui1(payload)   = payload.unpack1('C')
    def unpack_i2(payload)    = payload.unpack1('s')
    def unpack_ui2(payload)   = payload.unpack1('S')
    def unpack_ui4(payload)   = payload.unpack1('L')
    def unpack_int(payload)   = unpack_i4(payload)
    def unpack_uint(payload)  = unpack_ui4(payload)
    def unpack_ui8(payload)   = payload.unpack1('Q')
    def unpack_r4(payload)    = payload.unpack1('f')
    def unpack_error(payload) = unpack_i4(payload)
```

**Do not add `pack_cy`/`pack_date` or `VT_CY`/`VT_DATE` constants** — see Global Constraints.

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_win32.rb`
Expected: all tests PASS, including Phase 1/2's pre-existing ones in the same file (no regression).

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/win32.rb test/win32ole/jruby/test_win32.rb
git commit -m "jruby: generalize pack/unpack_variant, full scalar VT_* pack family, new VT_* constants"
```

---

## Task 2: `win32.rb` — `VT_BYREF` construction + `VariantChangeTypeEx` binding

**Files:**
- Modify: `lib/win32ole/jruby/win32.rb`
- Test: `test/win32ole/jruby/test_win32.rb`

**Interfaces:**
- Consumes: `Win32.pack_variant`/`unpack_variant`/`unpack_pointer`/`native_pointer_for` (Task 1/Phase 1), `VT_VARIANT`/`VT_BYREF` (Task 1).
- Produces: `Win32.pack_byref(vt, realvar_bytes)`, `Win32.variant_change_type` (memoized `Fiddle::Function`).

`pack_byref` is pure pointer arithmetic on an in-process Ruby `String` — `Fiddle::Pointer.to_ptr` works on any platform for any Ruby string, no COM/Windows API involved, so this half is locally testable. `variant_change_type` is a native `oleaut32` binding — CI-only, like Phase 2's Task 3 vtable bindings.

- [ ] **Step 1: Write the failing test**

```ruby
  def test_pack_byref_scalar_points_into_the_body_offset_of_realvar
    realvar = W.pack_variant(W::VT_I4, W.pack_i4(42))
    byref = W.pack_byref(W::VT_I4, realvar)
    vt, payload = W.unpack_variant(byref)
    assert_equal(W::VT_I4 | W::VT_BYREF, vt)
    assert_equal(W.native_pointer_for(realvar).to_i + 8, W.unpack_pointer(payload))
  end

  def test_pack_byref_variant_points_at_the_whole_realvar_buffer
    realvar = W.pack_variant(W::VT_I4, W.pack_i4(42))
    byref = W.pack_byref(W::VT_VARIANT, realvar)
    _vt, payload = W.unpack_variant(byref)
    assert_equal(W.native_pointer_for(realvar).to_i, W.unpack_pointer(payload))
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_win32.rb`
Expected: FAIL — `pack_byref` undefined.

- [ ] **Step 3: Implement**

```ruby
    # var's body holds a pointer INTO realvar's own already-packed bytes --
    # realvar must outlive var (Phase 1 §4.5's "native address embedded as
    # data" keep-alive discipline: the caller, WIN32OLE::Variant, is what
    # keeps realvar's String reachable for as long as var is in use).
    # VT_VARIANT|VT_BYREF is the one exception (mirrors MRI's
    # ole_set_byref): the pointer targets realvar's own start (offset 0,
    # the whole VARIANT), not offset 8 (one scalar slot within it).
    def pack_byref(vt, realvar_bytes)
      offset = vt == VT_VARIANT ? 0 : 8
      ptr = native_pointer_for(realvar_bytes)
      pack_variant(vt | VT_BYREF, pack_pointer((ptr + offset).to_i))
    end

    def variant_change_type
      @variant_change_type ||= Fiddle::Function.new(
        oleaut32['VariantChangeTypeEx'], [VOIDP, VOIDP, DWORD, WORD, WORD], LONG, STDCALL
      )
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_win32.rb`
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/win32.rb test/win32ole/jruby/test_win32.rb
git commit -m "jruby: VT_BYREF pointer construction + VariantChangeTypeEx binding"
```

Not fully independently runnable — `variant_change_type` itself has no Windows on this machine to call it against. Re-run the whole test file as a smoke check that nothing else broke.

---

## Task 3 (Tidy First — refactor, no new behavior): promote Phase 1's marshal methods to `WIN32OLE` class methods

**Files:**
- Modify: `lib/win32ole/jruby/win32ole.rb`

**Interfaces:**
- Consumes: nothing new.
- Produces: `WIN32OLE.ruby_value_to_variant_bytes(value, bstrs_to_free)`, `WIN32OLE.variant_bytes_to_ruby_value(bytes)`, `WIN32OLE.wrap_dispatch_pointer(ptr)` — all now callable from *outside* a `WIN32OLE` instance (needed by `array.rb`'s `VT_VARIANT`-element recursion and `record.rb`'s field marshaling and `variant.rb`'s inference path in later tasks).

**Why this is safe to do as a pure refactor:** reading the current `lib/win32ole/jruby/win32ole.rb`, `ruby_value_to_variant_bytes`/`variant_bytes_to_ruby_value` never actually touch `self` — `ruby_value_to_variant_bytes`'s `:dispatch` branch reads `value.instance_variable_get(:@ptr)` off its *argument*, not off `self`. The only real instance-dependency is `variant_bytes_to_ruby_value`'s call to `wrap_dispatch_pointer`, which itself only uses `self.class` (always plain `WIN32OLE`, no subclasses exist) — so `wrap_dispatch_pointer` becomes a class method cleanly too. This is a mechanical move, not a redesign: no call site's *behavior* changes, only *where the methods live*.

- [ ] **Step 1: Move the three methods**

In `lib/win32ole/jruby/win32ole.rb`, delete `wrap_dispatch_pointer`, `ruby_value_to_variant_bytes`, and `variant_bytes_to_ruby_value` from their current private-instance-method locations, and add this `class << self` block right after the `include Dispatch` / `W = Win32` lines near the top of `class WIN32OLE`:

```ruby
  class << self
    def wrap_dispatch_pointer(ptr)
      obj = allocate
      obj.instance_variable_set(:@ptr, ptr)
      obj.send(:install_finalizer)
      obj
    end

    def ruby_value_to_variant_bytes(value, bstrs_to_free)
      type = W.ruby_to_variant_type(value)
      payload =
        case type
        when :i4 then W.pack_i4(value)
        when :i8 then W.pack_i8(value)
        when :r8 then W.pack_r8(value)
        when :bool then W.pack_bool(value)
        when :empty then W.pack_empty
        when :bstr
          bstr = W.sys_alloc_string.call(W.wstr(value))
          bstrs_to_free << bstr
          W.pack_pointer(bstr)
        when :dispatch
          W.pack_pointer(value.instance_variable_get(:@ptr))
        end
      W.pack_variant(W::VT_FOR_TYPE.fetch(type), payload)
    end

    def variant_bytes_to_ruby_value(bytes)
      vt, payload = W.unpack_variant(bytes)
      type = W.variant_ruby_type(vt)
      case type
      when :empty then nil
      when :i4 then W.unpack_i4(payload)
      when :i8 then W.unpack_i8(payload)
      when :r8 then W.unpack_r8(payload)
      when :bool then W.unpack_bool(payload)
      when :bstr
        addr = W.unpack_pointer(payload)
        str = W.bstr_to_s(addr)
        W.sys_free_string.call(addr) unless addr.zero?
        str
      when :dispatch
        ptr = W.unpack_pointer(payload)
        ptr.zero? ? nil : wrap_dispatch_pointer(ptr)
      end
    end
  end
```

- [ ] **Step 2: Update the two call sites**

In `method_missing`, replace:

```ruby
    hr, result_bytes, excepinfo_bytes = ole_invoke(dispid, args, plan[:wflags], named_put: plan[:named_put])
    ...
    variant_bytes_to_ruby_value(result_bytes)
```

with:

```ruby
    hr, result_bytes, excepinfo_bytes = ole_invoke(dispid, args, plan[:wflags], named_put: plan[:named_put])
    ...
    self.class.variant_bytes_to_ruby_value(result_bytes)
```

`ruby_value_to_variant_bytes` itself is called from `dispatch.rb`'s `ole_invoke` (`arg_variants = arg_values.reverse.map { |v| ruby_value_to_variant_bytes(v, bstrs_to_free) }`) — since `Dispatch` is a module mixed into `WIN32OLE`, change that call site (in `lib/win32ole/jruby/dispatch.rb`) to `WIN32OLE.ruby_value_to_variant_bytes(v, bstrs_to_free)` (an explicit constant reference, since `Dispatch` itself isn't `WIN32OLE` and `self.class` inside a module method mixed into `WIN32OLE` does resolve to `WIN32OLE` — either spelling works; use the explicit `WIN32OLE.` form for clarity since `dispatch.rb` already references `WIN32OLE::RuntimeError` by explicit constant elsewhere in the same file).

- [ ] **Step 3: Run the full Phase 1/2 local + existing test suite to confirm zero behavior change**

Run: `bundle exec rake test`
Expected: identical pass/fail counts to the pre-refactor baseline (this machine is non-Windows, so this only confirms nothing broke at the require/parse/pure-logic level — full confirmation is Task 17's CI push).

- [ ] **Step 4: Commit**

```bash
git add lib/win32ole/jruby/win32ole.rb lib/win32ole/jruby/dispatch.rb
git commit -m "jruby: promote marshal helpers to WIN32OLE class methods (tidy, no behavior change)"
```

---

## Task 4: `array.rb` — `SAFEARRAY` `oleaut32` Fiddle bindings

**Files:**
- Create: `lib/win32ole/jruby/array.rb`

**Interfaces:**
- Consumes: `Win32.oleaut32`, `Win32::{VOIDP,DWORD,LONG,WORD,STDCALL}` (Phase 1).
- Produces: `WIN32OLE::SafeArray.safe_array_create`, `.safe_array_create_vector`, `.safe_array_destroy`, `.safe_array_lock`, `.safe_array_unlock`, `.safe_array_get_dim`, `.safe_array_get_lbound`, `.safe_array_get_ubound`, `.safe_array_ptr_of_index`, `.safe_array_put_element`, `.safe_array_access_data`, `.safe_array_unaccess_data` — each a memoized `Fiddle::Function`.

`SAFEARRAY`'s own byte layout (`oaidl.h`: `USHORT cDims; USHORT fFeatures; ULONG cbElements; ULONG cLocks; PVOID pvData; SAFEARRAYBOUND rgsabound[cDims]`, `SAFEARRAYBOUND: ULONG cElements; LONG lLbound`) is never hand-read or hand-built directly by this design — only via these `oleaut32` calls (spec §4.3) — so no `Fiddle::Importer.struct` is declared for it. This is where the file gains real Windows dependencies; like Phase 2's Task 3, nothing here is runnable on this machine — verification is CI-only (Task 6 uses these bindings; Task 17 is where CI actually exercises them).

- [ ] **Step 1: Implement**

```ruby
# lib/win32ole/jruby/array.rb
require 'fiddle'
require 'win32ole/jruby/win32'

class WIN32OLE
  module SafeArray
    W = Win32
    private_constant :W

    module_function

    def oleaut32
      W.oleaut32
    end

    def safe_array_create
      @safe_array_create ||= Fiddle::Function.new(
        oleaut32['SafeArrayCreate'], [W::WORD, W::DWORD, W::VOIDP], W::VOIDP, W::STDCALL
      )
    end

    def safe_array_create_vector
      @safe_array_create_vector ||= Fiddle::Function.new(
        oleaut32['SafeArrayCreateVector'], [W::WORD, W::LONG, W::DWORD], W::VOIDP, W::STDCALL
      )
    end

    def safe_array_destroy
      @safe_array_destroy ||= Fiddle::Function.new(oleaut32['SafeArrayDestroy'], [W::VOIDP], W::LONG, W::STDCALL)
    end

    def safe_array_lock
      @safe_array_lock ||= Fiddle::Function.new(oleaut32['SafeArrayLock'], [W::VOIDP], W::LONG, W::STDCALL)
    end

    def safe_array_unlock
      @safe_array_unlock ||= Fiddle::Function.new(oleaut32['SafeArrayUnlock'], [W::VOIDP], W::LONG, W::STDCALL)
    end

    def safe_array_get_dim
      @safe_array_get_dim ||= Fiddle::Function.new(oleaut32['SafeArrayGetDim'], [W::VOIDP], W::DWORD, W::STDCALL)
    end

    def safe_array_get_lbound
      @safe_array_get_lbound ||= Fiddle::Function.new(
        oleaut32['SafeArrayGetLBound'], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_get_ubound
      @safe_array_get_ubound ||= Fiddle::Function.new(
        oleaut32['SafeArrayGetUBound'], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_ptr_of_index
      @safe_array_ptr_of_index ||= Fiddle::Function.new(
        oleaut32['SafeArrayPtrOfIndex'], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_put_element
      @safe_array_put_element ||= Fiddle::Function.new(
        oleaut32['SafeArrayPutElement'], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_access_data
      @safe_array_access_data ||= Fiddle::Function.new(
        oleaut32['SafeArrayAccessData'], [W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_unaccess_data
      @safe_array_unaccess_data ||= Fiddle::Function.new(
        oleaut32['SafeArrayUnaccessData'], [W::VOIDP], W::LONG, W::STDCALL
      )
    end
  end
end
```

`GetDim` genuinely returns `UINT` directly (matches `TypeInfo.type_info_count_fn`'s identical `W::DWORD`-as-return-type precedent for `GetTypeInfoCount`) — not a copy-paste slip.

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/array.rb
git commit -m "jruby: SAFEARRAY oleaut32 Fiddle bindings"
```

---

## Task 5: `array.rb` — pure N-dimensional index-walk logic (local unit tests)

**Files:**
- Modify: `lib/win32ole/jruby/array.rb`
- Test: `test/win32ole/jruby/test_array.rb` (new — `RUBY_ENGINE == 'jruby'`-guarded per Global Constraints)

**Interfaces:**
- Consumes: nothing (pure Ruby `Array` logic, no native calls).
- Produces: `SafeArray.dimension_count(val)`, `.dimension_size(val, dim)`, `.dimension_sizes(val)`, `.nested_entry(val, pid)`, `.each_fill_index(sizes) { |pid| }`, `.each_read_index(lbounds, ubounds) { |pid| }`.

This is spec §7's explicitly called-out "OS-independent unit test... against synthetic nested-Array fixtures" and directly resolves §8 risk #3 ("diff its own index-walk against MRI's `ary_new_dim`/`ary_store_dim`/`dimension`/`ary_len_of_dim` line-by-line") — the functions below are a direct, function-by-function port of `ext/win32ole/win32ole.c`'s `dimension()` (lines 1148-1165), `ary_len_of_dim()` (1167-1191), `ole_ary_m_entry()` (963-974), `ole_set_safe_array()`'s index-increment loop (1116-1146), and `ole_variant2val()`'s array-branch index-increment loop (1476-1491) — **read while writing this plan, not reconstructed from the design doc's prose alone**.

**A real, non-obvious finding from that reading, worth stating plainly:** the fill direction (Ruby→`SAFEARRAY`, `ole_set_safe_array`) and the read direction (`SAFEARRAY`→Ruby, `ole_variant2val`) increment their index tuple in **opposite** orders — fill varies the *last* index (the innermost Ruby nesting level) fastest; read varies the *first* index (`SAFEARRAY` dimension 0) fastest. This is not a bug to "fix" into symmetry: `SAFEARRAY` dimension order is the reverse of Ruby-Array nesting order by COM convention, and this asymmetry is exactly what makes a round trip (`Array` → `SAFEARRAY` → `Array`) reconstruct the original nesting. `each_fill_index`/`each_read_index` below are intentionally two different functions, not one shared helper — do not "simplify" them into one.

- [ ] **Step 1: Write the failing test**

```ruby
# test/win32ole/jruby/test_array.rb
require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/array'

class TestSafeArray < Test::Unit::TestCase
  SA = WIN32OLE::SafeArray

  def test_dimension_count_for_a_flat_array
    assert_equal(1, SA.dimension_count([1, 2, 3]))
  end

  def test_dimension_count_for_a_nested_array
    assert_equal(2, SA.dimension_count([[1, 2], [3, 4]]))
  end

  def test_dimension_count_takes_the_max_depth_across_branches
    assert_equal(2, SA.dimension_count([[1, 2], 3]))
  end

  def test_dimension_sizes_for_a_2d_array
    assert_equal([2, 3], SA.dimension_sizes([[1, 2, 3], [4, 5, 6]]))
  end

  def test_dimension_sizes_takes_the_max_across_ragged_branches
    assert_equal([2, 3], SA.dimension_sizes([[1, 2, 3], [4]]))
  end

  def test_nested_entry_reads_the_addressed_element
    ary = [[1, 2], [3, 4]]
    assert_equal(4, SA.nested_entry(ary, [1, 1]))
  end

  def test_nested_entry_returns_nil_past_a_shorter_branch
    ary = [[1, 2, 3], [4]]
    assert_nil(SA.nested_entry(ary, [1, 2]))
  end

  def test_each_fill_index_enumerates_the_innermost_dimension_fastest
    tuples = SA.each_fill_index([2, 3]).to_a
    assert_equal([[0, 0], [0, 1], [0, 2], [1, 0], [1, 1], [1, 2]], tuples)
  end

  def test_each_fill_index_round_trips_a_2d_array_through_nested_entry
    ary = [[1, 2, 3], [4, 5, 6]]
    sizes = SA.dimension_sizes(ary)
    values = SA.each_fill_index(sizes).map { |pid| SA.nested_entry(ary, pid) }
    assert_equal([1, 2, 3, 4, 5, 6], values)
  end

  def test_each_read_index_enumerates_the_outermost_dimension_fastest
    tuples = SA.each_read_index([0, 0], [1, 2]).to_a
    assert_equal([[0, 0], [1, 0], [0, 1], [1, 1], [0, 2], [1, 2]], tuples)
  end

  def test_each_read_index_respects_nonzero_lower_bounds
    tuples = SA.each_read_index([1, 1], [2, 2]).to_a
    assert_equal([[1, 1], [2, 1], [1, 2], [2, 2]], tuples)
  end
end
end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_array.rb`
Expected: FAIL/ERROR — none of these methods exist yet.

- [ ] **Step 3: Implement**

Append inside `module SafeArray` in `lib/win32ole/jruby/array.rb` (before or after the bindings from Task 4, either position is fine — `module_function` already appears once):

```ruby
    # Port of ext/win32ole/win32ole.c's dimension() (line 1148) -- the
    # nesting depth is the MAX depth found across every branch, not just
    # the first element, so a ragged/mixed input still gets a well-defined
    # depth.
    def dimension_count(val)
      return 0 unless val.is_a?(::Array)

      val.reduce(0) { |max, v| [max, dimension_count(v)].max } + 1
    end

    # Port of ary_len_of_dim() (line 1167) -- the size at nesting level
    # `dim` (0-indexed, 0 == outermost) is the MAX size found across every
    # branch at that depth.
    def dimension_size(val, dim)
      return 0 unless val.is_a?(::Array)
      return val.size if dim.zero?

      val.reduce(0) { |max, v| [max, dimension_size(v, dim - 1)].max }
    end

    def dimension_sizes(val)
      dims = dimension_count(val)
      Array.new(dims) { |d| dimension_size(val, d) }
    end

    # Port of ole_ary_m_entry() (line 963) -- pid[0] indexes the outermost
    # Array, pid[1] the next level in, etc. Returns nil past a shorter
    # (ragged) branch, matching Ruby's own Array#[] out-of-range behavior.
    def nested_entry(val, pid)
      obj = val
      pid.each { |i| obj = obj.is_a?(::Array) ? obj[i] : nil }
      obj
    end

    # Port of ole_set_safe_array()'s pid increment loop (line 1116): the
    # LAST index (the innermost Ruby nesting level) varies fastest,
    # carrying left on overflow. This is the order Array->SAFEARRAY fill
    # (Task 6) writes elements in.
    def each_fill_index(sizes)
      return enum_for(:each_fill_index, sizes) unless block_given?

      pid = Array.new(sizes.size, 0)
      i = sizes.size - 1
      while i >= 0
        yield pid.dup
        pid[i] += 1
        if pid[i] >= sizes[i]
          pid[i] = 0
          i -= 1
        else
          i = sizes.size - 1
        end
      end
    end

    # Port of ole_variant2val()'s pid increment loop (line 1476): the
    # FIRST index (SAFEARRAY's own dimension 0) varies fastest --
    # deliberately the opposite order of each_fill_index above (see this
    # task's own note on why). lbounds/ubounds are inclusive, one entry
    # per dimension, as SafeArrayGetLBound/GetUBound return them.
    def each_read_index(lbounds, ubounds)
      return enum_for(:each_read_index, lbounds, ubounds) unless block_given?

      pid = lbounds.dup
      loop do
        yield pid.dup
        i = 0
        loop do
          pid[i] += 1
          break if pid[i] <= ubounds[i]

          pid[i] = lbounds[i]
          i += 1
          return if i == pid.size
        end
      end
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_array.rb`
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/array.rb test/win32ole/jruby/test_array.rb
git commit -m "jruby: N-dimensional SAFEARRAY index-walk logic, ported from MRI, local tests"
```

---

## Task 6: `array.rb` — `ruby_array_to_safearray`/`safearray_to_ruby_array` + `VT_UI1|VT_ARRAY` ↔ `String` fast path

**Files:**
- Modify: `lib/win32ole/jruby/array.rb`

**Interfaces:**
- Consumes: Task 4's bindings, Task 5's pure helpers, `WIN32OLE.ruby_value_to_variant_bytes`/`.variant_bytes_to_ruby_value` (Task 3) for `VT_VARIANT`-typed elements, `Win32::{VT_ARRAY,VT_VARIANT,VT_UI1,PACK_PTR}` (Task 1/Phase 1).
- Produces: `SafeArray.ruby_array_to_safearray(ary, elem_vt, bstrs_to_free = [])` (returns a raw `SAFEARRAY*` address), `SafeArray.safearray_to_ruby_array(psa_addr, elem_vt)` (returns a Ruby `Array`).

Native, CI-only (like Phase 2's Task 3-style bindings) — but unlike `Record` (see the plan header's finding), this one **does** have a real, always-available CI fixture: `Scripting.Dictionary#Keys`/`#Items` return a genuine 1-D `SAFEARRAY` of `VARIANT`s (`elem_vt == VT_VARIANT`), and that object is already proven present in this CI throughout Phase 1/2. Task 17 exercises this live.

**A distinction to get right, flagged explicitly because it's easy to miss:** `SafeArrayPutElement`'s `pv` argument for a **non**-`VT_VARIANT` element must point at the element's own *natural, unpadded* width (4 raw bytes for `VT_I4`, 1 for `VT_UI1`, etc.) — this is *not* the same shape as `Win32.pack_i4`'s 8-byte, zero-padded, VARIANT-body-slot format from Task 1 (that shape is specifically for a value living in a VARIANT's own 8-byte value slot, a different context). Only the `elem_vt == VT_VARIANT` branch below packs/unpacks a *full* `VARIANT` (via `WIN32OLE.ruby_value_to_variant_bytes`/`.variant_bytes_to_ruby_value`, Task 3) — every other `elem_vt` uses this task's own small `pack_scalar_element`/`unpack_scalar_element`, which are natural-width, not VARIANT-body-width.

- [ ] **Step 1: Implement**

Append inside `module SafeArray`:

```ruby
    ELEMENT_PACK_FORMAT = {
      W::VT_I1 => 'c', W::VT_UI1 => 'C', W::VT_I2 => 's', W::VT_UI2 => 'S',
      W::VT_I4 => 'l', W::VT_UI4 => 'L', W::VT_INT => 'l', W::VT_UINT => 'L',
      W::VT_I8 => 'q', W::VT_UI8 => 'Q', W::VT_R4 => 'f', W::VT_R8 => 'd',
      W::VT_ERROR => 'l', W::VT_BOOL => 's',
      W::VT_BSTR => W::PACK_PTR, W::VT_DISPATCH => W::PACK_PTR, W::VT_UNKNOWN => W::PACK_PTR
    }.freeze

    def pack_scalar_element(vt, value)
      fmt = ELEMENT_PACK_FORMAT.fetch(vt) { raise NotImplementedError, "VARTYPE #{vt} is not a supported array element type yet" }
      [value].pack(fmt)
    end

    def unpack_scalar_element(vt, bytes)
      fmt = ELEMENT_PACK_FORMAT.fetch(vt) { raise NotImplementedError, "VARTYPE #{vt} is not a supported array element type yet" }
      bytes.unpack1(fmt)
    end

    def ruby_array_to_safearray(ary, elem_vt, bstrs_to_free = [])
      base_vt = elem_vt & W::VT_TYPEMASK
      return ui1_safearray_from_bytes(ary) if base_vt == W::VT_UI1 && ary.is_a?(::String)

      sizes = dimension_sizes(ary)
      dims = sizes.size
      bounds = sizes.flat_map { |n| [n, 0] }.pack('L2' * dims)
      psa = safe_array_create.call(base_vt, dims, bounds)
      raise ::RuntimeError, 'memory allocation error' if psa.nil? || psa.to_i.zero?

      hr = safe_array_lock.call(psa)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayLock: #{W.hr_hex(hr)}" if W.failed?(hr)

      begin
        each_fill_index(sizes) do |pid|
          val = nested_entry(ary, pid)
          leaf = base_vt == W::VT_VARIANT ? WIN32OLE.ruby_value_to_variant_bytes(val, bstrs_to_free)
                                           : pack_scalar_element(base_vt, val)
          index_buf = pid.pack('l' * dims)
          hr = safe_array_put_element.call(psa, index_buf, W.native_pointer_for(leaf))
          raise WIN32OLE::RuntimeError, "failed to SafeArrayPutElement: #{W.hr_hex(hr)}" if W.failed?(hr)
        end
      ensure
        safe_array_unlock.call(psa)
      end
      psa
    end

    def safearray_to_ruby_array(psa, elem_vt)
      base_vt = elem_vt & W::VT_TYPEMASK
      return ui1_safearray_to_bytes(psa) if base_vt == W::VT_UI1

      dim = safe_array_get_dim.call(psa)
      lbounds = Array.new(dim) { |d| out = ("\x00" * 4).b; safe_array_get_lbound.call(psa, d + 1, out); out.unpack1('l') }
      ubounds = Array.new(dim) { |d| out = ("\x00" * 4).b; safe_array_get_ubound.call(psa, d + 1, out); out.unpack1('l') }

      hr = safe_array_lock.call(psa)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayLock: #{W.hr_hex(hr)}" if W.failed?(hr)

      result = nested_array_skeleton(ubounds.zip(lbounds).map { |u, l| u - l + 1 })
      begin
        each_read_index(lbounds, ubounds) do |pid|
          index_buf = pid.pack('l' * dim)
          elem_ptr_out = ("\x00" * W::PTR_SIZE).b
          hr = safe_array_ptr_of_index.call(psa, index_buf, elem_ptr_out)
          raise WIN32OLE::RuntimeError, "failed to SafeArrayPtrOfIndex: #{W.hr_hex(hr)}" if W.failed?(hr)

          elem_addr = elem_ptr_out.unpack1(W::PACK_PTR)
          val =
            if base_vt == W::VT_VARIANT
              WIN32OLE.variant_bytes_to_ruby_value(Fiddle::Pointer.new(elem_addr)[0, W::VARIANT_SIZE])
            else
              size = ELEMENT_PACK_FORMAT.fetch(base_vt) { raise NotImplementedError, "VARTYPE #{base_vt} is not a supported array element type yet" }
              unpack_scalar_element(base_vt, Fiddle::Pointer.new(elem_addr)[0, [1].pack(size).bytesize])
            end
          zero_based_pid = pid.each_with_index.map { |v, d| v - lbounds[d] }
          store_nested(result, zero_based_pid, val)
        end
      ensure
        safe_array_unlock.call(psa)
      end
      result
    end

    def nested_array_skeleton(sizes)
      return Array.new(sizes.first) if sizes.size == 1

      Array.new(sizes.first) { nested_array_skeleton(sizes[1..]) }
    end

    def store_nested(ary, pid, val)
      obj = ary
      pid[0..-2].each { |i| obj = obj[i] }
      obj[pid.last] = val
    end

    # VT_UI1|VT_ARRAY <-> String fast path: bulk-copy via
    # SafeArrayAccessData instead of the generic per-element path above --
    # a real, separate MRI code path (ole_val2olevariantdata's first
    # branch / folevariant_value's dim==1 reverse path), not an
    # optimization detail this design can skip: without it, a binary blob
    # argument would round-trip through an Array of small Integers instead
    # of staying a String.
    def ui1_safearray_from_bytes(bytes)
      bytes = bytes.b
      bounds = [bytes.bytesize, 0].pack('L2')
      psa = safe_array_create.call(W::VT_UI1, 1, bounds)
      raise ::RuntimeError, 'memory allocation error' if psa.nil? || psa.to_i.zero?

      data_out = ("\x00" * W::PTR_SIZE).b
      hr = safe_array_access_data.call(psa, data_out)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayAccessData: #{W.hr_hex(hr)}" if W.failed?(hr)

      Fiddle::Pointer.new(data_out.unpack1(W::PACK_PTR))[0, bytes.bytesize] = bytes
      safe_array_unaccess_data.call(psa)
      psa
    end

    def ui1_safearray_to_bytes(psa)
      data_out = ("\x00" * W::PTR_SIZE).b
      hr = safe_array_access_data.call(psa, data_out)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayAccessData: #{W.hr_hex(hr)}" if W.failed?(hr)

      lb_out = ("\x00" * 4).b
      ub_out = ("\x00" * 4).b
      safe_array_get_lbound.call(psa, 1, lb_out)
      safe_array_get_ubound.call(psa, 1, ub_out)
      len = ub_out.unpack1('l') - lb_out.unpack1('l') + 1
      bytes = Fiddle::Pointer.new(data_out.unpack1(W::PACK_PTR))[0, len].dup.b
      safe_array_unaccess_data.call(psa)
      bytes
    end
```

**`WIN32OLE` is referenced but not `require`d** — `array.rb` intentionally does not `require 'win32ole/jruby/win32ole'` (that would be circular: `win32ole.rb` itself requires `array.rb` in Task 15). This mirrors the established precedent in `dispatch.rb`, which references `WIN32OLE::RuntimeError` without requiring `win32ole.rb`, relying on `win32ole.rb` being the file that requires `dispatch.rb` (and, from Task 15 on, `array.rb`) in the first place — by the time any of these methods actually gets *called*, the whole require chain has finished loading.

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/array.rb
git commit -m "jruby: SAFEARRAY <-> Ruby Array marshaling (N-dimensional + VT_UI1<->String fast path)"
```

Not independently runnable — no Windows on this machine. Verified live in Task 17 against `Scripting.Dictionary#Keys`/`#Items`.

---

## Task 7: `record.rb` — `IRecordInfo` vtable slots + `GetRecordInfoFromTypeInfo` binding

**Files:**
- Create: `lib/win32ole/jruby/record.rb`

**Interfaces:**
- Consumes: `Win32.vtable_function`, `Win32.vtable_address`, `Win32.oleaut32` (Phase 1), `TypeInfo.type_info_count_fn`/`.type_info_fn`/`.documentation_fn_for_typeinfo` (Phase 2, reused for the typelib member scan in Task 8).
- Produces: `WIN32OLE::Record::IRECORDINFO_VTBL` (full 16-slot documentation table, spec §4.4), `.record_init_fn`/`.get_name_fn`/`.get_size_fn`/`.get_field_no_copy_fn`/`.put_field_fn`/`.get_field_names_fn` (the 6 bound members) and `.get_record_info_from_type_info` (the standalone `oleaut32` export).

Native, CI-only, same shape as Phase 2's Task 3 and this plan's Task 4 — no local test. **Per this plan's header finding, these vtable slot numbers cannot be empirically validated against a live `IRecordInfo*` in this project's CI** (unlike Phase 2's `ITypeInfo`/`ITypeLib`, which had `Scripting.Dictionary` to validate against) — they are transcribed from `oaidl.h` (spec §4.4) and from the same slot numbers already used by `win32ole_record.c`'s own C vtable calls, which is corroborating evidence but not a live-code-path confirmation.

- [ ] **Step 1: Implement**

```ruby
# lib/win32ole/jruby/record.rb
require 'fiddle'
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'

class WIN32OLE
  class Record
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    # Full slot table transcribed from oaidl.h, kept for documentation
    # completeness (spec §3 lists everything but the six below as out of
    # scope for this phase -- unbound, no Fiddle::Function declared).
    IRECORDINFO_VTBL = {
      RecordInit: 3, RecordClear: 4, RecordCopy: 5, GetGuid: 6, GetName: 7,
      GetSize: 8, GetTypeInfo: 9, GetField: 10, GetFieldNoCopy: 11,
      PutField: 12, PutFieldNoCopy: 13, GetFieldNames: 14, IsMatchingType: 15,
      RecordCreate: 16, RecordCreateCopy: 17, RecordDestroy: 18
    }.freeze

    def self.record_init_fn(pri)
      @record_init_fns ||= {}
      @record_init_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:RecordInit], [W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.get_name_fn(pri)
      @get_name_fns ||= {}
      @get_name_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:GetName], [W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.get_size_fn(pri)
      @get_size_fns ||= {}
      @get_size_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:GetSize], [W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.get_field_no_copy_fn(pri)
      @get_field_no_copy_fns ||= {}
      @get_field_no_copy_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:GetFieldNoCopy], [W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.put_field_fn(pri)
      @put_field_fns ||= {}
      @put_field_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:PutField], [W::VOIDP, W::DWORD, W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.get_field_names_fn(pri)
      @get_field_names_fns ||= {}
      @get_field_names_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:GetFieldNames], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.get_record_info_from_type_info_fn
      @get_record_info_from_type_info_fn ||= Fiddle::Function.new(
        W.oleaut32['GetRecordInfoFromTypeInfo'], [W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end
  end
end
```

`GetFieldNoCopy(PVOID pvData, LPCOLESTR szFieldName, VARIANT *pvarField, PVOID *ppvDataCArray)` and `PutField(ULONG wFlags, PVOID pvData, LPCOLESTR szFieldName, VARIANT *pvarField)` — both real 4/5-argument COM signatures per `oaidl.h`; do not drop `ppvDataCArray`/`wFlags` when calling these in Task 8/10 even though this phase's tests never exercise a `SAFEARRAY`-valued field.

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/record.rb
git commit -m "jruby: IRecordInfo vtable slots + GetRecordInfoFromTypeInfo binding"
```

---

## Task 8: `record.rb` — `WIN32OLE::Record.new` construction

**Files:**
- Modify: `lib/win32ole/jruby/record.rb`
- Test: `test/win32ole/jruby/test_record.rb` (new, `RUBY_ENGINE == 'jruby'`-guarded)

**Interfaces:**
- Consumes: Task 7's bindings, `TypeInfo.type_info_count_fn`/`.type_info_fn`/`.documentation_fn_for_typeinfo` (Phase 2), `WIN32OLE::TypeLib`/`WIN32OLE`'s `@ptr` ivars (Phase 2/1).
- Produces: `WIN32OLE::Record.new(typename, oleobj)`.

**A real correction to the design spec's own summary, found by reading `ext/win32ole/win32ole_record.c:281-322` (`folerecord_initialize`) and `:122-169` (`olerecord_set_ivar`) directly:** `Record.new` does **not** allocate a native buffer, call `RecordInit`, or read any field's current value at construction time. It only resolves `IRecordInfo*`, reads `@typename` (`GetName`), and reads the field *names* (`GetFieldNames`) — every field starts as `nil` in `@fields`. The buffer/`RecordInit`/`GetFieldNoCopy`-populated-values path only happens for a record *wrapping an existing native buffer* (a method's return value — Task 10's `Record.from_irecordinfo_and_buffer`, MRI's `create_win32ole_record`), which is a *different* construction path than `.new`. Do not conflate them — this plan's spec summary (§4.4) reads as if `.new` populates real values; the actual MRI source does not.

- [ ] **Step 1: Write the failing test**

Construction itself needs a live `ITypeLib*`/`WIN32OLE`, which this machine can't provide — so this step tests the pieces that *are* locally checkable: type-checking `oleobj`, and the resulting object's Hash-shape contract, via `allocate` + direct ivar injection (bypassing `initialize` entirely — the same bypass technique Task 9 also uses for the field-access methods, since those don't need a live COM object either).

```ruby
# test/win32ole/jruby/test_record.rb
require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/record'

class TestRecordConstruction < Test::Unit::TestCase
  def test_new_rejects_an_oleobj_that_is_neither_win32ole_nor_typelib
    assert_raise(TypeError) { WIN32OLE::Record.new('Book', Object.new) }
  end
end
end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_record.rb`
Expected: FAIL — `Record#initialize` doesn't exist yet (currently no `record.rb` class body beyond Task 7's bindings).

- [ ] **Step 3: Implement**

```ruby
    def initialize(typename, oleobj)
      typename = typename.to_s
      itypelib_ptr = resolve_itypelib_ptr(oleobj)
      pri = find_record_info(itypelib_ptr, typename)
      unless pri
        raise WIN32OLE::RuntimeError, "fail to query IRecordInfo interface for `#{typename}'"
      end

      set_record_info(pri, nil)
    end

    def to_h
      @fields
    end

    def typename
      @typename
    end

    def inspect
      "#<WIN32OLE::Record:#{@typename}>"
    end

    private

    def resolve_itypelib_ptr(oleobj)
      case oleobj
      when WIN32OLE::TypeLib
        oleobj.instance_variable_get(:@ptr)
      when WIN32OLE
        oleobj.ole_typelib.instance_variable_get(:@ptr)
      else
        raise TypeError, "2nd argument should be WIN32OLE object or WIN32OLE::TypeLib object, got #{oleobj.class}"
      end
    end

    # Port of recordinfo_from_itypelib (win32ole_record.c:29-59): linear-scan
    # the typelib's own members for a name match (GetDocumentation), same
    # already-open-typelib walk Phase 2's TypeLib#ole_types performs -- NOT
    # the registry-tree walk Phase 2 §3 excluded. Every non-matching
    # ITypeInfo* GetTypeInfo AddRef'd along the way must be Released, or
    # the scan leaks a reference per member it skips past.
    def find_record_info(itypelib_ptr, typename)
      count = TI.type_info_count_fn(itypelib_ptr).call(itypelib_ptr)
      count.times do |i|
        ti_out = ("\x00" * W::PTR_SIZE).b
        next if W.failed?(TI.type_info_fn(itypelib_ptr).call(itypelib_ptr, i, ti_out))

        itypeinfo_ptr = ti_out.unpack1(W::PACK_PTR)
        name_out = ("\x00" * W::PTR_SIZE).b
        TI.documentation_fn_for_typeinfo(itypeinfo_ptr).call(itypeinfo_ptr, -1, name_out, nil, nil, nil)
        name_bstr = name_out.unpack1(W::PACK_PTR)
        name = W.bstr_to_s(name_bstr)
        W.sys_free_string.call(name_bstr) unless name_bstr.zero?

        if name == typename
          pri_out = ("\x00" * W::PTR_SIZE).b
          hr = self.class.get_record_info_from_type_info_fn.call(itypeinfo_ptr, pri_out)
          release_itypeinfo(itypeinfo_ptr)
          return W.failed?(hr) ? nil : pri_out.unpack1(W::PACK_PTR)
        end
        release_itypeinfo(itypeinfo_ptr)
      end
      nil
    end

    def release_itypeinfo(itypeinfo_ptr)
      W.vtable_function(itypeinfo_ptr, 2, [W::VOIDP], W::DWORD).call(itypeinfo_ptr)
    end

    # Shared by .new (prec = nil, every field starts nil -- see this
    # task's own correction above) and Task 10's .from_irecordinfo_and_buffer
    # (prec = a real native buffer, fields read via GetFieldNoCopy). Ports
    # olerecord_set_ivar (win32ole_record.c:122-169).
    def set_record_info(pri, prec)
      @pri = pri
      install_finalizer

      name_out = ("\x00" * W::PTR_SIZE).b
      if self.class.get_name_fn(pri).call(pri, name_out).zero?
        bstr = name_out.unpack1(W::PACK_PTR)
        @typename = W.bstr_to_s(bstr)
        W.sys_free_string.call(bstr) unless bstr.zero?
      end

      count_out = ("\x00" * 4).b
      hr = self.class.get_field_names_fn(pri).call(pri, count_out, nil)
      count = count_out.unpack1('L')
      return if W.failed?(hr) || count.zero?

      names_out = ("\x00" * (count * W::PTR_SIZE)).b
      count_out = [count].pack('L')
      self.class.get_field_names_fn(pri).call(pri, count_out, names_out)
      bstrs = names_out.unpack(W::PACK_PTR * count)

      @fields = {}
      bstrs.each do |bstr|
        name = W.bstr_to_s(bstr)
        val = nil
        if prec
          var_out = ("\x00" * W::VARIANT_SIZE).b
          pdata_out = ("\x00" * W::PTR_SIZE).b
          hr = self.class.get_field_no_copy_fn(pri).call(pri, prec, bstr, var_out, pdata_out)
          val = WIN32OLE.variant_bytes_to_ruby_value(var_out) if hr.zero?
        end
        @fields[name] = val
        W.sys_free_string.call(bstr) unless bstr.zero?
      end
    end

    def install_finalizer
      pri = @pri
      release_fn = W.vtable_function(pri, 2, [W::VOIDP], W::DWORD)
      ObjectSpace.define_finalizer(self, self.class.finalizer(pri, release_fn))
    end

    def self.finalizer(pri, release_fn)
      proc { release_fn.call(pri) unless pri.zero? }
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_record.rb`
Expected: PASS. (Only the type-check path — the live-COM path is CI-only and, per the plan header, has no fixture in this project's CI to actually exercise it against.)

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/record.rb test/win32ole/jruby/test_record.rb
git commit -m "jruby: WIN32OLE::Record.new construction (typelib scan + IRecordInfo resolution)"
```

---

## Task 9: `record.rb` — field access (`method_missing`, `ole_instance_variable_get`/`_set`)

**Files:**
- Modify: `lib/win32ole/jruby/record.rb`
- Test: `test/win32ole/jruby/test_record.rb`

**Interfaces:**
- Consumes: `@fields` (Task 8).
- Produces: `Record#method_missing` (get/set), `#ole_instance_variable_get`/`#ole_instance_variable_set`, `#to_h` (already added in Task 8).

All pure `Hash` logic — fully local-testable via `allocate` + direct `@fields`/`@typename` injection (bypassing `initialize`/COM entirely, the same technique Task 8's own test used). Ports `olerecord_ivar_get`/`_set`/`folerecord_method_missing`/`folerecord_ole_instance_variable_get`/`_set` (`win32ole_record.c:400-520`) exactly, including the one behavior worth calling out: **`to_h` returns `@fields` directly, not a dup** — mutating the returned `Hash` mutates the `Record`'s own state too. This is MRI's actual behavior (`folerecord_to_h` is a bare `rb_ivar_get`), not an oversight to "fix."

- [ ] **Step 1: Write the failing test**

```ruby
class TestRecordFieldAccess < Test::Unit::TestCase
  def setup
    @record = WIN32OLE::Record.allocate
    @record.instance_variable_set(:@typename, 'Book')
    @record.instance_variable_set(:@fields, { 'title' => 'The Ruby Book', 'cost' => 20 })
  end

  def test_typename
    assert_equal('Book', @record.typename)
  end

  def test_to_h_returns_the_fields_hash
    assert_equal({ 'title' => 'The Ruby Book', 'cost' => 20 }, @record.to_h)
  end

  def test_to_h_is_not_a_defensive_copy
    @record.to_h['cost'] = 99
    assert_equal(99, @record.to_h['cost'])
  end

  def test_method_missing_getter
    assert_equal('The Ruby Book', @record.title)
  end

  def test_method_missing_setter
    @record.title = 'Ruby'
    assert_equal('Ruby', @record.title)
  end

  def test_method_missing_getter_raises_key_error_for_unknown_field
    assert_raise(KeyError) { @record.no_such_field }
  end

  def test_method_missing_setter_raises_key_error_for_unknown_field
    assert_raise(KeyError) { @record.no_such_field = 1 }
  end

  def test_ole_instance_variable_get
    assert_equal(20, @record.ole_instance_variable_get(:cost))
  end

  def test_ole_instance_variable_set
    @record.ole_instance_variable_set(:cost, 30)
    assert_equal(30, @record.ole_instance_variable_get(:cost))
  end

  def test_inspect
    assert_equal('#<WIN32OLE::Record:Book>', @record.inspect)
  end
end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_record.rb`
Expected: FAIL — `method_missing`/`ole_instance_variable_get`/`_set` don't exist yet.

- [ ] **Step 3: Implement**

Add to the `public` section of `WIN32OLE::Record` (above the `private` keyword already added in Task 8):

```ruby
    def method_missing(name, *args)
      sname = name.to_s
      case args.size
      when 0 then @fields.fetch(sname)
      when 1
        key = sname.end_with?('=') ? sname[0..-2] : sname
        @fields.fetch(key) # raises KeyError before writing, matching MRI
        @fields[key] = args.first
      else
        super
      end
    end

    def respond_to_missing?(name, include_private = false)
      @fields.key?(name.to_s.sub(/=\z/, '')) || super
    end

    def ole_instance_variable_get(name)
      unless name.is_a?(String) || name.is_a?(Symbol)
        raise TypeError, 'wrong argument type (expected String or Symbol)'
      end

      @fields.fetch(name.to_s)
    end

    def ole_instance_variable_set(name, val)
      unless name.is_a?(String) || name.is_a?(Symbol)
        raise TypeError, 'wrong argument type (expected String or Symbol)'
      end

      key = name.to_s
      @fields.fetch(key)
      @fields[key] = val
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_record.rb`
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/record.rb test/win32ole/jruby/test_record.rb
git commit -m "jruby: WIN32OLE::Record field access (method_missing, ole_instance_variable_get/set)"
```

---

## Task 10: `record.rb` — marshal `Record` → `VARIANT` + the incoming-result construction path

**Files:**
- Modify: `lib/win32ole/jruby/record.rb`

**Interfaces:**
- Consumes: Task 7's `put_field_fn`/`get_size_fn`/`record_init_fn`, Task 1's generalized `Win32.pack_variant`/`unpack_variant` (body-length-aware), Task 8's `set_record_info` (shared by both construction paths).
- Produces: `Record#to_variant_bytes` (marshal out, for win32ole.rb's Task 15 argument-building hook), `Record.from_irecordinfo_and_buffer(pri, prec)` (marshal in, for Task 15's `VT_RECORD` result-unmarshaling hook).

Native, CI-only. Ports `ole_rec2variant`/`hash2olerec` (`win32ole_record.c:61-120`) and `create_win32ole_record` (`:171-177`) exactly — **`to_variant_bytes` always allocates a *fresh* buffer (`GetSize` + `RecordInit`) and fills it from the record's *current* `@fields` Hash state every time it's called**, it does not reuse or mutate any buffer from construction time (`Record.new` never allocated one in the first place — Task 8's correction). `nil`-valued fields are skipped (not written via `PutField`), matching `hash2olerec`'s own `val != Qnil` guard.

- [ ] **Step 1: Implement**

```ruby
    VT_RECORD_BODY_SIZE = W::PTR_SIZE * 2 # BRECORD: { PVOID pvRecord; IRecordInfo *pRecInfo; }

    def to_variant_bytes
      size_out = ("\x00" * 4).b
      hr = self.class.get_size_fn(@pri).call(@pri, size_out)
      raise WIN32OLE::RuntimeError, "failed to get size for allocation of VT_RECORD object: #{W.hr_hex(hr)}" if W.failed?(hr)

      size = size_out.unpack1('L')
      buffer_ptr = Fiddle::Pointer.malloc(size)
      hr = self.class.record_init_fn(@pri).call(@pri, buffer_ptr.to_i)
      raise WIN32OLE::RuntimeError, "failed to initialize VT_RECORD object: #{W.hr_hex(hr)}" if W.failed?(hr)

      @fields.each do |name, val|
        next if val.nil?

        var_bytes = WIN32OLE.ruby_value_to_variant_bytes(val, [])
        hr = self.class.put_field_fn(@pri).call(
          @pri, W::DISPATCH_PROPERTYPUT, buffer_ptr.to_i, W.wstr(name), W.native_pointer_for(var_bytes)
        )
        raise WIN32OLE::RuntimeError, "failed to putfield of `#{name}': #{W.hr_hex(hr)}" if W.failed?(hr)
      end

      body = [buffer_ptr.to_i, @pri].pack("#{W::PACK_PTR}2")
      W.pack_variant(W::VT_RECORD, body)
    end

    def self.from_irecordinfo_and_buffer(pri, prec)
      allocate.tap { |rec| rec.send(:set_record_info, pri, prec) }
    end
```

**Note on `PACK_PTR` reuse:** `body` here is exactly `VT_RECORD_BODY_SIZE` (2 native pointers) bytes, which is `<= VARIANT_SIZE - 8` by construction (spec §4.2: this is *why* `VARIANT_SIZE` is 24 on x64 in the first place) — `Win32.pack_variant` (Task 1) accepts it unchanged, no new `pack_record` wrapper needed.

Add the corresponding unmarshal side to `win32.rb` is *not* needed as a separate function either — Task 15's dispatch hook reads a `VT_RECORD` result directly: `pri_ptr, buffer_ptr = *W.unpack_variant(bytes, body_size: Record::VT_RECORD_BODY_SIZE).last.unpack("#{W::PACK_PTR}2")`-shaped code, written out in Task 15 itself where it's used.

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/record.rb
git commit -m "jruby: Record#to_variant_bytes (PutField-based marshal out) + from_irecordinfo_and_buffer (marshal in)"
```

Not independently runnable — no Windows on this machine, and (per the plan header) no live `IRecordInfo`-bearing fixture in this project's CI either. This is the one piece of Phase 3 that stays structurally-verified-only through Task 17.

---

## Task 11: `variant.rb` — `WIN32OLE::VariantType`/`VARIANT` constant module

**Files:**
- Create: `lib/win32ole/jruby/variant.rb`
- Test: `test/win32ole/jruby/test_variant.rb` (new, `RUBY_ENGINE == 'jruby'`-guarded)

**Interfaces:**
- Consumes: nothing (pure constants).
- Produces: `WIN32OLE::VariantType::VT_*` (every VARTYPE MRI's `win32ole_variant_m.c` defines, independent of which ones this phase's marshaling code implements — spec §1.2's closing point), `WIN32OLE::VARIANT` (aliased to `VariantType`).

Pure, locally testable — same shape as Task 1's constant additions.

- [ ] **Step 1: Write the failing test**

```ruby
# test/win32ole/jruby/test_variant.rb
require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/variant'

class TestVariantType < Test::Unit::TestCase
  def test_vt_i4_matches_win32_constant
    assert_equal(WIN32OLE::Win32::VT_I4, WIN32OLE::VariantType::VT_I4)
  end

  def test_vt_array_and_byref_flags
    assert_equal(0x2000, WIN32OLE::VariantType::VT_ARRAY)
    assert_equal(0x4000, WIN32OLE::VariantType::VT_BYREF)
  end

  def test_vt_cy_and_vt_date_are_still_defined_as_constants_even_though_unsupported
    # Non-goal (spec §3) means the *marshaling* code doesn't implement
    # these -- the constant table itself is unconditional (spec §1.2).
    assert_equal(6, WIN32OLE::VariantType::VT_CY)
    assert_equal(7, WIN32OLE::VariantType::VT_DATE)
  end

  def test_variant_is_aliased_to_varianttype
    assert_same(WIN32OLE::VariantType, WIN32OLE::VARIANT)
  end
end
end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_variant.rb`
Expected: FAIL — file doesn't exist yet.

- [ ] **Step 3: Implement**

```ruby
# lib/win32ole/jruby/variant.rb
require 'win32ole/jruby/win32'

class WIN32OLE
  module VariantType
    VT_EMPTY = 0
    VT_NULL = 1
    VT_I2 = 2
    VT_I4 = 3
    VT_R4 = 4
    VT_R8 = 5
    VT_CY = 6
    VT_DATE = 7
    VT_BSTR = 8
    VT_DISPATCH = 9
    VT_ERROR = 10
    VT_BOOL = 11
    VT_VARIANT = 12
    VT_UNKNOWN = 13
    VT_DECIMAL = 14
    VT_I1 = 16
    VT_UI1 = 17
    VT_UI2 = 18
    VT_UI4 = 19
    VT_I8 = 20
    VT_UI8 = 21
    VT_INT = 22
    VT_UINT = 23
    VT_VOID = 24
    VT_HRESULT = 25
    VT_PTR = 26
    VT_SAFEARRAY = 27
    VT_CARRAY = 28
    VT_USERDEFINED = 29
    VT_LPSTR = 30
    VT_LPWSTR = 31
    VT_RECORD = 36
    VT_INT_PTR = 37
    VT_UINT_PTR = 38
    VT_FILETIME = 64
    VT_BLOB = 65
    VT_STREAM = 66
    VT_STORAGE = 67
    VT_STREAMED_OBJECT = 68
    VT_STORED_OBJECT = 69
    VT_BLOB_OBJECT = 70
    VT_CF = 71
    VT_CLSID = 72
    VT_VERSIONED_STREAM = 73
    VT_BSTR_BLOB = 0xFFF
    VT_VECTOR = 0x1000
    VT_ARRAY = 0x2000
    VT_BYREF = 0x4000
    VT_RESERVED = 0x8000
    VT_ILLEGAL = 0xFFFF
    VT_ILLEGALMASKED = 0xFFF
    VT_TYPEMASK = 0xFFF
  end

  VARIANT = VariantType
end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_variant.rb`
Expected: all PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/variant.rb test/win32ole/jruby/test_variant.rb
git commit -m "jruby: WIN32OLE::VariantType/VARIANT full VT_* constant table"
```

---

## Task 12: `variant.rb` — `Variant.new`, `value`/`value=`/`vartype`

**Files:**
- Modify: `lib/win32ole/jruby/variant.rb`
- Test: `test/win32ole/jruby/test_variant.rb`

**Interfaces:**
- Consumes: `WIN32OLE.ruby_value_to_variant_bytes`/`.variant_bytes_to_ruby_value` (Task 3), `Win32.pack_byref`/`.variant_change_type` (Task 2), `Win32`'s scalar pack/unpack family (Task 1), `SafeArray.ruby_array_to_safearray`/`.safearray_to_ruby_array` (Task 6).
- Produces: `WIN32OLE::Variant.new(val, vartype = nil, *)`, `#value`, `#value=`, `#vartype`.

Native for the explicit-mismatched-type (`VariantChangeTypeEx`) path only — the inference path (`Variant.new(val)`, one argument) reuses `WIN32OLE.ruby_value_to_variant_bytes` unchanged, which is itself pure/native-free except for the `:bstr` case's `SysAllocString` call (already true of ordinary dispatch argument-building, not new risk). CI-only overall, consistent with the rest of this file.

- [ ] **Step 1: Implement**

```ruby
class WIN32OLE
  class Variant
    W = Win32
    SA = SafeArray
    VT = VariantType
    private_constant :W, :SA, :VT

    def initialize(val, vartype = nil, _reserved = nil)
      if vartype.nil?
        @var = WIN32OLE.ruby_value_to_variant_bytes(val, @bstrs_to_free = [])
        return
      end

      raise ArgumentError, 'WIN32OLE::Variant does not support VT_RECORD; use WIN32OLE::Record instead' if (vartype & VT::VT_TYPEMASK) == VT::VT_RECORD

      base_vt = vartype & VT::VT_TYPEMASK
      byref = (vartype & VT::VT_BYREF) != 0

      @realvar =
        if (vartype & VT::VT_ARRAY) != 0
          @bstrs_to_free = []
          psa = SA.ruby_array_to_safearray(val, base_vt, @bstrs_to_free)
          W.pack_variant(vartype & ~VT::VT_BYREF, W.pack_pointer(psa.to_i))
        else
          @bstrs_to_free = []
          pack_scalar_explicit(base_vt, val)
        end

      @var = byref ? W.pack_byref(vartype & ~VT::VT_BYREF, @realvar) : @realvar
    end

    def value
      vt, = W.unpack_variant(@var)
      base_vt = vt & VT::VT_TYPEMASK
      return SA.safearray_to_ruby_array(current_array_ptr, base_vt) if (vt & VT::VT_ARRAY) != 0

      WIN32OLE.variant_bytes_to_ruby_value(current_scalar_bytes)
    end

    def value=(val)
      vt, = W.unpack_variant(@var)
      base_vt = vt & VT::VT_TYPEMASK
      if (vt & VT::VT_ARRAY) != 0
        unless val.is_a?(::String) && base_vt == VT::VT_UI1
          raise WIN32OLE::RuntimeError, 'array value can only be replaced with a String for a VT_UI1|VT_ARRAY Variant'
        end

        psa = SA.ruby_array_to_safearray(val, base_vt, @bstrs_to_free ||= [])
        @realvar = W.pack_variant(vt & ~VT::VT_BYREF, W.pack_pointer(psa.to_i))
      else
        @realvar = pack_scalar_explicit(base_vt, val)
      end
      @var = (vt & VT::VT_BYREF) != 0 ? W.pack_byref(vt & ~VT::VT_BYREF, @realvar) : @realvar
    end

    def vartype
      vt, = W.unpack_variant(@var)
      vt
    end

    private

    def current_array_ptr
      _vt, payload = W.unpack_variant(@var)
      W.unpack_pointer(payload)
    end

    def current_scalar_bytes
      @var
    end

    SCALAR_PACK = {
      VT::VT_I1 => :pack_i1, VT::VT_UI1 => :pack_ui1, VT::VT_I2 => :pack_i2, VT::VT_UI2 => :pack_ui2,
      VT::VT_I4 => :pack_i4, VT::VT_UI4 => :pack_ui4, VT::VT_INT => :pack_int, VT::VT_UINT => :pack_uint,
      VT::VT_I8 => :pack_i8, VT::VT_UI8 => :pack_ui8, VT::VT_R4 => :pack_r4, VT::VT_R8 => :pack_r8,
      VT::VT_BOOL => :pack_bool, VT::VT_ERROR => :pack_error, VT::VT_EMPTY => :pack_empty
    }.freeze

    def pack_scalar_explicit(base_vt, val)
      if base_vt == VT::VT_EMPTY || base_vt == VT::VT_NULL
        return W.pack_variant(base_vt, W.pack_empty)
      end
      if base_vt == VT::VT_BSTR
        bstr = W.sys_alloc_string.call(W.wstr(val))
        (@bstrs_to_free ||= []) << bstr
        return W.pack_variant(base_vt, W.pack_pointer(bstr))
      end
      if base_vt == VT::VT_DISPATCH || base_vt == VT::VT_UNKNOWN
        ptr = val.nil? ? 0 : val.instance_variable_get(:@ptr)
        return W.pack_variant(base_vt, W.pack_pointer(ptr))
      end

      pack_method = SCALAR_PACK[base_vt]
      unless pack_method
        raise NotImplementedError, "VARTYPE #{base_vt} is not supported yet"
      end

      inferred_vt = W.ruby_to_variant_type(val) rescue nil
      needs_coercion = inferred_vt && W::VT_FOR_TYPE[inferred_vt] != base_vt
      if needs_coercion
        coerce_via_variant_change_type(val, base_vt)
      else
        W.pack_variant(base_vt, W.send(pack_method, val))
      end
    end

    # Mirrors MRI's ole_val2variant_ex + VariantChangeTypeEx fallback for
    # a mismatched explicit VARTYPE (spec §4.2) -- e.g.
    # Variant.new("2e3", VariantType::VT_R4). Native call, CI-only.
    def coerce_via_variant_change_type(val, base_vt)
      src_type = W.ruby_to_variant_type(val)
      src = W.pack_variant(W::VT_FOR_TYPE.fetch(src_type), WIN32OLE.ruby_value_to_variant_bytes(val, @bstrs_to_free ||= [])[8, 8])
      dest = ("\x00" * W::VARIANT_SIZE).b
      hr = W.variant_change_type.call(dest, src, W::LOCALE_SYSTEM_DEFAULT, 0, base_vt)
      raise WIN32OLE::RuntimeError, "failed to change variant type: #{W.hr_hex(hr)}" if W.failed?(hr)

      dest
    end
  end
end
```

`require 'win32ole/jruby/array'` at the top of `variant.rb` (alongside the existing `require 'win32ole/jruby/win32'`) — unlike `array.rb`/`record.rb`'s deliberate non-require of `win32ole.rb` (Task 6's note), `variant.rb` requiring `array.rb` is *not* circular (`array.rb` doesn't require `variant.rb` back), so require it normally rather than relying on load order.

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/variant.rb
git commit -m "jruby: WIN32OLE::Variant.new (inference + explicit scalar/BYREF path), value/value=/vartype"
```

---

## Task 13: `variant.rb` — `.array(dims, vt)`, `#[]`, `#[]=`

**Files:**
- Modify: `lib/win32ole/jruby/variant.rb`

**Interfaces:**
- Consumes: `SafeArray.safe_array_create`/`.safe_array_ptr_of_index`/`.safe_array_put_element`/`.safe_array_lock`/`.safe_array_unlock` (Task 4), `SafeArray.pack_scalar_element`/`.unpack_scalar_element`/`ELEMENT_PACK_FORMAT` (Task 6).
- Produces: `WIN32OLE::Variant.array(dims, vt)`, `Variant#[]`, `Variant#[]=`.

Native, CI-only.

- [ ] **Step 1: Implement**

```ruby
    def self.array(dims, vt)
      bounds = dims.flat_map { |n| [n, 0] }.pack('L2' * dims.size)
      psa = SA.safe_array_create.call(vt & VT::VT_TYPEMASK, dims.size, bounds)
      raise ::RuntimeError, 'memory allocation error' if psa.nil? || psa.to_i.zero?

      allocate.tap { |v| v.send(:set_array_var, psa, vt) }
    end

    def [](*indices)
      base_vt, psa = array_state
      index_buf = indices.pack('l' * indices.size)
      elem_ptr_out = ("\x00" * W::PTR_SIZE).b
      hr = SA.safe_array_ptr_of_index.call(psa, index_buf, elem_ptr_out)
      raise WIN32OLE::RuntimeError, "failed to SafeArrayPtrOfIndex: #{W.hr_hex(hr)}" if W.failed?(hr)

      elem_addr = elem_ptr_out.unpack1(W::PACK_PTR)
      if base_vt == VT::VT_VARIANT
        WIN32OLE.variant_bytes_to_ruby_value(Fiddle::Pointer.new(elem_addr)[0, W::VARIANT_SIZE])
      else
        fmt = SA::ELEMENT_PACK_FORMAT.fetch(base_vt)
        SA.unpack_scalar_element(base_vt, Fiddle::Pointer.new(elem_addr)[0, [1].pack(fmt).bytesize])
      end
    end

    def []=(*args)
      val = args.pop
      indices = args
      base_vt, psa = array_state
      leaf = base_vt == VT::VT_VARIANT ? WIN32OLE.ruby_value_to_variant_bytes(val, @bstrs_to_free ||= [])
                                        : SA.pack_scalar_element(base_vt, val)
      index_buf = indices.pack('l' * indices.size)
      hr = SA.safe_array_put_element.call(psa, index_buf, W.native_pointer_for(leaf))
      raise WIN32OLE::RuntimeError, "failed to SafeArrayPutElement: #{W.hr_hex(hr)}" if W.failed?(hr)

      val
    end

    private

    def set_array_var(psa, vt)
      @var = @realvar = W.pack_variant(vt & ~VT::VT_BYREF, W.pack_pointer(psa.to_i))
    end

    def array_state
      vt, payload = W.unpack_variant(@var)
      [vt & VT::VT_TYPEMASK, W.unpack_pointer(payload)]
    end
```

`#[]`/`#[]=` intentionally skip `SafeArrayLock`/`Unlock` around a single-element access (matching `folevariant_ary_aref`/`_aset`'s own per-call, not batch-locked, shape per spec §4.5) — this differs from Task 6's `ruby_array_to_safearray`/`safearray_to_ruby_array`, which lock once around a whole-array walk. Both are faithful to their respective MRI call sites, not an inconsistency to reconcile.

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/variant.rb
git commit -m "jruby: Variant.array, #[], #[]="
```

---

## Task 14: `variant.rb` — `Empty`/`Null`/`Nothing`/`NoParam` constants

**Files:**
- Modify: `lib/win32ole/jruby/variant.rb`

**Interfaces:**
- Consumes: `Variant.new` (Task 12).
- Produces: `WIN32OLE::Variant::Empty`, `::Null`, `::Nothing`, `::NoParam`.

Depends on Task 12's explicit-VARTYPE path already working — a natural self-test of that path before any external COM call touches it (spec §4.5), matching MRI's own choice to build these eagerly at class-init time.

- [ ] **Step 1: Implement**

Append after the `class Variant` body closes (class-level constants, evaluated once the class itself is fully defined):

```ruby
    DISP_E_PARAMNOTFOUND = -2147352572 # 0x80020004

    Empty = new(nil, VariantType::VT_EMPTY)
    Null = new(nil, VariantType::VT_NULL)
    Nothing = new(nil, VariantType::VT_DISPATCH)
    NoParam = new(DISP_E_PARAMNOTFOUND, VariantType::VT_ERROR)
```

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/variant.rb
git commit -m "jruby: Variant::Empty/Null/Nothing/NoParam"
```

Not independently runnable (calls `Variant.new`'s explicit path, which allocates real VARIANT byte buffers via Task 1/2's pure logic — actually locally exercisable at `require` time! Re-run `ruby -Ilib -Itest test/win32ole/jruby/test_variant.rb` after adding this: if `variant.rb` fails to even `require` cleanly because these four lines raise, that's a real, catchable bug this smoke check would surface immediately, unlike Record's construction path which needs live COM.)

- [ ] **Step 3: Smoke check**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_variant.rb`
Expected: still all PASS (confirms `variant.rb` requires cleanly end-to-end, including the four eager constants).

---

## Task 15: `win32ole.rb` — dispatch hook (`Array`/`Record`/`Variant` marshaling)

**Files:**
- Modify: `lib/win32ole/jruby/win32ole.rb`

**Interfaces:**
- Consumes: everything from Tasks 1-14.
- Produces: extended `WIN32OLE.ruby_value_to_variant_bytes`/`.variant_bytes_to_ruby_value` (Task 3's class methods) — the integration point tying the whole phase together.

This is the one task where Phase 1/2's existing, CI-green dispatch path is directly touched — the new branches must come *before* the existing `else raise TypeError`/`NotImplementedError` fallthroughs, not replace them, so every Phase 1/2 behavior is unchanged for values this phase doesn't add handling for.

- [ ] **Step 1: Implement**

Add `require 'win32ole/jruby/array'` and `require 'win32ole/jruby/record'` and `require 'win32ole/jruby/variant'` to the top of `lib/win32ole/jruby/win32ole.rb`, alongside the existing `require 'win32ole/jruby/win32'`/`require 'win32ole/jruby/dispatch'`.

Replace the `class << self` block's two methods (from Task 3) with:

```ruby
  class << self
    def wrap_dispatch_pointer(ptr)
      obj = allocate
      obj.instance_variable_set(:@ptr, ptr)
      obj.send(:install_finalizer)
      obj
    end

    def ruby_value_to_variant_bytes(value, bstrs_to_free)
      case value
      when ::Array
        psa = SafeArray.ruby_array_to_safearray(value, W::VT_VARIANT, bstrs_to_free)
        return W.pack_variant(W::VT_VARIANT | W::VT_ARRAY, W.pack_pointer(psa.to_i))
      when WIN32OLE::Record
        return value.to_variant_bytes
      when WIN32OLE::Variant
        return value.instance_variable_get(:@var)
      end

      type = W.ruby_to_variant_type(value)
      payload =
        case type
        when :i4 then W.pack_i4(value)
        when :i8 then W.pack_i8(value)
        when :r8 then W.pack_r8(value)
        when :bool then W.pack_bool(value)
        when :empty then W.pack_empty
        when :bstr
          bstr = W.sys_alloc_string.call(W.wstr(value))
          bstrs_to_free << bstr
          W.pack_pointer(bstr)
        when :dispatch
          W.pack_pointer(value.instance_variable_get(:@ptr))
        end
      W.pack_variant(W::VT_FOR_TYPE.fetch(type), payload)
    end

    def variant_bytes_to_ruby_value(bytes)
      vt, = W.unpack_variant(bytes)
      base_vt = vt & W::VT_TYPEMASK

      if (vt & W::VT_ARRAY) != 0
        _vt, payload = W.unpack_variant(bytes)
        psa = W.unpack_pointer(payload)
        return SafeArray.safearray_to_ruby_array(psa, base_vt)
      end

      if base_vt == W::VT_RECORD
        _vt, body = W.unpack_variant(bytes, body_size: WIN32OLE::Record::VT_RECORD_BODY_SIZE)
        buffer_ptr, pri = body.unpack("#{W::PACK_PTR}2")
        return WIN32OLE::Record.from_irecordinfo_and_buffer(pri, buffer_ptr)
      end

      type = W.variant_ruby_type(vt)
      _vt2, payload = W.unpack_variant(bytes)
      case type
      when :empty then nil
      when :i4 then W.unpack_i4(payload)
      when :i8 then W.unpack_i8(payload)
      when :r8 then W.unpack_r8(payload)
      when :bool then W.unpack_bool(payload)
      when :bstr
        addr = W.unpack_pointer(payload)
        str = W.bstr_to_s(addr)
        W.sys_free_string.call(addr) unless addr.zero?
        str
      when :dispatch
        ptr = W.unpack_pointer(payload)
        ptr.zero? ? nil : wrap_dispatch_pointer(ptr)
      end
    end
  end
```

`Win32.variant_ruby_type` (Phase 1, unchanged) still raises `NotImplementedError` for any `base_vt` this phase doesn't recognize — `VT_ARRAY`/`VT_RECORD` are handled *before* that call is even reached above, so `variant_ruby_type` itself needs no new `case` branches; this is Phase 1 §6.4's own table actually growing two covered entries the way that section's framing anticipated (spec §4.6's closing point), not a change to `variant_ruby_type`.

- [ ] **Step 2: Run local smoke check**

Run: `bundle exec rake test`
Expected: same pass/fail shape as before this task (pure-logic-reachable parts unaffected; everything else is CI-only, confirmed in Task 17).

- [ ] **Step 3: Commit**

```bash
git add lib/win32ole/jruby/win32ole.rb
git commit -m "jruby: dispatch hook -- Array/Record/Variant argument marshaling, VT_ARRAY/VT_RECORD result unmarshaling"
```

---

## Task 16: `GC.stress` finalizer test + structural `VT_BYREF` round-trip test

**Files:**
- Create: `test/win32ole/jruby/test_record_variant_gc_stress.rb` (`RUBY_ENGINE == 'jruby'`-guarded, `defined?(WIN32OLE)`-guarded, matching `test_typelib_gc_stress.rb`'s exact shape)
- Test additions: `test/win32ole/jruby/test_variant.rb` (structural `VT_BYREF` test — this part is Windows-free, add it to the existing `RUBY_ENGINE == 'jruby'` guarded file, not the GC-stress file)

**Interfaces:** none new.

Per spec §7 and this plan's header finding: the `GC.stress` finalizer probe needs a live COM object (CI-only, and per the header finding, has no `Record`-capable fixture — so it exercises `Variant`'s `SAFEARRAY` finalizer path via `Scripting.Dictionary`, not `Record`'s). The `VT_BYREF` round-trip is downgraded from spec §7's "live OLE out-parameter call" to a structural, locally-runnable pointer-aliasing check (consistent with the `Record` decision in this plan's header) — it confirms `realvar`/`var` stay correctly linked at the byte level without needing a live out-parameter call this project's CI has no fixture for.

- [ ] **Step 1: Add the structural `VT_BYREF` round-trip test**

Append to `test/win32ole/jruby/test_variant.rb` (inside the existing `if RUBY_ENGINE == 'jruby'` guard):

```ruby
class TestVariantByRefRoundTrip < Test::Unit::TestCase
  def test_mutating_realvar_bytes_is_visible_through_the_byref_pointer
    v = WIN32OLE::Variant.new(42, WIN32OLE::VariantType::VT_I4 | WIN32OLE::VariantType::VT_BYREF)
    realvar = v.instance_variable_get(:@realvar)
    var = v.instance_variable_get(:@var)

    _vt, payload = WIN32OLE::Win32.unpack_variant(var)
    ptr_into_realvar = WIN32OLE::Win32.unpack_pointer(payload)
    assert_equal(WIN32OLE::Win32.native_pointer_for(realvar).to_i + 8, ptr_into_realvar)

    # Simulate an out-parameter callee overwriting *ptr_into_realvar in place
    Fiddle::Pointer.new(ptr_into_realvar)[0, 4] = [99].pack('l')
    assert_equal(99, WIN32OLE::Win32.unpack_i4(realvar[8, 4]))
  end
end
```

- [ ] **Step 2: Run test to verify it passes**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_variant.rb`
Expected: PASS. This is a genuine, locally-runnable confirmation of the `realvar`/`var` aliasing invariant (spec §4.7) — it doesn't need Windows because it's pure in-process pointer arithmetic on a Ruby `String`'s own bytes, same technique as Task 2's `pack_byref` tests.

- [ ] **Step 3: Add the `GC.stress` finalizer test**

```ruby
# test/win32ole/jruby/test_record_variant_gc_stress.rb
begin
  require 'win32ole'
rescue LoadError
end
require 'test/unit'

if defined?(WIN32OLE) && RUBY_ENGINE == 'jruby'
  class TestRecordVariantGCStress < Test::Unit::TestCase
    def test_gc_stress_survives_repeated_variant_array_construction
      dict = WIN32OLE.new('Scripting.Dictionary')
      dict.add('a', 1)
      dict.add('b', 2)
      GC.stress = true
      50.times do
        keys = dict.Keys
        items = dict.Items
        assert_kind_of(Array, keys)
        assert_kind_of(Array, items)
      end
    ensure
      GC.stress = false
    end
  end
end
```

- [ ] **Step 4: Commit and push**

```bash
git add test/win32ole/jruby/test_variant.rb test/win32ole/jruby/test_record_variant_gc_stress.rb
git commit -m "test: structural VT_BYREF round-trip + GC.stress probe for SAFEARRAY finalizer path"
```

(Confirm with whoever is driving this plan's execution before pushing — same established practice this session as Phase 2's plan noted: pushing to `origin/jruby-support` triggers real GitHub Actions billing.)

---

## Task 17: CI verification against the legacy test suite

**Files:**
- None created — this task pushes Tasks 1-16's work to CI and fixes real bugs the existing test suite finds.

**Interfaces:** none new.

- [ ] **Step 1: Push and read the CI log**

Run: `git push origin jruby-support` (only after explicit confirmation, per Task 16's note), then `gh run list --branch jruby-support --limit 1`, then `gh run view <run-id> --job=<jruby-job-id> --log-failed`.

- [ ] **Step 2: Triage every failure in `test_win32ole_record.rb`/`test_win32ole_variant.rb`/`test_win32ole_variant_m.rb`/`test_win32ole_variant_outarg.rb`**

Classify each:

- **Expected — `RbComTest.ComSrvTest` not installed**: per this plan's header finding, this test environment has no such server registered. These tests self-`omit` (`test/win32ole/test_win32ole_record.rb:79-86` and the equivalent guards in the other three files) — confirm they show as *omissions*, not failures or passes. This covers nearly all of `test_win32ole_record.rb` and the struct-passing parts of `test_win32ole_variant.rb`/`test_win32ole_variant_outarg.rb`.
- **Expected (§3 non-goal)**: anything exercising `VT_CY`/`VT_DATE` conversion. Confirm it raises `NotImplementedError`, not a silent wrong value.
- **Real bug**: anything else — including any `test_win32ole_variant_m.rb` constant-table mismatch, or any `Scripting.Dictionary#Keys`/`#Items`-based array test failure (this plan's one genuinely CI-verifiable live fixture — a failure here is a real bug in Task 4-6's `SAFEARRAY` marshaling, not an environment gap). Fix it, re-reading the relevant `oleaut32` signature/struct field against Microsoft's documentation and `ext/win32ole/win32ole.c`/`win32ole_variant.c` (present in this repo), commit, push, and re-check CI.

- [ ] **Step 3: Confirm zero regressions in the pre-existing Phase 1/2 test set**

Diff the new failure list against Phase 2's last known-good baseline the same way Phase 2's own CI-driven fix rounds did (its plan's Task 11, Step 4) — any Phase-1/2-scoped test that was passing before this phase's changes and now fails is a real regression this phase introduced (most likely via the Task 3 refactor or a shared-file edit to `win32.rb`/`win32ole.rb`).

- [ ] **Step 4: Final commit once CI is clean**

```bash
git add -A
git commit -m "jruby: Phase 3 fixes from CI triage" # only if Step 2 produced fixes beyond earlier commits
```

- [ ] **Step 5: Record, explicitly, what CI green here does and doesn't mean**

Green CI at this point confirms: `Variant`'s scalar + `SAFEARRAY` (1-D, `VT_VARIANT`-element) marshaling against a real live fixture, the full constant table, and that nothing in Phase 1/2's own test set regressed. It does **not** confirm `Record`/`IRecordInfo`'s vtable slots, `PutField`/`GetFieldNoCopy` call shapes, N-dimensional (2-D+) `SAFEARRAY` marshaling, or a live `VT_BYREF` out-parameter round trip — those remain open per this plan's header finding and design spec §8 risk #1, carried forward exactly as decided, not silently resolved by a green build.

---

## Self-Review

**Spec coverage:** §4.1 (file layout) → Tasks 4-15 (one task group per file). §4.2 (wider VARIANT substrate) → Tasks 1-2. §4.3 (`array.rb`) → Tasks 4-6. §4.4 (`record.rb`) → Tasks 7-10 (with the construction-timing correction found by reading `win32ole_record.c` directly, flagged explicitly rather than silently matching the spec's own summary). §4.5 (`variant.rb`) → Tasks 11-14. §4.6 (dispatch hook) → Task 15. §4.7 (resource lifetime) → each task's own construction/finalizer code (Task 8/10 for `Record`, Task 12/13 for `Variant`). §5 (per-class API table) → cross-checked one-by-one: `Record`'s `new`/`to_h`/`typename`/`method_missing`/`ole_instance_variable_get`/`_set`/`inspect` (Tasks 8-9) match the table exactly; `Variant`'s `new`/`.array`/`value`/`value=`/`vartype`/`[]`/`[]=` (Tasks 12-13) match exactly; `WIN32OLE` dispatch additions (Task 15) match. §6 (error translation) → `WIN32OLE::RuntimeError` at every `Record`/`SAFEARRAY` failure site (Tasks 6, 8, 10), `ArgumentError` for `Variant.new(val, VT_RECORD)` (Task 12), `KeyError` via `Hash#fetch` for unknown `Record` fields (Task 9), plain `RuntimeError` (not `WIN32OLE::RuntimeError`) for a `NULL` `SafeArrayCreate` return (Tasks 6, 13). §7 (testing/CI) → Task 5 (local N-dim unit tests), Task 16 (`GC.stress` + structural `VT_BYREF`), Task 17 (legacy suite + triage). §8 risks: #1 (`IRecordInfo` vtable unverified) → explicitly *not* resolved, stated in the plan header and Task 7/8/10/17's own text, not glossed over; #2 (n/a — no new struct layouts this phase, `SAFEARRAY` is never hand-packed per Task 4); #3 (N-dim algorithm fidelity) → Task 5's direct line-cited port, with the fill/read asymmetry finding documented; #4 (`olerecord_free` no `RecordClear`) → not applicable since this port never even reaches an allocated buffer in most paths, and Task 10's `to_variant_bytes` matches `ole_rec2variant`'s own "no `RecordClear`, just `free`" shape via the finalizer added in Task 8; #5 (x86) → inherited, not re-litigated; #6 (performance) → inherited, not re-litigated.

**Placeholder scan:** no TBD/TODO/"implement later"/"add appropriate error handling" phrasing anywhere in this plan. An earlier draft of Task 10's `to_variant_bytes` contained a syntax error in its `PutField`-failure branch and an errant trailing comma in its final `pack_variant` call — both were corrected in place before being written into this document, consistent with Phase 1/2's own precedent of showing only the final, correct code rather than a wrong-then-fixed sequence.

**Type consistency:** `Record.new(typename, oleobj)` (Task 8) → `resolve_itypelib_ptr` accepts exactly `WIN32OLE`/`WIN32OLE::TypeLib`, matching Task 15's dispatch hook which only ever constructs a `Record` via `.from_irecordinfo_and_buffer` (Task 10), never `.new`, for a dispatch *result* — the two construction paths are never confused at a call site. `SafeArray.ruby_array_to_safearray(ary, elem_vt, bstrs_to_free = [])`/`.safearray_to_ruby_array(psa, elem_vt)` (Task 6) match their call sites in `Variant#initialize`/`#value=`/`.array` (Task 12-13, though `.array`/`#[]`/`#[]=` build/read `SAFEARRAY`s directly via Task 4's bindings rather than through Task 6's whole-array functions — a deliberate, noted-in-Task-13 difference, not an inconsistency) and in `WIN32OLE.ruby_value_to_variant_bytes`/`.variant_bytes_to_ruby_value` (Task 15). `Win32.pack_variant(vt, body)`/`.unpack_variant(bytes, body_size: 8)` (Task 1) signatures match every call site across Tasks 2, 6, 10, 12, 15 — no call site passes a keyword argument to `pack_variant` (it never gained one) or a positional `body_size` to `unpack_variant` (it's keyword-only).
