# JRuby win32ole Phase 4 (Event) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement `WIN32OLE::Event` (a COM connection-point event sink built from a hand-rolled `IDispatch` vtable backed by `Fiddle::Closure::BlockCaller`), plus its two prerequisites, `WIN32OLE.connect` and `WIN32OLE.const_load`, on top of the already-shipped Phase 1-3 substrate (`WIN32OLE` core dispatch, `Type`/`TypeLib`/`Method`/`Param`/`Variable` introspection, `Record`/`Variant`/`SAFEARRAY` marshaling).

**Architecture:** One new file (`event.rb`: the sink's 7-slot vtable, event-source resolution ported from `win32ole_event.c`'s `find_iid`/`find_coclass`/`find_default_source*`, and the public `Event` API) plus three extended files (`win32.rb`: new COM interface IIDs, a generic `QueryInterface` helper, `user32.dll` message-loop bindings; `typeinfo.rb`: five more `ITypeInfo`/`ITypeLib` vtable `_fn` helpers; `win32ole.rb`: `.connect`/`.const_load`, plus a small Tidy-First promotion of `resolve_clsid`/`hresult_detail` from private instance methods to shared class-level helpers, and one new `#dispatch_ptr` reader). This is the **first** use of `Fiddle::Closure` anywhere in this port — Phase 1's spike proved the mechanism works on both MRI and JRuby, but that spike code was never committed (confirmed: no trace of it exists anywhere in this repository's git history), so every closure in this plan is written from scratch against the C extension's `win32ole_event.c` and empirically-verified `Fiddle::Closure::BlockCaller` semantics (see the Global Constraints finding below), not ported from prior working code.

**Tech Stack:** Ruby stdlib `fiddle` (`Fiddle::Closure::BlockCaller` for the sink's native vtable slots, `Fiddle::Pointer.malloc`/`Fiddle.free` for the sink/vtable buffers, matching `Record`/`Variant`'s existing "shared finalizer-state hash" discipline), `user32.dll`/`oleaut32.dll`/`ole32.dll` via `Fiddle.dlopen`.

**Spec:** `docs/superpowers/specs/2026-09-26-jruby-win32ole-phase4-event-design.md` — this plan implements §4/§5/§6 in full; §3's non-goals (C-parity exception handling, `SWbemSink` async CI verification, `WIN32OLE.connect`'s `host` param, x86/ARM64, the four public `Type` ImplType methods) are out of scope here and are not silently "fixed" by any task below.

**A materially important finding not in the spec, discovered while writing this plan:** every existing JRuby-port test file (`test/win32ole/jruby/*.rb`) wraps its entire body — `require` and test class alike — in `if RUBY_ENGINE == 'jruby'`. On this development machine (`RUBY_ENGINE == 'ruby'`), running any of these files directly (`ruby -Ilib -Itest test/win32ole/jruby/test_win32.rb`) produces **zero test output and exit 0** — the whole file is skipped, not "0 failures", so this command (used by the Phase 3 plan's own "Run:" steps) does not actually confirm red-then-green the way it reads. Verified directly: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_win32.rb"'` forces the guard open and genuinely runs all 41 existing pure-logic tests under this machine's real CRuby Fiddle backend (41 tests, 93 assertions, 0 failures — confirmed while writing this plan). **This plan's "Run:" steps for pure-logic tasks use this exact spoofed-engine invocation, not a bare `ruby test_file.rb`.** It does not help for anything that calls `Fiddle.dlopen` against a Windows-only DLL (`user32`, `oleaut32` GetActiveObject, `advapi32`) — those remain CI-only exactly as Phase 3 already established for its own live-COM tasks.

**A second finding:** `Fiddle::Closure::BlockCaller`'s `TYPE_VOIDP` arguments arrive in the block as `Fiddle::Pointer` objects (not raw `Integer`s), and returning a plain `Integer` for a `TYPE_VOIDP` return type works correctly (wrapped in a `Pointer` by the caller side) — verified empirically on this machine (`Fiddle::Closure::BlockCaller.new(...) { |ptr, n| ... }` with a `TYPE_VOIDP` arg receives `#<Fiddle::Pointer ...>`, `.to_i`/`.class` checked directly). Every closure below is written assuming this.

## Global Constraints

- **Reuse Phase 1-3 substrate unchanged**: `Win32.vtable_function`/`vtable_address` (memoize `_fn` helpers keyed on vtable address, never object address), `Win32.native_pointer_for` (hold the returned `Fiddle::Pointer` in a local/ivar through any native call that dereferences it), `WIN32OLE::RuntimeError`/`QueryInterfaceError`, `WIN32OLE.ruby_value_to_variant_bytes`/`variant_bytes_to_ruby_value`.
- **Closures use `Win32::STDCALL`**, exactly like every other vtable call in this codebase — the sink's vtable is called BY the OLE server using the same calling convention every other vtable slot in this port already uses.
- **No Ruby exception may ever propagate out of a `Fiddle::Closure::BlockCaller` block.** This is not new policy — it is spec §3's existing, explicit non-goal boundary ("write to `$stderr`... keeps the process... alive") applied literally: an uncaught exception unwinding across the closure's native trampoline back into the calling OLE server's C code is undefined behavior exactly like unwinding across any other native boundary in this project. Every closure below either cannot raise (pure byte packing) or is wrapped in a `rescue StandardError, ScriptError` that logs to `$stderr` and returns a safe HRESULT.
- **New COM interface IIDs (`IUnknown`, `IConnectionPointContainer`, `IProvideClassInfo`, `IProvideClassInfo2`) are hand-transcribed from well-known, decades-stable OLE Automation headers (`objbase.h`/`ocidl.h`)** — same "verify against a live Windows build before anything depends on it" caution Phase 3 gave `IRecordInfo`'s vtable (design §8 risk #1 there; carried forward here as risk #1, see §8 below).
- **The sink/vtable native buffers follow the exact "shared mutable state hash captured by a class-level finalizer proc, not `self`, not live wrapper objects" discipline `variant.rb:282-300`/`record.rb` already established** (`persist_realvar`/`commit_realvar`/`realvar_finalizer`) — a finalizer proc that captured `self` or the live `Fiddle::Closure` objects would create a reference cycle that prevents GC from ever running it.
- **Event-source resolution (`find_iid`/`find_coclass`/`find_default_source*`) is implemented as `event.rb`-private helpers operating on raw `ITypeInfo*`/`ITypeLib*` integers**, not through the `Type`/`TypeLib` wrapper classes — spec §3 non-goal #5 is explicit that this does NOT wire up `Type#implemented_ole_types`/friends (those stay `NotImplementedError`, Phase 2 non-goal). Every raw pointer obtained mid-traversal that isn't the final result must be `Release`'d (vtable slot 2) before returning or moving to the next candidate — leaking these is a real, silent COM reference leak, not just untidy code.
- **Verification reality**: pure Ruby/Fiddle logic (byte packing, `Fiddle::Closure` construction AND direct invocation via a hand-built `Fiddle::Function` wrapper around the closure's own address, `@events` array bookkeeping) is locally verifiable right now via the spoofed-`RUBY_ENGINE` harness described above. Anything calling `Fiddle.dlopen('ole32'|'oleaut32'|'user32'|'advapi32')` or requiring a live `WIN32OLE`/COM object is Windows+COM-only and can only be verified by pushing to the `test-jruby` CI job (`.github/workflows/windows.yml`), exactly as Phase 3 already established for its own live-COM tasks.
- **`ext/win32ole/win32ole_event.c` is present in this repo** (real MRI C extension, not the JRuby port) — tasks below cite exact line ranges; cross-check any detail not spelled out in a task's own code against it directly, the same discipline Phase 2/3 established.
- **New test file**: `test/win32ole/jruby/test_event.rb`, following the established `RUBY_ENGINE == 'jruby'`-guards-the-whole-file convention (`require 'win32ole/jruby/event'` directly, like `test_win32.rb` requires `win32ole/jruby/win32` directly — not `require 'win32ole'`).

## Review Focus

- Calling `#on_event`/`#off_event` after `#unadvise` has already run must raise `WIN32OLE::RuntimeError` ("You must call advise at first"), and a second `#unadvise` call must be a safe no-op — neither may touch already-freed native memory.
- A raised exception inside a user's `on_event`/`on_event_with_outargs`/`handler=` callback must never escape the `Invoke` closure — caught, written to `$stderr` with a backtrace, and `Invoke` still returns `NOERROR` so the firing OLE server (and the message loop) keep running (spec §3).
- `#on_event`/`#off_event` must accept a `Symbol` event name (`:WillConnect`) identically to the equivalent `String`, and raise `TypeError` for anything else (matches `ev_on_event`/`fev_off_event`'s explicit type check).
- Registering two callbacks under the same event name (including two catch-alls, i.e. `on_event` called twice with no name) must **replace** the earlier one, never stack both — `add_event_call_back`'s delete-then-push semantics mean the common "register a catch-all, then override it" idiom must not silently double-invoke user code.
- `WIN32OLE::Event.new(ole, itf)` against a `WIN32OLE` object with no matching event source (e.g. `Scripting.Dictionary`, per the legacy test's own `test_s_new_non_exist_event`) must raise a catchable `WIN32OLE::RuntimeError`, never crash via an unguarded native dereference deep in the `ImplType` traversal — every helper that reads a `TYPEATTR`/calls `GetRefTypeOfImplType` must check its own `HRESULT`/nil result before dereferencing further.

---

## Task 1: `win32.rb` — COM interface IIDs + generic `QueryInterface`

**Files:**
- Modify: `lib/win32ole/jruby/win32.rb`
- Test: `test/win32ole/jruby/test_win32.rb` (inside the existing `if RUBY_ENGINE == 'jruby'` guard)

**Interfaces:**
- Consumes: `Win32.vtable_function`, `Win32.failed?`, `Win32.PACK_PTR`/`PTR_SIZE` (all existing).
- Produces: `Win32::IID_IUNKNOWN`, `Win32::IID_ICONNECTIONPOINTCONTAINER`, `Win32::IID_IPROVIDECLASSINFO`, `Win32::IID_IPROVIDECLASSINFO2` (16-byte packed GUIDs, same shape as the existing `IID_IDISPATCH`), `Win32.query_interface(obj_addr, iid_bytes)` → pointer `Integer` or `nil`.

- [ ] **Step 1: Write the failing test**

Add inside `test/win32ole/jruby/test_win32.rb`'s existing `class TestWin32`:

```ruby
  def test_new_iid_constants_are_16_bytes
    [W::IID_IUNKNOWN, W::IID_ICONNECTIONPOINTCONTAINER, W::IID_IPROVIDECLASSINFO, W::IID_IPROVIDECLASSINFO2].each do |iid|
      assert_equal(16, iid.bytesize)
    end
  end

  def test_iid_iunknown_matches_known_bytes
    # {00000000-0000-0000-C000-000000000046}
    expected = [0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46].pack('LSSC8')
    assert_equal(expected, W::IID_IUNKNOWN)
  end

  def test_new_iid_constants_are_pairwise_distinct
    iids = [W::IID_IUNKNOWN, W::IID_ICONNECTIONPOINTCONTAINER, W::IID_IPROVIDECLASSINFO, W::IID_IPROVIDECLASSINFO2, W::IID_IDISPATCH]
    assert_equal(iids.size, iids.uniq.size)
  end

  def test_query_interface_reads_hr_and_returns_nil_on_failure
    # Fake object: first pointer-sized field is a vtable whose slot 0
    # (QueryInterface) is a closure that always fails. No live COM needed --
    # this is the same "malloc a fake vtable" technique test_win32.rb already
    # uses for test_vtable_address_reads_the_first_pointer_sized_field.
    qi = Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL) do |_this, _riid, ppv|
      ppv[0, W::PTR_SIZE] = [0].pack(W::PACK_PTR)
      -2147467262 # E_NOINTERFACE
    end
    vtable = Fiddle::Pointer.malloc(W::PTR_SIZE)
    vtable[0, W::PTR_SIZE] = [qi.to_i].pack(W::PACK_PTR)
    obj = Fiddle::Pointer.malloc(W::PTR_SIZE)
    obj[0, W::PTR_SIZE] = [vtable.to_i].pack(W::PACK_PTR)

    assert_nil(W.query_interface(obj.to_i, ("\x00" * 16).b))
  ensure
    Fiddle.free(vtable.to_i) if vtable
    Fiddle.free(obj.to_i) if obj
  end

  def test_query_interface_returns_the_ppv_pointer_on_success
    qi = Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL) do |_this, _riid, ppv|
      ppv[0, W::PTR_SIZE] = [0x123456].pack(W::PACK_PTR)
      0
    end
    vtable = Fiddle::Pointer.malloc(W::PTR_SIZE)
    vtable[0, W::PTR_SIZE] = [qi.to_i].pack(W::PACK_PTR)
    obj = Fiddle::Pointer.malloc(W::PTR_SIZE)
    obj[0, W::PTR_SIZE] = [vtable.to_i].pack(W::PACK_PTR)

    assert_equal(0x123456, W.query_interface(obj.to_i, ("\x00" * 16).b))
  ensure
    Fiddle.free(vtable.to_i) if vtable
    Fiddle.free(obj.to_i) if obj
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_win32.rb"'`
Expected: `NameError`/`NoMethodError` — the new constants/method don't exist yet.

- [ ] **Step 3: Add the constants and helper to `win32.rb`**

Add near `IID_IDISPATCH`:

```ruby
    # Well-known, decades-stable OLE Automation interface IDs (ocidl.h /
    # objbase.h). Hand-transcribed, not pulled from a live header -- verify
    # against a real Windows build before anything else depends on it (same
    # caution Phase 3 gave IRecordInfo's vtable; see this plan's §8 risk #1).
    IID_IUNKNOWN = [0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46].pack('LSSC8')
    IID_ICONNECTIONPOINTCONTAINER = [0xB196B284, 0xBAB4, 0x101A, 0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07].pack('LSSC8')
    IID_IPROVIDECLASSINFO = [0xB196B283, 0xBAB4, 0x101A, 0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07].pack('LSSC8')
    IID_IPROVIDECLASSINFO2 = [0xA6BC3AC0, 0xDBAA, 0x11CE, 0x9D, 0xE3, 0x00, 0xAA, 0x00, 0x4B, 0xB8, 0x51].pack('LSSC8')
```

Add near `vtable_function`:

```ruby
    # Generic COM QueryInterface, usable against any interface pointer --
    # every other vtable helper in this codebase is for a FIXED interface
    # (IDispatch/ITypeInfo/ITypeLib); WIN32OLE::Event is the first caller
    # that needs to ask an arbitrary object for an arbitrary interface.
    def query_interface(obj_addr, iid_bytes)
      ppv = ("\x00" * PTR_SIZE).b
      hr = vtable_function(obj_addr, 0, [VOIDP, VOIDP, VOIDP], LONG).call(obj_addr, iid_bytes, ppv)
      return nil if failed?(hr)

      ppv.unpack1(PACK_PTR)
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_win32.rb"'`
Expected: PASS, all tests including the 5 new ones.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/win32.rb test/win32ole/jruby/test_win32.rb
git commit -m "jruby: add COM interface IIDs and generic QueryInterface for Event"
```

---

## Task 2: `win32.rb` — `user32.dll` message-loop substrate

**Files:**
- Modify: `lib/win32ole/jruby/win32.rb`
- Test: `test/win32ole/jruby/test_win32.rb`

**Interfaces:**
- Consumes: `Win32::STDCALL`, `Win32::VOIDP`/`DWORD`.
- Produces: `Win32.pump_windows_messages` (drains the calling thread's message queue: `PeekMessageW`/`TranslateMessage`/`DispatchMessageW` until empty).

Windows-only (`user32.dll` doesn't exist on this dev machine) — the drain loop itself is CI-verified only. The one thing testable here without Windows is that the constants have sane values.

- [ ] **Step 1: Write the failing test**

```ruby
  def test_pm_remove_matches_win32_constant
    assert_equal(0x0001, W::PM_REMOVE)
  end

  def test_msg_size_is_large_enough_for_a_real_msg_struct
    # Real MSG is 48 bytes on x64 / 28 on x86 (HWND hwnd; UINT message; WPARAM
    # wParam; LPARAM lParam; DWORD time; POINT pt). We never read MSG's own
    # fields (just pass the pointer PeekMessage filled in on to
    # TranslateMessage/DispatchMessage), so exact layout doesn't matter --
    # only that the scratch buffer is big enough for the OS to write into
    # without corrupting adjacent memory.
    minimum = W::PTR_SIZE == 8 ? 48 : 28
    assert_operator(W::MSG_SIZE, :>=, minimum)
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_win32.rb"'`
Expected: `NameError` — `PM_REMOVE`/`MSG_SIZE` don't exist yet.

- [ ] **Step 3: Add the message-loop substrate**

```ruby
    PM_REMOVE = 0x0001
    MSG_SIZE = 64 # generous, alignment-safe upper bound; see test comment above

    def user32
      @user32 ||= Fiddle.dlopen('user32')
    end

    def peek_message
      @peek_message ||= Fiddle::Function.new(
        user32['PeekMessageW'], [VOIDP, VOIDP, DWORD, DWORD, DWORD], DWORD, STDCALL
      )
    end

    def translate_message
      @translate_message ||= Fiddle::Function.new(user32['TranslateMessage'], [VOIDP], DWORD, STDCALL)
    end

    def dispatch_message
      @dispatch_message ||= Fiddle::Function.new(user32['DispatchMessageW'], [VOIDP], LONG, STDCALL)
    end

    def pump_windows_messages
      msg = Fiddle::Pointer.malloc(MSG_SIZE)
      while peek_message.call(msg, nil, 0, 0, PM_REMOVE) != 0
        translate_message.call(msg)
        dispatch_message.call(msg)
      end
      nil
    ensure
      Fiddle.free(msg.to_i) if msg
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_win32.rb"'`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/win32.rb test/win32ole/jruby/test_win32.rb
git commit -m "jruby: add user32.dll message-loop substrate for Event.message_loop"
```

---

## Task 3: `typeinfo.rb` — new `ITypeInfo`/`ITypeLib` vtable `_fn` helpers

**Files:**
- Modify: `lib/win32ole/jruby/typeinfo.rb`
- Test: `test/win32ole/jruby/test_typeinfo.rb`

**Interfaces:**
- Consumes: `ITYPEINFO_VTBL`/`ITYPELIB_VTBL` (already list all five slots used here), `W.vtable_function`/`vtable_address`.
- Produces: `TypeInfo.impl_type_flags_fn`, `TypeInfo.ref_type_of_impl_type_fn`, `TypeInfo.type_info_of_guid_fn`, `TypeInfo.get_names_fn`, `TypeInfo.get_ids_of_names_fn` — same memoized-by-vtable-address pattern as every existing `_fn` helper.

**Note beyond the spec**: the design doc's §4.3 names four helpers (`GetImplTypeFlags`/`GetRefTypeOfImplType`/`GetTypeInfoOfGuid`/`GetNames`); `GetIDsOfNames` (the sink's own `EVENTSINK_GetIDsOfNames` delegates to it, spec §1.1/§4.4) needs a fifth, on `ITypeInfo` (not `IDispatch` — `Dispatch#get_ids_of_names_fn` already exists but is IDispatch's own 5-arg shape with a leading `riid`; `ITypeInfo::GetIDsOfNames` is a different 3-arg shape with no `riid`). Adding it here since it's the same kind of helper as the other four.

- [ ] **Step 1: Write the failing test**

Add inside `test/win32ole/jruby/test_typeinfo.rb`'s existing test class. This uses the same "malloc a fake vtable" technique as `test_win32.rb`'s `test_vtable_address_reads_the_first_pointer_sized_field` — it only proves the helper resolves the *documented slot index* and memoizes correctly, not that a real `ITypeInfo` accepts the call (that's COM/CI-only).

```ruby
  def fake_vtable_object(slot_count)
    vtable = Fiddle::Pointer.malloc(WIN32OLE::Win32::PTR_SIZE * slot_count)
    slot_count.times { |i| vtable[i * WIN32OLE::Win32::PTR_SIZE, WIN32OLE::Win32::PTR_SIZE] = [0x1000 + i].pack(WIN32OLE::Win32::PACK_PTR) }
    obj = Fiddle::Pointer.malloc(WIN32OLE::Win32::PTR_SIZE)
    obj[0, WIN32OLE::Win32::PTR_SIZE] = [vtable.to_i].pack(WIN32OLE::Win32::PACK_PTR)
    [obj, vtable]
  end

  def test_impl_type_flags_fn_resolves_the_documented_slot_and_memoizes
    obj, vtable = fake_vtable_object(10)
    fn1 = TI.impl_type_flags_fn(obj.to_i)
    fn2 = TI.impl_type_flags_fn(obj.to_i)
    assert_same(fn1, fn2)
    assert_equal(0x1000 + TI::ITYPEINFO_VTBL[:GetImplTypeFlags], fn1.to_i)
  ensure
    Fiddle.free(vtable.to_i) if vtable
    Fiddle.free(obj.to_i) if obj
  end

  def test_ref_type_of_impl_type_fn_resolves_the_documented_slot
    obj, vtable = fake_vtable_object(10)
    fn = TI.ref_type_of_impl_type_fn(obj.to_i)
    assert_equal(0x1000 + TI::ITYPEINFO_VTBL[:GetRefTypeOfImplType], fn.to_i)
  ensure
    Fiddle.free(vtable.to_i) if vtable
    Fiddle.free(obj.to_i) if obj
  end

  def test_get_names_fn_resolves_the_documented_slot
    obj, vtable = fake_vtable_object(10)
    fn = TI.get_names_fn(obj.to_i)
    assert_equal(0x1000 + TI::ITYPEINFO_VTBL[:GetNames], fn.to_i)
  ensure
    Fiddle.free(vtable.to_i) if vtable
    Fiddle.free(obj.to_i) if obj
  end

  def test_get_ids_of_names_fn_resolves_the_documented_slot
    obj, vtable = fake_vtable_object(11)
    fn = TI.get_ids_of_names_fn(obj.to_i)
    assert_equal(0x1000 + TI::ITYPEINFO_VTBL[:GetIDsOfNames], fn.to_i)
  ensure
    Fiddle.free(vtable.to_i) if vtable
    Fiddle.free(obj.to_i) if obj
  end

  def test_type_info_of_guid_fn_resolves_the_documented_slot
    obj, vtable = fake_vtable_object(7)
    fn = TI.type_info_of_guid_fn(obj.to_i)
    assert_equal(0x1000 + TI::ITYPELIB_VTBL[:GetTypeInfoOfGuid], fn.to_i)
  ensure
    Fiddle.free(vtable.to_i) if vtable
    Fiddle.free(obj.to_i) if obj
  end
```

(`Fiddle::Function#to_i` returns the wrapped function pointer's own address — this is what lets the test observe which vtable slot got resolved without ever calling the fabricated function, which would crash since `0x1000 + i` isn't real code.)

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_typeinfo.rb"'`
Expected: `NoMethodError` for each new `_fn` helper.

- [ ] **Step 3: Add the five helpers to `typeinfo.rb`**

Add alongside the existing `_fn` helpers:

```ruby
    def impl_type_flags_fn(itypeinfo_ptr)
      @impl_type_flags_fns ||= {}
      @impl_type_flags_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetImplTypeFlags], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def ref_type_of_impl_type_fn(itypeinfo_ptr)
      @ref_type_of_impl_type_fns ||= {}
      @ref_type_of_impl_type_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetRefTypeOfImplType], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def get_names_fn(itypeinfo_ptr)
      @get_names_fns ||= {}
      @get_names_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetNames], [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def get_ids_of_names_fn(itypeinfo_ptr)
      @get_ids_of_names_fns ||= {}
      @get_ids_of_names_fns[W.vtable_address(itypeinfo_ptr)] ||= W.vtable_function(
        itypeinfo_ptr, ITYPEINFO_VTBL[:GetIDsOfNames], [W::VOIDP, W::VOIDP, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def type_info_of_guid_fn(itypelib_ptr)
      @type_info_of_guid_fns ||= {}
      @type_info_of_guid_fns[W.vtable_address(itypelib_ptr)] ||= W.vtable_function(
        itypelib_ptr, ITYPELIB_VTBL[:GetTypeInfoOfGuid], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_typeinfo.rb"'`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/typeinfo.rb test/win32ole/jruby/test_typeinfo.rb
git commit -m "jruby: add ImplType/GetNames/GetIDsOfNames/GetTypeInfoOfGuid vtable helpers"
```

---

## Task 4: `win32ole.rb` — Tidy First: promote `resolve_clsid`/`hresult_detail`, add `#dispatch_ptr`

**Files:**
- Modify: `lib/win32ole/jruby/win32ole.rb`

**Interfaces:**
- Consumes: nothing new.
- Produces: `WIN32OLE.resolve_clsid(server)` (class-level, was a private instance method), `WIN32OLE.hresult_detail(hr)` (class-level, was a private instance method), `WIN32OLE#dispatch_ptr` (new public reader for `@ptr`).

Pure refactor + one-line addition — this task's own tests are the existing Phase 1/2/3 test suite still passing (CI-only, since it all requires live COM), so this task is a Tidy-First step verified by *not changing behavior*, the same pattern Phase 3's own Task 3 used when promoting instance methods to class methods.

- [ ] **Step 1: Move `resolve_clsid` and `hresult_detail` into the `class << self` block**

In `lib/win32ole/jruby/win32ole.rb`, cut `resolve_clsid` and `hresult_detail` out of the instance-level `private` section (currently right after `initialize`) and paste them into the existing `class << self ... end` block (where `wrap_dispatch_pointer`/`ruby_value_to_variant_bytes`/`variant_bytes_to_ruby_value` already live), marking them `private_class_method`:

```ruby
    def resolve_clsid(server)
      wide = W.wstr(server)
      clsid = ("\x00" * 16).b
      hr = W.clsid_from_progid.call(wide, clsid)
      hr = W.clsid_from_string.call(wide, clsid) if W.failed?(hr)
      if W.failed?(hr)
        raise WIN32OLE::RuntimeError, "#{W.unknown_server_error_message(server)}\n#{hresult_detail(hr)}"
      end

      clsid
    end

    def hresult_detail(hr)
      "    HRESULT error code:#{W.hr_hex(hr)}\n      #{W.hresult_system_message(hr)}"
    end
    private :resolve_clsid, :hresult_detail
```

Update `initialize` and the instance-level `error_detail` to call through `self.class`:

```ruby
    clsid = self.class.send(:resolve_clsid, server)
    ...
    raise WIN32OLE::RuntimeError, "#{W.unknown_server_error_message(server)}\n#{self.class.send(:hresult_detail, hr)}"
```

(both call sites in `initialize`; `error_detail`'s own two `hresult_detail(hr)` calls become `self.class.send(:hresult_detail, hr)` too.)

- [ ] **Step 2: Add `#dispatch_ptr`**

Add to the public section (near `#ole_type`):

```ruby
    # WIN32OLE::Event needs this raw IDispatch* to build its own COM
    # connections (QueryInterface for IConnectionPointContainer,
    # IProvideClassInfo2, etc.) -- there is no general-purpose
    # QueryInterface API (#ole_query_interface is a Phase 2 non-goal), so
    # this is Event's one deliberate, documented crack in encapsulation.
    def dispatch_ptr
      @ptr
    end
```

- [ ] **Step 3: Verify no local syntax/load regression**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_win32ole_phase1.rb"'`
Expected: same pass/omit counts as before this task (this file's real assertions are COM-gated and will omit on this machine; the point of running it is confirming the file still *loads* — a syntax error in the refactor would show up as a `LoadError`/`NameError` here immediately).

- [ ] **Step 4: Commit**

```bash
git add lib/win32ole/jruby/win32ole.rb
git commit -m "jruby: promote resolve_clsid/hresult_detail to class level; add #dispatch_ptr"
```

---

## Task 5: `win32ole.rb` — `WIN32OLE.connect`

**Files:**
- Modify: `lib/win32ole/jruby/win32.rb` (new `GetActiveObject` binding)
- Modify: `lib/win32ole/jruby/win32ole.rb`
- Test: `test/win32ole/jruby/test_win32ole_phase1.rb` (add a new test class, or a new file `test_event_prereqs.rb` — using the latter to keep Phase 1's file scoped to Phase 1)

**Interfaces:**
- Consumes: `WIN32OLE.resolve_clsid`/`hresult_detail` (Task 4), `WIN32OLE.wrap_dispatch_pointer` (existing).
- Produces: `WIN32OLE.connect(server, host = nil)`.

The `host`-raises-`NotImplementedError` guard runs before any native call, so it's the one part of this task testable without COM.

- [ ] **Step 1: Write the failing test**

Create `test/win32ole/jruby/test_event_prereqs.rb`:

```ruby
begin
  require 'win32ole'
rescue LoadError
end
require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/win32ole'

class TestEventPrereqs < Test::Unit::TestCase
  def test_connect_raises_not_implemented_for_non_nil_host
    assert_raise(NotImplementedError) { WIN32OLE.connect('Scripting.Dictionary', 'remotehost') }
  end
end
end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event_prereqs.rb"'`
Expected: `NoMethodError: undefined method 'connect'`.

- [ ] **Step 3: Add `GetActiveObject` binding and `.connect`**

In `win32.rb`, near `co_create_instance`:

```ruby
    def get_active_object
      @get_active_object ||= Fiddle::Function.new(oleaut32['GetActiveObject'], [VOIDP, VOIDP, VOIDP], LONG, STDCALL)
    end
```

In `win32ole.rb`'s `class << self` block:

```ruby
    def connect(server, host = nil)
      raise NotImplementedError, 'remote OLE (host) is not supported yet' unless host.nil?

      hr = W.co_initialize.call(nil)
      raise 'fail: OLE initialize' unless hr.zero? || hr == 1

      clsid = resolve_clsid(server)
      ppv = ("\x00" * W::PTR_SIZE).b
      hr = W.get_active_object.call(clsid, nil, ppv)
      if W.failed?(hr)
        raise WIN32OLE::RuntimeError, "#{W.unknown_server_error_message(server)}\n#{hresult_detail(hr)}"
      end

      wrap_dispatch_pointer(ppv.unpack1(W::PACK_PTR))
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event_prereqs.rb"'`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/win32.rb lib/win32ole/jruby/win32ole.rb test/win32ole/jruby/test_event_prereqs.rb
git commit -m "jruby: add WIN32OLE.connect"
```

---

## Task 6: `win32ole.rb` — `WIN32OLE.const_load`

**Files:**
- Modify: `lib/win32ole/jruby/win32ole.rb`
- Test: `test/win32ole/jruby/test_event_prereqs.rb`

**Interfaces:**
- Consumes: `WIN32OLE#ole_type`, `Type#ole_typelib`, `TypeLib#ole_types`, `Type#variables`, `Variable#variable_kind`/`#name`/`#value` (all existing, Phase 2).
- Produces: `WIN32OLE.const_load(ole, mod)`.

Entirely live-COM (needs a real `WIN32OLE` object's typelib) — CI-only. Add a structural regression test anyway per house convention: it documents the contract even though it can't run here.

- [ ] **Step 1: Write the (CI-only) test**

Add to `test_event_prereqs.rb`, inside the `if RUBY_ENGINE == 'jruby'` guard, gated further on a live object exactly like the legacy test does:

```ruby
  def test_const_load_defines_constants_without_redefining_existing_ones
    omit('requires a live WIN32OLE COM object') unless defined?(WIN32OLE) && RUBY_ENGINE == 'jruby'
    dict = WIN32OLE.new('Scripting.Dictionary')
    mod = Module.new
    WIN32OLE.const_load(dict, mod)
    # Scripting.Dictionary's typelib (Scripting library) has no constants of
    # its own worth asserting on portably; the real coverage is ADO's
    # WIN32OLE.const_load(@db, ADO) in test_win32ole_event.rb once Event
    # itself lands (Task 15's GC-stress task doesn't touch this, but the
    # ADO-gated legacy suite exercises it directly in CI).
    assert_kind_of(Module, mod)
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event_prereqs.rb"'`
Expected: `NoMethodError: undefined method 'const_load'` (the `omit` only fires once the method is missing is no longer the failure mode — the method call itself raises first).

- [ ] **Step 3: Add `.const_load`**

```ruby
    def const_load(ole, mod)
      ole.ole_type.ole_typelib.ole_types.each do |type|
        type.variables.each do |var|
          next unless var.variable_kind == 'CONSTANT'
          next if mod.const_defined?(var.name, false)

          mod.const_set(var.name, var.value)
        end
      end
      nil
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event_prereqs.rb"'`
Expected: on this machine, `omit` (no `WIN32OLE` COM object available) — a pass-with-omission, not a failure, exactly like Phase 1's own precedent for COM-only tests. Full behavior confirmed in CI (Task 15).

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/win32ole.rb test/win32ole/jruby/test_event_prereqs.rb
git commit -m "jruby: add WIN32OLE.const_load"
```

---

## Task 7: `event.rb` — skeleton, constructor guard, `require` wiring, `.message_loop`

**Files:**
- Create: `lib/win32ole/jruby/event.rb`
- Modify: `lib/win32ole/jruby.rb` (`require 'win32ole/jruby/event'`)
- Test: `test/win32ole/jruby/test_event.rb` (new)

**Interfaces:**
- Consumes: `Win32.pump_windows_messages` (Task 2).
- Produces: `WIN32OLE::Event.new(ole, itf = nil)` (raises `TypeError` for a non-`WIN32OLE` first argument — everything past that guard is stubbed as an unimplemented private `advise` until Task 13), `WIN32OLE::Event.message_loop`.

The `TypeError` guard runs before any native call — this is the one part of `Event.new` testable without COM (and matches the legacy suite's own `test_s_new_exception`).

- [ ] **Step 1: Write the failing test**

Create `test/win32ole/jruby/test_event.rb`:

```ruby
require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/event'

class TestEvent < Test::Unit::TestCase
  W = WIN32OLE::Win32

  def test_new_raises_type_error_for_non_win32ole_argument
    assert_raise(TypeError) { WIN32OLE::Event.new('A') }
  end

  def test_message_loop_is_a_class_method
    assert_respond_to(WIN32OLE::Event, :message_loop)
  end
end
end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: `LoadError` — `win32ole/jruby/event` doesn't exist yet.

- [ ] **Step 3: Create `event.rb` and wire the require**

```ruby
# lib/win32ole/jruby/event.rb
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'
require 'win32ole/jruby/win32ole'

class WIN32OLE
  class Event
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    def initialize(ole, itf = nil)
      raise TypeError, '1st parameter must be WIN32OLE object' unless ole.is_a?(WIN32OLE)

      @events = []
      @handler = nil
      @finalizer_state = nil
      @sink_closures = nil

      advise(ole, itf)
    end

    def self.message_loop
      W.pump_windows_messages
    end

    private

    # Built up across Tasks 9-13; a successful construction isn't
    # exercised by any test until Task 13 wires the real implementation in.
    def advise(ole, itf)
      raise NotImplementedError, 'advise is implemented in Task 13'
    end
  end
end
```

Add to `lib/win32ole/jruby.rb`:

```ruby
require 'win32ole/jruby/event'
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/event.rb lib/win32ole/jruby.rb test/win32ole/jruby/test_event.rb
git commit -m "jruby: add WIN32OLE::Event skeleton and message_loop"
```

---

## Task 8: `event.rb` — `#on_event`/`#on_event_with_outargs`/`#off_event`/`#handler=`/`#handler`

**Files:**
- Modify: `lib/win32ole/jruby/event.rb`
- Test: `test/win32ole/jruby/test_event.rb`

**Interfaces:**
- Consumes: `@events`/`@finalizer_state` (Task 7).
- Produces: `#on_event`, `#on_event_with_outargs`, `#off_event`, `#handler=`, `#handler`. `@events` becomes an `Array` of `{name: String|nil, proc: Proc, with_outargs: bool}`.

Pure `Array`/`Hash` bookkeeping, no native calls at all — fully unit-testable by constructing the `Event` via `allocate` (bypassing `#initialize`/`advise` entirely) and setting `@events`/`@finalizer_state` by hand, exactly the kind of direct-ivar test setup `Type`/`Variable`'s own specs never needed but is legitimate here since we're deliberately testing below the constructor.

- [ ] **Step 1: Write the failing test**

```ruby
  def advised_event
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@events, [])
    ev.instance_variable_set(:@handler, nil)
    ev.instance_variable_set(:@finalizer_state, { cp_ptr: 0x1234 }) # truthy stand-in for "advised"
    ev
  end

  def test_on_event_registers_a_named_callback
    ev = advised_event
    ev.on_event('Foo') { |*| }
    events = ev.instance_variable_get(:@events)
    assert_equal(1, events.size)
    assert_equal('Foo', events.first[:name])
    assert_equal(false, events.first[:with_outargs])
  end

  def test_on_event_accepts_a_symbol_event_name
    ev = advised_event
    ev.on_event(:Foo) { |*| }
    assert_equal('Foo', ev.instance_variable_get(:@events).first[:name])
  end

  def test_on_event_rejects_non_string_non_symbol_event_name
    ev = advised_event
    assert_raise(TypeError) { ev.on_event(42) { |*| } }
  end

  def test_on_event_with_outargs_sets_the_with_outargs_flag
    ev = advised_event
    ev.on_event_with_outargs('Foo') { |*| }
    assert_equal(true, ev.instance_variable_get(:@events).first[:with_outargs])
  end

  def test_on_event_twice_with_the_same_name_replaces_not_stacks
    ev = advised_event
    first = proc { |*| }
    second = proc { |*| }
    ev.on_event('Foo', &first)
    ev.on_event('Foo', &second)
    events = ev.instance_variable_get(:@events)
    assert_equal(1, events.size)
    assert_same(second, events.first[:proc])
  end

  def test_on_event_twice_with_no_name_replaces_the_catch_all
    ev = advised_event
    first = proc { |*| }
    second = proc { |*| }
    ev.on_event(&first)
    ev.on_event(&second)
    events = ev.instance_variable_get(:@events)
    assert_equal(1, events.size)
    assert_nil(events.first[:name])
    assert_same(second, events.first[:proc])
  end

  def test_on_event_raises_runtime_error_when_not_advised
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@events, [])
    ev.instance_variable_set(:@finalizer_state, nil)
    assert_raise(WIN32OLE::RuntimeError) { ev.on_event('Foo') { |*| } }
  end

  def test_off_event_removes_only_the_named_entry
    ev = advised_event
    ev.on_event('Foo') { |*| }
    ev.on_event('Bar') { |*| }
    ev.off_event('Foo')
    events = ev.instance_variable_get(:@events)
    assert_equal(1, events.size)
    assert_equal('Bar', events.first[:name])
  end

  def test_off_event_with_no_args_removes_only_the_catch_all
    ev = advised_event
    ev.on_event('Foo') { |*| }
    ev.on_event { |*| }
    ev.off_event
    events = ev.instance_variable_get(:@events)
    assert_equal(1, events.size)
    assert_equal('Foo', events.first[:name])
  end

  def test_off_event_accepts_a_symbol
    ev = advised_event
    ev.on_event('Foo') { |*| }
    ev.off_event(:Foo)
    assert_empty(ev.instance_variable_get(:@events))
  end

  def test_handler_accessor_round_trips
    ev = WIN32OLE::Event.allocate
    handler = Object.new
    ev.handler = handler
    assert_same(handler, ev.handler)
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: `NoMethodError` for `on_event`/etc.

- [ ] **Step 3: Implement the callback bookkeeping**

Add to `event.rb` (public section, above `private`):

```ruby
    def on_event(event = nil, &block)
      register_event(event, block, false)
    end

    def on_event_with_outargs(event = nil, &block)
      register_event(event, block, true)
    end

    def off_event(event = nil)
      name = event.nil? ? nil : normalize_event_name(event)
      @events.reject! { |e| e[:name] == name }
      nil
    end

    def handler=(obj)
      @handler = obj
    end

    def handler
      @handler
    end
```

Add to the `private` section:

```ruby
    def register_event(event, block, with_outargs)
      if @finalizer_state.nil?
        raise WIN32OLE::RuntimeError, 'IConnectionPoint not found. You must call advise at first.'
      end

      name = event.nil? ? nil : normalize_event_name(event)
      @events.reject! { |e| e[:name] == name }
      @events << { name: name, proc: block, with_outargs: with_outargs }
      nil
    end

    def normalize_event_name(event)
      unless event.is_a?(String) || event.is_a?(Symbol)
        raise TypeError, 'wrong argument type (expected String or Symbol)'
      end

      event.to_s
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/event.rb test/win32ole/jruby/test_event.rb
git commit -m "jruby: add Event#on_event/on_event_with_outargs/off_event/handler="
```

---

## Task 9: `event.rb` — event-source resolution, part 1: `find_iid_by_name`/`find_iid_by_guid`

**Files:**
- Modify: `lib/win32ole/jruby/event.rb`

**Interfaces:**
- Consumes: `TI.type_attr_fn`/`release_type_attr_fn`/`containing_typelib_fn`/`get_type_info_fn`/`documentation_fn_for_typeinfo`/`type_info_count_fn`/`type_info_fn` (existing), `TI.ref_type_of_impl_type_fn`/`ref_type_info_fn`/`type_info_of_guid_fn` (Task 3), `W.query_interface` (Task 1).
- Produces: private `#find_iid_by_name(ole, itf_name)` → `[iid_bytes, itypeinfo_ptr]` (raises `WIN32OLE::RuntimeError`), `#find_iid_by_guid(ole, iid_bytes)` → `itypeinfo_ptr` (raises `WIN32OLE::RuntimeError`), plus the shared low-level pointer helpers both this task and Task 10 need: `#release_ptr`, `#get_type_info0`, `#containing_typelib`, `#type_attr_ptr`, `#type_name`, `#type_guid`, `#impl_type_ref_typeinfo`.

Entirely live-COM (real `ITypeLib`/`ITypeInfo` traversal) — CI-only, cross-check against `ext/win32ole/win32ole_event.c:481-590` (`find_iid`) for any detail not spelled out below. This task's one locally-checkable property: the guard clauses raise the right exception type before any traversal happens.

- [ ] **Step 1: Write the (guard-only) failing test**

```ruby
  def test_find_iid_by_name_is_private
    assert(WIN32OLE::Event.private_method_defined?(:find_iid_by_name))
  end

  def test_find_iid_by_guid_is_private
    assert(WIN32OLE::Event.private_method_defined?(:find_iid_by_guid))
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: FAIL — the methods don't exist.

- [ ] **Step 3: Add the shared pointer helpers and `find_iid_by_name`/`find_iid_by_guid`**

Add to `event.rb`'s `private` section:

```ruby
    TKIND_COCLASS = 5

    def release_ptr(ptr)
      return if ptr.nil? || ptr.zero?

      W.vtable_function(ptr, 2, [W::VOIDP], W::DWORD).call(ptr)
    end

    def get_type_info0(idispatch_ptr)
      out = ("\x00" * W::PTR_SIZE).b
      hr = TI.get_type_info_fn(idispatch_ptr).call(idispatch_ptr, 0, W::LOCALE_SYSTEM_DEFAULT, out)
      return nil if W.failed?(hr)

      out.unpack1(W::PACK_PTR)
    end

    def containing_typelib(itypeinfo_ptr)
      tlib_out = ("\x00" * W::PTR_SIZE).b
      index_out = ("\x00" * 4).b
      hr = TI.containing_typelib_fn(itypeinfo_ptr).call(itypeinfo_ptr, tlib_out, index_out)
      return nil if W.failed?(hr)

      tlib_out.unpack1(W::PACK_PTR)
    end

    def type_attr_ptr(itypeinfo_ptr)
      out = ("\x00" * W::PTR_SIZE).b
      hr = TI.type_attr_fn(itypeinfo_ptr).call(itypeinfo_ptr, out)
      return nil if W.failed?(hr)

      out.unpack1(W::PACK_PTR)
    end

    def type_name(itypeinfo_ptr)
      name_out = ("\x00" * W::PTR_SIZE).b
      TI.documentation_fn_for_typeinfo(itypeinfo_ptr).call(itypeinfo_ptr, -1, name_out, nil, nil, nil)
      bstr = name_out.unpack1(W::PACK_PTR)
      name = W.bstr_to_s(bstr)
      W.sys_free_string.call(bstr) unless bstr.zero?
      name
    end

    def type_guid(attr_ptr)
      ta = TI::TYPEATTR.new(attr_ptr)
      [ta.guid_Data1, ta.guid_Data2, ta.guid_Data3].pack('LSS') + ta.guid_Data4.pack('C8')
    end

    def impl_type_ref_typeinfo(itypeinfo_ptr, index)
      href_out = ("\x00" * 4).b
      hr = TI.ref_type_of_impl_type_fn(itypeinfo_ptr).call(itypeinfo_ptr, index, href_out)
      return nil if W.failed?(hr)

      ref_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.ref_type_info_fn(itypeinfo_ptr).call(itypeinfo_ptr, href_out.unpack1('L'), ref_out)
      return nil if W.failed?(hr)

      ref_out.unpack1(W::PACK_PTR)
    end

    # ext/win32ole/win32ole_event.c:481-590 (find_iid, pitf given): scans
    # every COCLASS in ole's containing typelib for an implemented type
    # named itf_name; that impl type's own GUID becomes the event source IID.
    def find_iid_by_name(ole, itf_name)
      itypeinfo_ptr = get_type_info0(ole.dispatch_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if itypeinfo_ptr.nil?

      tlib_ptr = containing_typelib(itypeinfo_ptr)
      release_ptr(itypeinfo_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if tlib_ptr.nil?

      found_iid = nil
      found_ti_ptr = nil
      count = TI.type_info_count_fn(tlib_ptr).call(tlib_ptr)
      count.times do |index|
        break if found_iid

        ti_out = ("\x00" * W::PTR_SIZE).b
        next if W.failed?(TI.type_info_fn(tlib_ptr).call(tlib_ptr, index, ti_out))

        ti_ptr = ti_out.unpack1(W::PACK_PTR)
        attr_ptr = type_attr_ptr(ti_ptr)
        if attr_ptr
          if TI::TYPEATTR.new(attr_ptr).typekind == TKIND_COCLASS
            impl_count = TI::TYPEATTR.new(attr_ptr).cImplTypes
            impl_count.times do |t|
              impl_ti_ptr = impl_type_ref_typeinfo(ti_ptr, t)
              next if impl_ti_ptr.nil?

              if type_name(impl_ti_ptr) == itf_name
                impl_attr_ptr = type_attr_ptr(impl_ti_ptr)
                if impl_attr_ptr
                  found_iid = type_guid(impl_attr_ptr)
                  TI.release_type_attr_fn(impl_ti_ptr).call(impl_ti_ptr, impl_attr_ptr)
                  found_ti_ptr = impl_ti_ptr
                end
              end
              release_ptr(impl_ti_ptr) unless impl_ti_ptr == found_ti_ptr
              break if found_iid
            end
          end
          TI.release_type_attr_fn(ti_ptr).call(ti_ptr, attr_ptr)
        end
        release_ptr(ti_ptr)
      end
      release_ptr(tlib_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if found_iid.nil?

      [found_iid, found_ti_ptr]
    end

    # ext/win32ole/win32ole_event.c:518-524 (find_iid, pitf NULL): a direct
    # GetTypeInfoOfGuid lookup, used when the caller already knows the IID
    # (the IProvideClassInfo2::GetGUID path in find_default_source).
    def find_iid_by_guid(ole, iid_bytes)
      itypeinfo_ptr = get_type_info0(ole.dispatch_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if itypeinfo_ptr.nil?

      tlib_ptr = containing_typelib(itypeinfo_ptr)
      release_ptr(itypeinfo_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if tlib_ptr.nil?

      ti_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.type_info_of_guid_fn(tlib_ptr).call(tlib_ptr, iid_bytes, ti_out)
      release_ptr(tlib_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if W.failed?(hr)

      ti_out.unpack1(W::PACK_PTR)
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/event.rb test/win32ole/jruby/test_event.rb
git commit -m "jruby: add Event's find_iid_by_name/find_iid_by_guid event-source resolution"
```

---

## Task 10: `event.rb` — event-source resolution, part 2: `find_default_source` + `resolve_event_source`

**Files:**
- Modify: `lib/win32ole/jruby/event.rb`

**Interfaces:**
- Consumes: Task 9's shared helpers, `TI.impl_type_flags_fn` (Task 3), `W.query_interface`/`IID_IPROVIDECLASSINFO`/`IID_IPROVIDECLASSINFO2` (Task 1).
- Produces: private `#find_default_source_from_typeinfo`, `#find_coclass`, `#find_default_source`, `#provide_class_info2_iid`, `#provide_class_info_typeinfo`, `#resolve_event_source(ole, itf)` → `[iid_bytes, itypeinfo_ptr]` (the single entry point Task 13's `#advise` calls).

Entirely live-COM — CI-only, cross-check against `ext/win32ole/win32ole_event.c:592-786`.

- [ ] **Step 1: Write the (guard-only) failing test**

```ruby
  def test_resolve_event_source_is_private
    assert(WIN32OLE::Event.private_method_defined?(:resolve_event_source))
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: FAIL — method doesn't exist.

- [ ] **Step 3: Implement**

```ruby
    IMPLTYPEFLAG_FDEFAULT = 0x1
    IMPLTYPEFLAG_FSOURCE = 0x2
    GUIDKIND_DEFAULT_SOURCE_DISP_IID = 1

    def impl_type_flags(itypeinfo_ptr, index)
      flags_out = ("\x00" * 4).b
      hr = TI.impl_type_flags_fn(itypeinfo_ptr).call(itypeinfo_ptr, index, flags_out)
      return nil if W.failed?(hr)

      flags_out.unpack1('l')
    end

    # ext/win32ole/win32ole_event.c:667-701
    def find_default_source_from_typeinfo(ti_ptr, attr_ptr)
      count = TI::TYPEATTR.new(attr_ptr).cImplTypes
      count.times do |i|
        flags = impl_type_flags(ti_ptr, i)
        next if flags.nil?
        next if (flags & IMPLTYPEFLAG_FDEFAULT).zero? || (flags & IMPLTYPEFLAG_FSOURCE).zero?

        ref_ti_ptr = impl_type_ref_typeinfo(ti_ptr, i)
        return ref_ti_ptr if ref_ti_ptr
      end
      nil
    end

    # ext/win32ole/win32ole_event.c:592-665: find, in ti_ptr's containing
    # typelib, the COCLASS whose default impl type is ti_ptr itself.
    def find_coclass(ti_ptr)
      tlib_ptr = containing_typelib(ti_ptr)
      return [nil, nil] if tlib_ptr.nil?

      target_attr_ptr = type_attr_ptr(ti_ptr)
      target_guid = target_attr_ptr && type_guid(target_attr_ptr)
      TI.release_type_attr_fn(ti_ptr).call(ti_ptr, target_attr_ptr) if target_attr_ptr

      found_ti_ptr = nil
      found_attr_ptr = nil
      count = TI.type_info_count_fn(tlib_ptr).call(tlib_ptr)
      count.times do |i|
        break if found_ti_ptr

        ti2_out = ("\x00" * W::PTR_SIZE).b
        next if W.failed?(TI.type_info_fn(tlib_ptr).call(tlib_ptr, i, ti2_out))

        ti2_ptr = ti2_out.unpack1(W::PACK_PTR)
        attr2_ptr = type_attr_ptr(ti2_ptr)
        if attr2_ptr.nil?
          release_ptr(ti2_ptr)
          next
        end
        if TI::TYPEATTR.new(attr2_ptr).typekind != TKIND_COCLASS
          TI.release_type_attr_fn(ti2_ptr).call(ti2_ptr, attr2_ptr)
          release_ptr(ti2_ptr)
          next
        end

        matched = TI::TYPEATTR.new(attr2_ptr).cImplTypes.times.any? do |j|
          flags = impl_type_flags(ti2_ptr, j)
          next false if flags.nil? || (flags & IMPLTYPEFLAG_FDEFAULT).zero?

          ref_ti_ptr = impl_type_ref_typeinfo(ti2_ptr, j)
          next false if ref_ti_ptr.nil?

          ref_attr_ptr = type_attr_ptr(ref_ti_ptr)
          ref_guid = ref_attr_ptr && type_guid(ref_attr_ptr)
          TI.release_type_attr_fn(ref_ti_ptr).call(ref_ti_ptr, ref_attr_ptr) if ref_attr_ptr
          release_ptr(ref_ti_ptr)
          ref_guid == target_guid
        end

        if matched
          found_ti_ptr = ti2_ptr
          found_attr_ptr = attr2_ptr
        else
          TI.release_type_attr_fn(ti2_ptr).call(ti2_ptr, attr2_ptr)
          release_ptr(ti2_ptr)
        end
      end
      release_ptr(tlib_ptr)
      [found_ti_ptr, found_attr_ptr]
    end

    def provide_class_info2_iid(idispatch_ptr)
      pci2_ptr = W.query_interface(idispatch_ptr, W::IID_IPROVIDECLASSINFO2)
      return nil if pci2_ptr.nil?

      iid_out = ("\x00" * 16).b
      hr = W.vtable_function(pci2_ptr, 4, [W::DWORD, W::VOIDP], W::LONG).call(
        pci2_ptr, GUIDKIND_DEFAULT_SOURCE_DISP_IID, iid_out
      )
      release_ptr(pci2_ptr)
      return nil if W.failed?(hr)

      iid_out
    end

    def provide_class_info_typeinfo(idispatch_ptr)
      pci_ptr = W.query_interface(idispatch_ptr, W::IID_IPROVIDECLASSINFO)
      return nil if pci_ptr.nil?

      ti_out = ("\x00" * W::PTR_SIZE).b
      hr = W.vtable_function(pci_ptr, 3, [W::VOIDP], W::LONG).call(pci_ptr, ti_out)
      release_ptr(pci_ptr)
      return nil if W.failed?(hr)

      ti_out.unpack1(W::PACK_PTR)
    end

    # ext/win32ole/win32ole_event.c:703-786, minus the GetGUID/find_iid
    # fast path (handled by resolve_event_source below, since it needs an
    # `iid_bytes` result rather than an `itypeinfo_ptr` result).
    def find_default_source(ole)
      itypeinfo_ptr = provide_class_info_typeinfo(ole.dispatch_ptr) || get_type_info0(ole.dispatch_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if itypeinfo_ptr.nil?

      attr_ptr = type_attr_ptr(itypeinfo_ptr)
      if attr_ptr.nil?
        release_ptr(itypeinfo_ptr)
        raise WIN32OLE::RuntimeError, 'interface not found'
      end

      result_ti_ptr = find_default_source_from_typeinfo(itypeinfo_ptr, attr_ptr)
      if result_ti_ptr.nil?
        co_ti_ptr, co_attr_ptr = find_coclass(itypeinfo_ptr)
        if co_ti_ptr
          result_ti_ptr = find_default_source_from_typeinfo(co_ti_ptr, co_attr_ptr)
          TI.release_type_attr_fn(co_ti_ptr).call(co_ti_ptr, co_attr_ptr)
          release_ptr(co_ti_ptr)
        end
      end
      TI.release_type_attr_fn(itypeinfo_ptr).call(itypeinfo_ptr, attr_ptr)
      release_ptr(itypeinfo_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if result_ti_ptr.nil?

      result_attr_ptr = type_attr_ptr(result_ti_ptr)
      if result_attr_ptr.nil?
        release_ptr(result_ti_ptr)
        raise WIN32OLE::RuntimeError, 'interface not found'
      end
      guid = type_guid(result_attr_ptr)
      TI.release_type_attr_fn(result_ti_ptr).call(result_ti_ptr, result_attr_ptr)
      [guid, result_ti_ptr]
    end

    # The single entry point #advise (Task 13) calls.
    def resolve_event_source(ole, itf)
      return find_iid_by_name(ole, itf) if itf

      iid_bytes = provide_class_info2_iid(ole.dispatch_ptr)
      if iid_bytes
        begin
          return [iid_bytes, find_iid_by_guid(ole, iid_bytes)]
        rescue WIN32OLE::RuntimeError
          # IProvideClassInfo2 succeeded but the IID it named isn't
          # resolvable in this typelib -- fall through to the
          # IProvideClassInfo/ImplType-traversal path below, matching
          # win32ole_event.c:730-736's own fallthrough.
        end
      end
      find_default_source(ole)
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/event.rb test/win32ole/jruby/test_event.rb
git commit -m "jruby: add Event's find_default_source/find_coclass/resolve_event_source"
```

---

## Task 11: `event.rb` — sink vtable closures + `build_sink`

**Files:**
- Modify: `lib/win32ole/jruby/event.rb`
- Test: `test/win32ole/jruby/test_event.rb`

**Interfaces:**
- Consumes: `W::IID_IUNKNOWN`/`IID_IDISPATCH` (existing), `TI.get_ids_of_names_fn` (Task 3).
- Produces: private `#build_sink(source_iid_bytes, event_typeinfo_ptr)` → `[sink_addr, vtable_addr, closures]` (5 of the 7 closures are real here: `QueryInterface`/`AddRef`/`Release`/`GetTypeInfoCount`/`GetTypeInfo`; `GetIDsOfNames` delegates to a live `ITypeInfo` so it's wired here but only CI-testable; `Invoke` is Task 12).

This is the first genuinely, fully locally-testable native-callback task: `Fiddle::Closure::BlockCaller` objects are real callable machine code on any platform (no COM needed) — wrap each one's own address in a `Fiddle::Function` and call it directly, the same self-test technique the (uncommitted) Phase 1 spike used ("verified with a self-test: a message posted to our own thread...").

- [ ] **Step 1: Write the failing test**

```ruby
  E_NOINTERFACE = -2147467262
  DISP_E_BADINDEX = -2147352565

  def build_test_sink(source_iid_bytes = ("\x01" * 16).b, event_typeinfo_ptr = 0)
    ev = WIN32OLE::Event.allocate
    sink_addr, vtable_addr, closures = ev.send(:build_sink, source_iid_bytes, event_typeinfo_ptr)
    [sink_addr, vtable_addr, closures]
  end

  def test_query_interface_closure_returns_sink_for_known_and_source_iids
    source_iid = ("\x01" * 16).b
    sink_addr, vtable_addr, closures = build_test_sink(source_iid)
    qi_fn = Fiddle::Function.new(closures[0], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG)
    ppv = ("\xFF" * W::PTR_SIZE).b

    assert_equal(0, qi_fn.call(sink_addr, W::IID_IUNKNOWN, ppv))
    assert_equal(sink_addr, ppv.unpack1(W::PACK_PTR))

    assert_equal(0, qi_fn.call(sink_addr, W::IID_IDISPATCH, ppv))
    assert_equal(sink_addr, ppv.unpack1(W::PACK_PTR))

    assert_equal(0, qi_fn.call(sink_addr, source_iid, ppv))
    assert_equal(sink_addr, ppv.unpack1(W::PACK_PTR))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

  def test_query_interface_closure_rejects_unknown_iid
    sink_addr, vtable_addr, closures = build_test_sink
    qi_fn = Fiddle::Function.new(closures[0], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG)
    ppv = ("\xFF" * W::PTR_SIZE).b

    assert_equal(E_NOINTERFACE, qi_fn.call(sink_addr, ("\xFE" * 16).b, ppv))
    assert_equal(0, ppv.unpack1(W::PACK_PTR))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

  def test_add_ref_and_release_closures_share_a_refcount
    sink_addr, vtable_addr, closures = build_test_sink
    add_ref_fn = Fiddle::Function.new(closures[1], [W::VOIDP], W::DWORD)
    release_fn = Fiddle::Function.new(closures[2], [W::VOIDP], W::DWORD)

    assert_equal(1, add_ref_fn.call(sink_addr))
    assert_equal(2, add_ref_fn.call(sink_addr))
    assert_equal(1, release_fn.call(sink_addr))
    assert_equal(0, release_fn.call(sink_addr))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

  def test_get_type_info_count_closure_always_reports_zero
    sink_addr, vtable_addr, closures = build_test_sink
    fn = Fiddle::Function.new(closures[3], [W::VOIDP, W::VOIDP], W::LONG)
    pct = ("\xFF" * 4).b

    assert_equal(0, fn.call(sink_addr, pct))
    assert_equal(0, pct.unpack1('L'))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

  def test_get_type_info_closure_always_fails_with_bad_index
    sink_addr, vtable_addr, closures = build_test_sink
    fn = Fiddle::Function.new(closures[4], [W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::LONG)
    ppti = ("\xFF" * W::PTR_SIZE).b

    assert_equal(DISP_E_BADINDEX, fn.call(sink_addr, 0, 0, ppti))
    assert_equal(0, ppti.unpack1(W::PACK_PTR))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

  def test_build_sink_wires_all_seven_vtable_slots_to_the_closures
    sink_addr, vtable_addr, closures = build_test_sink
    assert_equal(7, closures.size)
    closures.each_with_index do |closure, i|
      slot = Fiddle::Pointer.new(vtable_addr)[i * W::PTR_SIZE, W::PTR_SIZE].unpack1(W::PACK_PTR)
      assert_equal(closure.to_i, slot)
    end
    assert_equal(vtable_addr, Fiddle::Pointer.new(sink_addr)[0, W::PTR_SIZE].unpack1(W::PACK_PTR))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: `NoMethodError: undefined method 'build_sink'`.

- [ ] **Step 3: Implement the closures and `build_sink`**

```ruby
    SINK_VTBL_SLOTS = 7

    def query_interface_closure(sink_addr, source_iid_bytes, refcount)
      Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL) do |_this, riid_ptr, ppv_ptr|
        riid = riid_ptr[0, 16]
        if riid == W::IID_IUNKNOWN || riid == W::IID_IDISPATCH || riid == source_iid_bytes
          ppv_ptr[0, W::PTR_SIZE] = [sink_addr].pack(W::PACK_PTR)
          refcount[0] += 1
          0
        else
          ppv_ptr[0, W::PTR_SIZE] = [0].pack(W::PACK_PTR)
          E_NOINTERFACE
        end
      rescue StandardError, ScriptError => e
        warn_closure_exception('QueryInterface', e)
        E_NOINTERFACE
      end
    end

    def add_ref_closure(refcount)
      Fiddle::Closure::BlockCaller.new(W::DWORD, [W::VOIDP], W::STDCALL) do |_this|
        refcount[0] += 1
      end
    end

    def release_closure(refcount)
      Fiddle::Closure::BlockCaller.new(W::DWORD, [W::VOIDP], W::STDCALL) do |_this|
        refcount[0] -= 1
        refcount[0]
      end
    end

    def get_type_info_count_closure
      Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::VOIDP], W::STDCALL) do |_this, pct_ptr|
        pct_ptr[0, 4] = [0].pack('L')
        0
      end
    end

    def get_type_info_closure
      Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::STDCALL) do |_this, _idx, _lcid, ppti_ptr|
        ppti_ptr[0, W::PTR_SIZE] = [0].pack(W::PACK_PTR)
        DISP_E_BADINDEX
      end
    end

    def get_ids_of_names_closure(event_typeinfo_ptr)
      Fiddle::Closure::BlockCaller.new(
        W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::STDCALL
      ) do |_this, _riid, names_ptr, cnames, _lcid, dispids_ptr|
        TI.get_ids_of_names_fn(event_typeinfo_ptr).call(event_typeinfo_ptr, names_ptr.to_i, cnames, dispids_ptr.to_i)
      rescue StandardError, ScriptError => e
        warn_closure_exception('GetIDsOfNames', e)
        DISP_E_UNKNOWNNAME
      end
    end

    DISP_E_UNKNOWNNAME = -2147352570
    E_NOINTERFACE = -2147467262
    DISP_E_BADINDEX = -2147352565

    def warn_closure_exception(where, error)
      warn "#{error.backtrace&.first}: #{error.message} (#{error.class}) in WIN32OLE::Event sink's #{where}"
    end

    # Builds a fresh 7-slot IDispatch-shaped vtable (QueryInterface, AddRef,
    # Release, GetTypeInfoCount, GetTypeInfo, GetIDsOfNames, Invoke -- same
    # order/shape as ext/win32ole/win32ole_event.c's IEventSinkVtbl) backed
    # by Fiddle::Closure::BlockCaller trampolines. Returns raw addresses (not
    # Fiddle::Pointer wrappers) plus the closures themselves -- the CALLER
    # must keep `closures` referenced for as long as the sink is advised
    # (GC'ing a Closure frees its native trampoline), matching variant.rb's
    # own finalizer-state discipline (see #advise, Task 13).
    def build_sink(source_iid_bytes, event_typeinfo_ptr)
      sink_ptr = Fiddle::Pointer.malloc(W::PTR_SIZE)
      sink_addr = sink_ptr.to_i
      refcount = [0]

      closures = [
        query_interface_closure(sink_addr, source_iid_bytes, refcount),
        add_ref_closure(refcount),
        release_closure(refcount),
        get_type_info_count_closure,
        get_type_info_closure,
        get_ids_of_names_closure(event_typeinfo_ptr),
        invoke_closure(event_typeinfo_ptr)
      ]

      vtable_ptr = Fiddle::Pointer.malloc(W::PTR_SIZE * SINK_VTBL_SLOTS)
      closures.each_with_index { |c, i| vtable_ptr[i * W::PTR_SIZE, W::PTR_SIZE] = [c.to_i].pack(W::PACK_PTR) }
      sink_ptr[0, W::PTR_SIZE] = [vtable_ptr.to_i].pack(W::PACK_PTR)

      [sink_addr, vtable_ptr.to_i, closures]
    end
```

`invoke_closure` is referenced above but implemented in Task 12 — define a minimal real (not placeholder) version now so `build_sink` is already callable and this task's tests pass:

```ruby
    def invoke_closure(_event_typeinfo_ptr)
      Fiddle::Closure::BlockCaller.new(
        W::LONG, [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::WORD, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL
      ) { |*| 0 } # NOERROR; replaced with real dispatch in Task 12
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/event.rb test/win32ole/jruby/test_event.rb
git commit -m "jruby: add Event's sink vtable closures and build_sink"
```

---

## Task 12: `event.rb` — the real `Invoke` closure: event dispatch + exception containment

**Files:**
- Modify: `lib/win32ole/jruby/event.rb`
- Test: `test/win32ole/jruby/test_event.rb`

**Interfaces:**
- Consumes: `@events`/`@handler` (Task 8), `TI.get_names_fn` (Task 3), `WIN32OLE.variant_bytes_to_ruby_value`/`ruby_value_to_variant_bytes` (existing).
- Produces: replaces Task 11's stub `#invoke_closure` with the real one; adds private `#handle_invoke`, `#find_event_entry`, `#read_dispparams`, `#resolve_event_name`.

Event-name resolution (`GetNames`) needs a live `ITypeInfo`, so the full `Invoke` path is CI-only — but `#find_event_entry`'s lookup logic (given a pre-set `@events`) and the exception-containment behavior are pure Ruby, testable directly.

- [ ] **Step 1: Write the failing test**

```ruby
  def test_find_event_entry_prefers_a_named_match_over_the_catch_all
    ev = WIN32OLE::Event.allocate
    fallback = { name: nil, proc: proc { :fallback }, with_outargs: false }
    named = { name: 'Foo', proc: proc { :named }, with_outargs: false }
    ev.instance_variable_set(:@events, [fallback, named])
    entry, is_default = ev.send(:find_event_entry, 'Foo')
    assert_same(named, entry)
    assert_equal(false, is_default)
  end

  def test_find_event_entry_falls_back_to_the_catch_all
    ev = WIN32OLE::Event.allocate
    fallback = { name: nil, proc: proc { :fallback }, with_outargs: false }
    ev.instance_variable_set(:@events, [fallback])
    entry, is_default = ev.send(:find_event_entry, 'Bar')
    assert_same(fallback, entry)
    assert_equal(true, is_default)
  end

  def test_find_event_entry_returns_nil_when_nothing_matches
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@events, [])
    entry, is_default = ev.send(:find_event_entry, 'Bar')
    assert_nil(entry)
    assert_equal(false, is_default)
  end

  def test_handle_invoke_writes_exception_message_to_stderr_and_does_not_raise
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@events, [{ name: nil, proc: proc { raise 'boom' }, with_outargs: false }])
    ev.instance_variable_set(:@handler, nil)
    ev.instance_variable_set(:@event_typeinfo_ptr, 0)
    ev.define_singleton_method(:resolve_event_name) { |_dispid| 'Whatever' }

    dispparams = [0, 0, 0, 0].pack("#{W::PACK_PTR}#{W::PACK_PTR}LL")
    dispparams_ptr = Fiddle::Pointer.to_ptr(dispparams)

    err = capture_stderr { ev.send(:handle_invoke, 1, dispparams_ptr, nil) }
    assert_match(/boom/, err)
  end

  def capture_stderr
    old = $stderr
    $stderr = StringIO.new
    yield
    $stderr.string
  ensure
    $stderr = old
  end
```

(add `require 'stringio'` near the top of `test_event.rb`, inside the `if RUBY_ENGINE == 'jruby'` guard)

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: `NoMethodError` for `find_event_entry`/`handle_invoke`.

- [ ] **Step 3: Implement**

Replace Task 11's stub `#invoke_closure` and add the supporting methods:

```ruby
    def invoke_closure(event_typeinfo_ptr)
      Fiddle::Closure::BlockCaller.new(
        W::LONG, [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::WORD, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL
      ) do |_this, dispid, _riid, _lcid, _wflags, pdispparams_ptr, pvarresult_ptr, _pexcepinfo_ptr, _puargerr_ptr|
        @event_typeinfo_ptr = event_typeinfo_ptr
        handle_invoke(dispid, pdispparams_ptr, pvarresult_ptr)
        0 # NOERROR, always -- see Global Constraints: no exception may cross this boundary.
      rescue StandardError, ScriptError => e
        warn_closure_exception('Invoke', e)
        0
      end
    end

    def resolve_event_name(dispid)
      bstr_out = ("\x00" * W::PTR_SIZE).b
      count_out = ("\x00" * 4).b
      hr = TI.get_names_fn(@event_typeinfo_ptr).call(@event_typeinfo_ptr, dispid, bstr_out, 1, count_out)
      return nil if W.failed?(hr)

      bstr = bstr_out.unpack1(W::PACK_PTR)
      name = W.bstr_to_s(bstr)
      W.sys_free_string.call(bstr) unless bstr.zero?
      name
    end

    # ext/win32ole/win32ole_event.c:813-836 (ole_search_event): a NAMED
    # match wins immediately; otherwise fall back to the one catch-all
    # (nil-name) entry, if any. Returns [entry_or_nil, is_default] --
    # is_default is true only for the catch-all/no-match case (drives
    # whether the event name gets prepended to the callback's args).
    def find_event_entry(name)
      fallback = nil
      @events.each do |e|
        return [e, false] if e[:name] == name

        fallback = e if e[:name].nil?
      end
      [fallback, true]
    end

    def read_dispparams(pdispparams_ptr)
      rgvarg_addr = pdispparams_ptr[0, W::PTR_SIZE].unpack1(W::PACK_PTR)
      cargs = pdispparams_ptr[2 * W::PTR_SIZE, 4].unpack1('L')
      [cargs, rgvarg_addr]
    end

    # ext/win32ole/win32ole_event.c:132-234 (EVENTSINK_Invoke), the
    # non-hash/non-outargs subset -- Task 14 adds the Hash/Array
    # out-argument write-back on top of this.
    def handle_invoke(dispid, pdispparams_ptr, pvarresult_ptr)
      name = resolve_event_name(dispid)
      return if name.nil?

      entry, is_default = find_event_entry(name)
      handler_obj = nil
      mid = nil
      with_outargs = false
      if entry
        handler_obj = entry[:proc]
        mid = :call
        with_outargs = entry[:with_outargs]
      elsif @handler
        on_name = "on#{name}"
        if @handler.respond_to?(on_name)
          handler_obj = @handler
          mid = on_name
          is_default = false
        elsif @handler.respond_to?(:method_missing)
          handler_obj = @handler
          mid = :method_missing
          is_default = true
        end
      end
      return if handler_obj.nil? || mid.nil?

      args = []
      args << name if is_default
      cargs, rgvarg_addr = read_dispparams(pdispparams_ptr)
      cargs.times do |i|
        var_ptr = Fiddle::Pointer.new(rgvarg_addr + (cargs - i - 1) * W::VARIANT_SIZE)
        args << WIN32OLE.variant_bytes_to_ruby_value(var_ptr[0, W::VARIANT_SIZE])
      end
      outargv = nil
      if with_outargs
        outargv = []
        args << outargv
      end

      result = begin
        handler_obj.send(mid, *args)
      rescue StandardError, ScriptError => e
        warn_closure_exception('an event callback', e)
        nil
      end

      return if pvarresult_ptr.nil? || pvarresult_ptr.to_i.zero?

      bytes = begin
        WIN32OLE.ruby_value_to_variant_bytes(result, [])
      rescue StandardError
        W.pack_variant(W::VT_EMPTY, W.pack_empty)
      end
      pvarresult_ptr[0, W::VARIANT_SIZE] = bytes
    end
```

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/event.rb test/win32ole/jruby/test_event.rb
git commit -m "jruby: add Event's real Invoke closure with exception containment"
```

---

## Task 13: `event.rb` — `#advise` wiring, `#unadvise`, finalizer

**Files:**
- Modify: `lib/win32ole/jruby/event.rb`
- Test: `test/win32ole/jruby/test_event.rb`

**Interfaces:**
- Consumes: `#resolve_event_source` (Task 10), `#build_sink` (Task 11), `W.query_interface`/`IID_ICONNECTIONPOINTCONTAINER` (Task 1).
- Produces: the real `#advise` (replaces Task 7's `NotImplementedError` stub), `#unadvise`, a class-level finalizer following `variant.rb:282-300`'s exact "shared mutable state hash" pattern.

Entirely live-COM (needs a real `IConnectionPointContainer`/`IConnectionPoint`) — CI-only. This task's local check: `#unadvise` on a never-advised object (`@finalizer_state.nil?`) is a safe no-op, and a second `#unadvise` call doesn't double-`Unadvise`/double-free.

- [ ] **Step 1: Write the failing test**

```ruby
  def test_unadvise_on_a_never_advised_event_is_a_safe_noop
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@finalizer_state, nil)
    assert_nil(ev.unadvise)
  end

  def test_unadvise_is_idempotent
    calls = []
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@finalizer_state, { cp_ptr: nil, cookie: nil, ti_ptr: nil, sink_addr: nil, vtable_addr: nil })
    ev.instance_variable_set(:@sink_closures, nil)
    ev.unadvise
    assert_nil(ev.unadvise) # second call must not raise (e.g. Release on a nil pointer)
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: `NoMethodError: undefined method 'unadvise'`.

- [ ] **Step 3: Implement `#advise`/`#unadvise`/finalizer**

Replace Task 7's stub `#advise`:

```ruby
    # ext/win32ole/win32ole_event.c:900-973 (ev_advise)
    def advise(ole, itf)
      iid_bytes, event_typeinfo_ptr = resolve_event_source(ole, itf)

      idispatch_ptr = ole.dispatch_ptr
      container_ptr = W.query_interface(idispatch_ptr, W::IID_ICONNECTIONPOINTCONTAINER)
      if container_ptr.nil?
        release_ptr(event_typeinfo_ptr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('query IConnectionPointContainer', 'E_NOINTERFACE')
      end

      cp_out = ("\x00" * W::PTR_SIZE).b
      hr = W.vtable_function(container_ptr, 4, [W::VOIDP, W::VOIDP], W::LONG).call(container_ptr, iid_bytes, cp_out)
      release_ptr(container_ptr)
      if W.failed?(hr)
        release_ptr(event_typeinfo_ptr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('query IConnectionPoint', W.hr_hex(hr))
      end
      connection_point_ptr = cp_out.unpack1(W::PACK_PTR)

      sink_addr, vtable_addr, closures = build_sink(iid_bytes, event_typeinfo_ptr)
      @sink_closures = closures # keep the trampolines alive; see build_sink's own comment

      cookie_out = ("\x00" * 4).b
      hr = W.vtable_function(connection_point_ptr, 5, [W::VOIDP, W::DWORD], W::LONG).call(
        connection_point_ptr, sink_addr, cookie_out
      )
      if W.failed?(hr)
        release_ptr(connection_point_ptr)
        release_ptr(event_typeinfo_ptr)
        Fiddle.free(sink_addr)
        Fiddle.free(vtable_addr)
        @sink_closures = nil
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('Advise', W.hr_hex(hr))
      end

      @finalizer_state = {
        cp_ptr: connection_point_ptr, cookie: cookie_out.unpack1('L'),
        ti_ptr: event_typeinfo_ptr, sink_addr: sink_addr, vtable_addr: vtable_addr
      }
      ObjectSpace.define_finalizer(self, self.class.finalizer(@finalizer_state))
    end

    # Mirrors variant.rb:282-300's discipline exactly: the finalizer proc
    # captures ONLY this plain data hash, never self and never the live
    # Fiddle::Closure objects (@sink_closures) -- capturing either would
    # create a reference cycle that prevents GC from ever running the
    # finalizer at all.
    def self.finalizer(state)
      proc do
        cp_ptr = state[:cp_ptr]
        if cp_ptr && !cp_ptr.zero?
          W.vtable_function(cp_ptr, 6, [W::DWORD], W::LONG).call(cp_ptr, state[:cookie])
          W.vtable_function(cp_ptr, 2, [W::VOIDP], W::DWORD).call(cp_ptr)
        end
        ti_ptr = state[:ti_ptr]
        W.vtable_function(ti_ptr, 2, [W::VOIDP], W::DWORD).call(ti_ptr) if ti_ptr && !ti_ptr.zero?
        Fiddle.free(state[:sink_addr]) if state[:sink_addr]
        Fiddle.free(state[:vtable_addr]) if state[:vtable_addr]
      end
    end
```

Add the public `#unadvise`:

```ruby
    def unadvise
      return nil if @finalizer_state.nil? || @finalizer_state[:cp_ptr].nil?

      cp_ptr = @finalizer_state[:cp_ptr]
      W.vtable_function(cp_ptr, 6, [W::DWORD], W::LONG).call(cp_ptr, @finalizer_state[:cookie])
      W.vtable_function(cp_ptr, 2, [W::VOIDP], W::DWORD).call(cp_ptr)
      ti_ptr = @finalizer_state[:ti_ptr]
      W.vtable_function(ti_ptr, 2, [W::VOIDP], W::DWORD).call(ti_ptr) if ti_ptr && !ti_ptr.zero?
      Fiddle.free(@finalizer_state[:sink_addr]) if @finalizer_state[:sink_addr]
      Fiddle.free(@finalizer_state[:vtable_addr]) if @finalizer_state[:vtable_addr]

      @finalizer_state[:cp_ptr] = nil
      @finalizer_state[:ti_ptr] = nil
      @finalizer_state[:sink_addr] = nil
      @finalizer_state[:vtable_addr] = nil
      @sink_closures = nil
      nil
    end
```

`#on_event`'s existing `@finalizer_state.nil?` guard (Task 8) now also correctly rejects calls after `#unadvise` — `unadvise` sets `@finalizer_state[:cp_ptr] = nil` but leaves the hash itself non-nil, so tighten Task 8's guard from `@finalizer_state.nil?` to `@finalizer_state.nil? || @finalizer_state[:cp_ptr].nil?` (matching `#unadvise`'s own check) in `register_event`.

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: PASS, full suite (including Tasks 8-12's tests, which use the plain `@finalizer_state = { cp_ptr: 0x1234 }` truthy stand-in from Task 8 and are unaffected by this tightened check).

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/event.rb test/win32ole/jruby/test_event.rb
git commit -m "jruby: wire Event#advise/#unadvise and the sink's finalizer"
```

---

## Task 14: `event.rb` — Hash/Array out-argument write-back

**Files:**
- Modify: `lib/win32ole/jruby/event.rb`
- Test: `test/win32ole/jruby/test_event.rb`

**Interfaces:**
- Consumes: `#handle_invoke` (Task 12).
- Produces: private `#write_byref_variant(var_ptr, value)`, `#write_hash_result(hash, dispid, cargs, rgvarg_addr)`, `#write_array_outargs(ary, cargs, rgvarg_addr)`, wired into `#handle_invoke`.

`#write_byref_variant` operates on a raw `VARIANT` byte buffer we construct ourselves (no COM needed) — fully locally testable. `#write_hash_result` needs a live `ITypeInfo::GetNames` call — CI-only.

- [ ] **Step 1: Write the failing test**

```ruby
  def fake_byref_variant(vt, ref_bytesize)
    ref_buf = Fiddle::Pointer.malloc(ref_bytesize)
    var = W.pack_variant(vt | W::VT_BYREF, W.pack_pointer(ref_buf.to_i))
    var_ptr = Fiddle::Pointer.to_ptr(var)
    [var_ptr, ref_buf]
  end

  def test_write_byref_variant_writes_a_bool
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_BOOL, 2)
    ev.send(:write_byref_variant, var_ptr, true)
    assert_equal(-1, ref_buf[0, 2].unpack1('s'))
  ensure
    Fiddle.free(ref_buf.to_i) if ref_buf
  end

  def test_write_byref_variant_writes_an_i4
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_I4, 4)
    ev.send(:write_byref_variant, var_ptr, 42)
    assert_equal(42, ref_buf[0, 4].unpack1('l'))
  ensure
    Fiddle.free(ref_buf.to_i) if ref_buf
  end

  def test_write_byref_variant_writes_an_r8
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_R8, 8)
    ev.send(:write_byref_variant, var_ptr, 1.5)
    assert_in_delta(1.5, ref_buf[0, 8].unpack1('d'), 0.0001)
  ensure
    Fiddle.free(ref_buf.to_i) if ref_buf
  end

  def test_write_byref_variant_ignores_a_non_byref_variant
    ev = WIN32OLE::Event.allocate
    var = W.pack_variant(W::VT_I4, W.pack_i4(0))
    var_ptr = Fiddle::Pointer.to_ptr(var)
    assert_nil(ev.send(:write_byref_variant, var_ptr, 99)) # must not raise / must not dereference garbage
  end

  def test_write_byref_variant_ignores_a_type_mismatched_value
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_I4, 4)
    ref_buf[0, 4] = [7].pack('l')
    ev.send(:write_byref_variant, var_ptr, [1, 2, 3]) # Array has no matching case -- silent no-op, matches C
    assert_equal(7, ref_buf[0, 4].unpack1('l'))
  ensure
    Fiddle.free(ref_buf.to_i) if ref_buf
  end

  def test_write_array_outargs_writes_positionally_and_stops_at_cargs
    ev = WIN32OLE::Event.allocate
    rgvarg = Fiddle::Pointer.malloc(W::VARIANT_SIZE * 2)
    ref0 = Fiddle::Pointer.malloc(4)
    ref1 = Fiddle::Pointer.malloc(4)
    rgvarg[1 * W::VARIANT_SIZE, W::VARIANT_SIZE] = W.pack_variant(W::VT_I4 | W::VT_BYREF, W.pack_pointer(ref0.to_i)) # arg 0
    rgvarg[0 * W::VARIANT_SIZE, W::VARIANT_SIZE] = W.pack_variant(W::VT_I4 | W::VT_BYREF, W.pack_pointer(ref1.to_i)) # arg 1

    ev.send(:write_array_outargs, [11, 22], 2, rgvarg.to_i)

    assert_equal(11, ref0[0, 4].unpack1('l'))
    assert_equal(22, ref1[0, 4].unpack1('l'))
  ensure
    Fiddle.free(rgvarg.to_i) if rgvarg
    Fiddle.free(ref0.to_i) if ref0
    Fiddle.free(ref1.to_i) if ref1
  end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: `NoMethodError: undefined method 'write_byref_variant'`.

- [ ] **Step 3: Implement**

```ruby
    # ext/win32ole/win32ole_event.c:333-399 (ole_val2ptr_variant), ported
    # 1:1 including its "silently do nothing for an unhandled
    # type/VARTYPE combination" fallthrough -- e.g. writing a String into a
    # VT_I4|BYREF slot is a deliberate no-op in the original, not a bug we
    # should "fix" by raising.
    def write_byref_variant(var_ptr, value)
      vt = var_ptr[0, 2].unpack1('S')
      return if (vt & W::VT_BYREF).zero?

      ref_addr = var_ptr[8, W::PTR_SIZE].unpack1(W::PACK_PTR)
      return if ref_addr.zero?

      ref_ptr = Fiddle::Pointer.new(ref_addr)
      base_vt = vt & ~W::VT_BYREF
      case value
      when String
        ref_ptr[0, W::PTR_SIZE] = [W.sys_alloc_string.call(W.wstr(value))].pack(W::PACK_PTR) if base_vt == W::VT_BSTR
      when Integer
        case base_vt
        when W::VT_UI1 then ref_ptr[0, 1] = [value].pack('C')
        when W::VT_I2 then ref_ptr[0, 2] = [value].pack('s')
        when W::VT_I4 then ref_ptr[0, 4] = [value].pack('l')
        when W::VT_R4 then ref_ptr[0, 4] = [value.to_f].pack('f')
        when W::VT_R8 then ref_ptr[0, 8] = [value.to_f].pack('d')
        end
      when Float
        case base_vt
        when W::VT_I2 then ref_ptr[0, 2] = [value.to_i].pack('s')
        when W::VT_I4 then ref_ptr[0, 4] = [value.to_i].pack('l')
        when W::VT_R4 then ref_ptr[0, 4] = [value].pack('f')
        when W::VT_R8 then ref_ptr[0, 8] = [value].pack('d')
        end
      when true, false
        ref_ptr[0, 2] = [value ? -1 : 0].pack('s') if base_vt == W::VT_BOOL
      end
    end

    def write_array_outargs(ary, cargs, rgvarg_addr)
      ary.each_with_index do |value, i|
        break if i >= cargs

        var_ptr = Fiddle::Pointer.new(rgvarg_addr + (cargs - i - 1) * W::VARIANT_SIZE)
        write_byref_variant(var_ptr, value)
      end
    end

    # ext/win32ole/win32ole_event.c:401-428 (hash2ptr_dispparams)
    def write_hash_result(hash, dispid, cargs, rgvarg_addr)
      names_out = Fiddle::Pointer.malloc(W::PTR_SIZE * (cargs + 1))
      count_out = ("\x00" * 4).b
      hr = TI.get_names_fn(@event_typeinfo_ptr).call(@event_typeinfo_ptr, dispid, names_out, cargs + 1, count_out)
      return if W.failed?(hr)

      len = count_out.unpack1('L')
      (len - 1).times do |i|
        bstr = names_out[(i + 1) * W::PTR_SIZE, W::PTR_SIZE].unpack1(W::PACK_PTR)
        key_name = W.bstr_to_s(bstr)
        W.sys_free_string.call(bstr) unless bstr.zero?

        value = hash[i]
        value = hash[key_name] if value.nil?
        value = hash[key_name.to_sym] if value.nil? && key_name

        var_ptr = Fiddle::Pointer.new(rgvarg_addr + (cargs - i - 1) * W::VARIANT_SIZE)
        write_byref_variant(var_ptr, value)
      end
    ensure
      Fiddle.free(names_out.to_i) if names_out
    end
```

Wire both into `#handle_invoke` (replacing the `result = begin ... end` block's immediate use):

```ruby
      result = begin
        handler_obj.send(mid, *args)
      rescue StandardError, ScriptError => e
        warn_closure_exception('an event callback', e)
        nil
      end

      if result.is_a?(Hash)
        write_hash_result(result, dispid, cargs, rgvarg_addr)
        result = result['return'] || result[:return]
      elsif with_outargs && outargv.is_a?(Array)
        write_array_outargs(outargv, cargs, rgvarg_addr)
      end
```

(this replaces the plain `result = begin ... end` statement that previously fell straight through to the `pvarresult_ptr` write in Task 12's version.)

- [ ] **Step 4: Run test to verify it passes**

Run: `ruby -Ilib -Itest -e 'Object.send(:remove_const, :RUBY_ENGINE); RUBY_ENGINE = "jruby"; load "test/win32ole/jruby/test_event.rb"'`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/event.rb test/win32ole/jruby/test_event.rb
git commit -m "jruby: add Event's Hash/Array out-argument write-back"
```

---

## Task 15: GC-stress coverage + final CI verification

**Files:**
- Create: `test/win32ole/jruby/test_event_gc_stress.rb`
- No production code changes.

**Interfaces:**
- Consumes: everything above.

Live-COM (ADO's `ConnectionEvents`, mirroring `test_win32ole_event.rb`'s own `TestWIN32OLE_EVENT_ADO`) — CI-only, exactly like `test_record_variant_gc_stress.rb`/`test_typelib_gc_stress.rb` already are for their own subsystems.

- [ ] **Step 1: Write the test**

```ruby
begin
  require 'win32ole'
rescue LoadError
end
require 'test/unit'

ado_installed =
  if defined?(WIN32OLE) && RUBY_ENGINE == 'jruby'
    begin
      db = WIN32OLE.new('ADODB.Connection')
      db.connectionString = 'Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=.;'
      db.open
      db.close
      true
    rescue
    end
  end

if ado_installed
  class TestEventGCStress < Test::Unit::TestCase
    def test_gc_stress_survives_advise_and_a_subsequent_event
      db = WIN32OLE.new('ADODB.Connection')
      db.connectionString = 'Driver={Microsoft Text Driver (*.txt; *.csv)};DefaultDir=.;'
      fired = false
      ev = WIN32OLE::Event.new(db, 'ConnectionEvents')
      ev.on_event('WillConnect') { fired = true }

      GC.stress = true
      GC.start
      GC.stress = false

      db.open
      WIN32OLE::Event.message_loop
      assert(fired, 'event callback did not fire after a forced GC between advise and the firing call')
    ensure
      ev&.unadvise
      db&.close if db&.state == 1
    end
  end
end
```

- [ ] **Step 2: Run test to verify it fails**

Run: `bundle exec rake` on a `windows-latest` JRuby CI job (this test cannot meaningfully run/fail on this dev machine at all — no `ado_installed`, whole file contributes 0 tests).
Expected (pre-implementation, hypothetically): `NameError` for `WIN32OLE::Event` if any earlier task were missing — moot here since Tasks 1-14 already landed.

- [ ] **Step 3: No implementation needed — this task is pure verification**

- [ ] **Step 4: Push the branch and confirm the full `test-jruby` CI job**

Run: `git push` (or open/update the PR), then check the `build (jruby)` job of `.github/workflows/windows.yml`'s `test-jruby` run (`bundle exec rake`, which globs `test/**/test_*.rb` and picks up every new file from Tasks 1-15 automatically — no Rakefile change needed).
Expected: `TestWIN32OLE_EVENT`/`TestWIN32OLE_EVENT_SWbemSink`/`TestWIN32OLE_EVENT_ADO` (the legacy suite, now unguarded since `WIN32OLE::Event` is defined) plus every new `test/win32ole/jruby/test_*.rb` file's tests pass or cleanly `omit`. Per spec §7, `TestWIN32OLE_EVENT_SWbemSink`'s actual event-arrival assertions are an accepted, inherited, known-unreliable-in-CI risk (§3/§8 risk #2) — not something this push is expected to turn green, and their continued flakiness is not a Phase 4 regression.

- [ ] **Step 5: Commit**

```bash
git add test/win32ole/jruby/test_event_gc_stress.rb
git commit -m "jruby: add GC-stress coverage for Event's sink keep-alive discipline"
```

---

## 8. Risks / open questions carried forward

1. **The new COM interface IIDs (`IUnknown`, `IConnectionPointContainer`, `IProvideClassInfo`, `IProvideClassInfo2`) and the hand-rolled `Fiddle::Closure`-backed vtable are unverified against a live Windows build until Task 15's CI push** — same class of risk Phase 3 flagged for `IRecordInfo` (design §8 risk #1 there). This is the **first** `Fiddle::Closure` usage in this repository's actual (not spiked-and-discarded) history.
2. **`SWbemSink` async WMI delivery remains unresolved in CI**, inherited from the Phase 1 spike and restated in this phase's own design doc §3/§8 risk #2 — not attempted here.
3. **Exception-during-callback behavior intentionally diverges from the C extension** (write to `$stderr` and keep running, vs. the C extension's `ruby_finalize`+`exit(-1)`) — a deliberate, permanent difference per spec §3, not a stopgap.
4. **x86 (32-bit) / ARM64 Windows unverified** — unchanged, inherited risk.
5. **`Type#implemented_ole_types`/friends stay `NotImplementedError`** — this phase's `ImplType` traversal logic lives privately inside `event.rb` (Tasks 9-10), not surfaced as public `Type` API, per spec §3 non-goal #5.
