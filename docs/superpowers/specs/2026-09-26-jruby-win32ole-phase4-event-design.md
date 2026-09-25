# JRuby support for win32ole — Phase 4 (Event) design

Date: 2026-09-26
Status: proposed

## 1. Background

Phase 1–3 built `WIN32OLE` core dispatch, `WIN32OLE::TYPE`/`TYPELIB`/
`METHOD`/`PARAM`/`VARIABLE` introspection, and `WIN32OLE::Record`/
`Variant`, all as a pure-Ruby + Fiddle client of COM objects. Every
interface used so far (`IDispatch`, `ITypeInfo`, `ITypeLib`,
`IRecordInfo`) is one this code only *calls into* — the OLE server
always owns the vtable.

`WIN32OLE::Event` inverts that: this process must act as a COM
**server**. It builds a connection-point event sink (an object
implementing `IDispatch`, handed to `IConnectionPoint::Advise`) so an
OLE server can call *into* our process when an event fires. The design
doc's Phase 1 spike (`2026-09-22-jruby-win32ole-support-design.md`
§1.2) validated the core mechanism — a hand-rolled `IUnknown`/
`IDispatch` vtable backed by `Fiddle::Closure::BlockCaller` native
callbacks, handed to a real out-of-process COM server — on both MRI
and JRuby, but that spike code was never committed to this repository
and no wrapper class was built around it.

### 1.1 What the C extension actually does (verified against
`ext/win32ole/win32ole_event.c`)

- A static `IEventSinkVtbl` (7 slots: `QueryInterface`, `AddRef`,
  `Release`, `GetTypeInfoCount`, `GetTypeInfo`, `GetIDsOfNames`,
  `Invoke`) backs every sink instance. `GetTypeInfoCount`/`GetTypeInfo`
  are stubs (`*pct = 0`; `DISP_E_BADINDEX`); `GetIDsOfNames` delegates
  to the resolved event `ITypeInfo`; `Invoke` is where all the real
  work happens.
- Sink identity: a small heap struct (`IEVENTSINKOBJ`) holding the
  vtable pointer, a refcount, the source IID, an index into a
  process-global `ary_ole_event` array (`m_event_id`), and the
  resolved event-interface `ITypeInfo*`.
- Interface/event-source resolution (`find_iid`, `find_coclass`,
  `find_default_source_from_typeinfo`, `find_default_source`) — see
  §4.2 below, ported near-verbatim.
- `ev_advise`: resolve the IID → `QueryInterface` the target for
  `IConnectionPointContainer` → `FindConnectionPoint(iid)` →
  `Advise(sink, &cookie)`.
- `EVENTSINK_Invoke`: resolve the firing DISPID to an event name via
  `ITypeInfo::GetNames`, look up a registered callback (block or
  handler-object method), build the Ruby argument list from
  `pdispparams->rgvarg` (reversed, like every other DISPPARAMS access
  in this codebase), run it, and handle a `Hash` or (for
  `on_event_with_outargs`) trailing-`Array` return by writing back
  into the VARIANT `BYREF` out-parameters.
- A process-wide `ary_ole_event` (`rb_gc_register_mark_object`) is
  purely a workaround for C having no closures — Invoke is a bare
  function pointer with no captured state, so it needs *some* way to
  get from "which sink fired" back to a Ruby object. This does not
  carry over to the Ruby port (§4.1).

### 1.2 The `WIN32OLE.connect` / `WIN32OLE.const_load` gap

Neither method exists yet in `lib/win32ole/jruby/win32ole.rb`. Both
were explicitly out of scope for Phase 1 (design doc §6.1) and were
never picked up by Phase 2 or 3. Every behavioral test in
`test/win32ole/test_win32ole_event.rb` beyond the two bare constructor-
error checks (`TestWIN32OLE_EVENT`) depends on one or the other:
`TestWIN32OLE_EVENT_SWbemSink` calls `WIN32OLE.connect`;
`TestWIN32OLE_EVENT_ADO` calls `WIN32OLE.const_load`. Phase 4 includes
both as prerequisites — implementing `Event` without them would leave
almost the entire real test suite unexercised.

## 2. Goals

- Implement `WIN32OLE::Event`: construct via an existing `WIN32OLE`
  object (optionally naming a source interface), register callbacks
  with `#on_event`/`#on_event_with_outargs`, remove them with
  `#off_event`, disconnect with `#unadvise`, support a delegate
  `#handler=`/`#handler` object, and pump the Windows message queue
  with `WIN32OLE::Event.message_loop`.
- Implement `WIN32OLE.connect` (attach to a running OLE server
  instance) and `WIN32OLE.const_load` (load a typelib's constants into
  a Ruby module) as prerequisites for exercising `Event` against the
  existing MRI test suite.
- Reuse the existing Phase 1–3 substrate (`win32.rb`, `dispatch.rb`,
  `typeinfo.rb`, `type.rb`, `typelib.rb`, `variable.rb`) wherever it
  already does what's needed; extend it with the same naming/memoization
  conventions where it doesn't (generic `QueryInterface`, a few more
  `ITypeInfo`/`ITypeLib` vtable slots, `user32.dll` message-loop
  functions).

## 3. Non-goals (for this document / Phase 4)

- **Exception-in-callback parity with the C extension.** MRI prints
  the backtrace and calls `exit(-1)` (via `ruby_finalize`) when a Ruby
  exception escapes an event callback, because letting a C++ exception
  unwind back across a COM call boundary is undefined behavior for the
  calling OLE server. This port instead writes the exception to
  `$stderr` and lets `Invoke` return `NOERROR`, keeping the process (and
  the message loop) alive. Unwinding a Ruby/JRuby exception across a
  `Fiddle::Closure::BlockCaller` boundary into arbitrary native calling
  code is exactly the kind of cross-boundary unwind Phase 1 (§4.5) and
  Phase 3 already avoid elsewhere; this is a deliberate, permanent
  behavior difference, not a stopgap.
- **CI verification of `WIN32OLE::Event::SWbemSink`'s asynchronous WMI
  delivery.** The Phase 1 spike (design doc §1.2) already found this
  doesn't arrive within 30s on GitHub Actions `windows-latest`,
  independent of the vtable/closure mechanism itself (everything up to
  the actual async notification — vtable construction, marshaling the
  sink across the process boundary, `PeekMessageW`/`DispatchMessageW`
  — worked). Phase 4 does not attempt to root-cause or fix this; the
  existing `swbemsink_available`-gated tests continue to be exercised
  as thoroughly as they can be (real dispid resolution, real COM calls)
  and are expected to leave the WMI-specific event-firing assertions
  unverified in CI, as inherited risk. `TestWIN32OLE_EVENT_ADO`
  (`ConnectionEvents`, which fires synchronously within the same
  process during `@db.open`) is the primary CI-verified path for the
  actual event-firing/argument-marshaling/out-arg/hash-return logic.
- **`WIN32OLE.connect`'s `host` (remote DCOM) parameter** — mirrors
  Phase 1's existing `WIN32OLE.new` restriction; raise
  `NotImplementedError` for a non-nil host rather than silently
  ignoring it.
- **Windows-on-ARM64 / x86 (32-bit) verification** — inherited,
  unresolved risk from the Phase 1 design doc (§8.3), unchanged by
  Phase 4.
- **`Type#implemented_ole_types`/`#default_ole_types`/
  `#source_ole_types`/`#default_event_sources`** stay
  `NotImplementedError` ("Phase 2 non-goal"). Phase 4's event-source
  resolution (§4.2) needs the same underlying `ImplType` vtable calls,
  but implements them as private helpers inside `event.rb`, not as a
  public `Type` API — the existing tests that exercise those four
  public methods (`test_win32ole_type_event.rb` and friends) are gated
  on `AvailableOLE.sysmon_available?`, which itself calls the
  currently-`NotImplementedError` `WIN32OLE::Type.new(typelib, class)`
  and is expected to keep omitting in CI regardless. Wiring the public
  methods up is a natural follow-on, not required for Phase 4's own
  goals.

## 4. Architecture

### 4.1 File layout

```
lib/win32ole/jruby/event.rb      # WIN32OLE::Event                   (new)
lib/win32ole/jruby/win32.rb      # + generic QueryInterface, user32.dll message-loop fns
lib/win32ole/jruby/typeinfo.rb   # + GetImplTypeFlags/GetRefTypeOfImplType/
                                  #   GetTypeInfoOfGuid/GetNames _fn helpers
lib/win32ole/jruby/win32ole.rb   # + .connect, .const_load
```

`lib/win32ole/jruby.rb` gains `require 'win32ole/jruby/event'`.

### 4.2 The event sink: `Fiddle::Closure`, not a global registry

The sink is a persistent native buffer holding one pointer (to a
7-entry vtable, itself a persistent native buffer of 7 function
pointers). Each of the 7 functions is a `Fiddle::Closure::BlockCaller`
built with `STDCALL` (§Win32 substrate) and the same argument-type
signature as its `IEventSinkVtbl` counterpart in the C source (§1.1).

Unlike the C extension, each closure's block is a real Ruby closure
over the owning `WIN32OLE::Event` instance — created fresh per
`#advise` call, not looked up from a shared table by index. This
removes the entire `ary_ole_event`/`m_event_id`/`evs_push`/
`evs_entry`/`evs_delete` apparatus; the event-name → callback lookup
(§4.4) becomes a plain private method call on `self`.

**Lifetime**: the sink buffer, the vtable buffer, and the 7 `Closure`
objects are held in instance variables on the `Event` object
(`@sink_ptr`, `@vtable_ptr`, `@sink_closures`) for as long as the
connection point is advised — the same "native address embedded as
data needs its owning Ruby object kept alive" discipline as Phase 1
§4.5 and Phase 3 §4.7. `#unadvise` and the finalizer both clear these
after calling `IConnectionPoint::Unadvise`; a `GC.stress` test (§7)
probes this the same way Phase 1/3 already do for their own
keep-alive paths.

`QueryInterface`'s closure always returns the sink's own address for
`IID_IUnknown`/`IID_IDispatch`/the resolved source IID (mirroring the
C extension exactly — this sink never implements more than one real
interface identity). `AddRef`/`Release` maintain a plain Ruby integer
refcount captured by the closure's block (no native struct field
needed for this, unlike the C version, since nothing outside the
closures ever reads it).

### 4.3 Event-source resolution (`WIN32OLE::Event.new(ole, itf = nil)`)

Ported near-verbatim from `find_iid`/`find_coclass`/
`find_default_source_from_typeinfo`/`find_default_source`
(`win32ole_event.c` lines 481–786), as private `event.rb` helpers
operating on the already-existing `TypeInfo`/`Type`/`TypeLib`
primitives:

- `itf` given: `ole`'s `IDispatch` → `GetTypeInfo(0)` →
  `GetContainingTypeLib` → scan every type in that typelib for a
  `TKIND_COCLASS` whose implemented types include one named `itf`
  (`GetDocumentation(-1)` on each impl type) → that impl type's GUID
  is the source IID.
- `itf` omitted: try `IProvideClassInfo2::GetGUID
  (GUIDKIND_DEFAULT_SOURCE_DISP_IID)` first (this is the path
  `ADODB.Connection` actually takes) → resolve that IID the same way
  as the `itf`-given case. If `IProvideClassInfo2` isn't supported (or
  fails), fall back to `IProvideClassInfo::GetClassInfo` or plain
  `GetTypeInfo(0)`, then walk the resulting type's `ImplType` entries
  for one flagged both `[default]` and `[source]`
  (`GetImplTypeFlags`/`GetRefTypeOfImplType`/`GetRefTypeInfo`); if the
  starting type itself isn't a COCLASS with such an entry, search the
  same typelib for a COCLASS that implements it (`find_coclass`) and
  retry.
- Resolved IID → `ole`'s `IDispatch` → `QueryInterface` for
  `IConnectionPointContainer` → `FindConnectionPoint(iid)` →
  `Advise(sink, &cookie)`. A failure at any step raises
  `WIN32OLE::RuntimeError` ("interface not found") or
  `WIN32OLE::QueryInterfaceError`, matching the C messages closely
  enough for message-content assertions in the test suite.

New substrate needed: a generic `Win32.query_interface(obj_addr,
iid_bytes)` (vtable slot 0, usable against any COM pointer — the
existing code has never needed one before now, only the fixed
`IDispatch`/`ITypeInfo`/`ITypeLib` vtable slot helpers) plus
`TypeInfo` `_fn` helpers for `GetImplTypeFlags` (slot 9),
`GetRefTypeOfImplType` (slot 8), `GetTypeInfoOfGuid` (`ITypeLib` slot
6), and `GetNames` (`ITypeInfo` slot 7) — all four vtable slots are
already listed in `ITYPEINFO_VTBL`/`ITYPELIB_VTBL`, only the memoized
`Fiddle::Function` wrapper is missing, following the exact pattern
every other `_fn` method already uses.

### 4.4 `Invoke`: firing an event

1. `GetNames(dispid, &bstr, 1)` on the resolved event `ITypeInfo` →
   the event name.
2. Look up a registered callback: the `Event`'s own event table (an
   `Array` of `{name:, proc:, with_outargs:}`, appended by `#on_event`/
   `#on_event_with_outargs`, matching by name or falling back to a
   nameless "catch-all" entry exactly like `ole_search_event`) takes
   priority; if none matches, fall back to `#handler`'s `onXXX` method
   (or `method_missing`, flagged as the "default handler" case that
   also gets the event name prepended to its argument list — same as
   `ole_search_handler_method`).
3. Build the Ruby argument array from `pdispparams->rgvarg`, iterated
   in reverse (same convention `dispatch.rb#ole_invoke` already uses
   for the outgoing direction) and converted with the existing
   `WIN32OLE.variant_bytes_to_ruby_value`. `on_event_with_outargs`
   appends a trailing mutable `Array` the callback can write into.
4. Run the callback. A raised exception is written to `$stderr` (with
   backtrace) and swallowed — §3 non-goal, not a C-parity choice.
5. If the callback returned a `Hash`: for each parameter name (from
   `GetNames(dispid, cArgs+1)`) look up a value by index, string key,
   or symbol key, and write it into the corresponding `BYREF`
   VARIANT's pointee via a new `write_byref_variant(var_ptr, value)`
   helper (a direct port of `ole_val2ptr_variant`'s `VT_BSTR|BYREF` /
   `VT_UI1|I2|I4|R4|R8|BOOL|BYREF` cases); a `'return'`/`:return` key
   becomes the VARIANT result. If `on_event_with_outargs` was used and
   the callback returned an `Array` instead, write each element
   positionally the same way (`ary2ptr_dispparams`).

### 4.5 `WIN32OLE.connect` / `WIN32OLE.const_load`

- `.connect(server)`: resolve CLSID the same way `WIN32OLE.new`
  already does (`resolve_clsid`, reused as-is) → `GetActiveObject`
  (`oleaut32.dll`, new binding — distinct from `CoCreateInstance`,
  which always creates a *new* instance) → wrap the returned
  `IDispatch*` with the existing `WIN32OLE.wrap_dispatch_pointer`.
  `host` non-nil raises `NotImplementedError` (§3).
- `.const_load(ole, mod)`: `ole.ole_type.ole_typelib` (existing
  `WIN32OLE#ole_type`/`Type#ole_typelib`) → `TypeLib#ole_types` →
  for each `Type`, its `#variables` (existing `WIN32OLE::Variable`,
  Phase 2) filtered to `varkind == VARKIND_CONSTANT` → `mod.const_set`
  keyed on the variable's name, value read via `Variable#value`
  (already implemented). Constants already defined on `mod` are left
  alone (mirrors the C extension not re-defining, and the test suite's
  `defined?(ADO::AdStateOpen)` guard in `setup`).

## 5. Per-class API (Phase 4 scope)

| Method | Notes |
|---|---|
| `WIN32OLE::Event.new(ole, itf = nil)` | §4.3. `TypeError` if `ole` isn't a `WIN32OLE`; `WIN32OLE::RuntimeError` if the interface/source can't be resolved. |
| `WIN32OLE::Event.message_loop` | `PeekMessageW`/`TranslateMessage`/`DispatchMessageW` drain loop (new `user32.dll` bindings). |
| `#on_event([event]) { ... }` | Registers/replaces the callback for `event` (or the catch-all if omitted). `String` or `Symbol`. |
| `#on_event_with_outargs([event]) { ... }` | Same, but the block's last argument is a mutable out-arg array (§4.4 step 3). |
| `#off_event([event])` | Removes a registered callback. |
| `#unadvise` | Disconnects; subsequent `#on_event` raises `WIN32OLE::RuntimeError`. |
| `#handler=(obj)` / `#handler` | Delegate object for `onXXX`/`method_missing`-style handling (§4.4 step 2). |

`WIN32OLE.connect(server, host = nil)` and `WIN32OLE.const_load(ole,
mod)` (§4.5) round out the prerequisites.

## 6. Error translation

Reuses the existing `WIN32OLE::RuntimeError`/`QueryInterfaceError`
classes and `Win32.query_interface_error_message`/
`unknown_server_error_message` formatting helpers as-is; no new error
classes. `#advise` failures use the same
`"interface not found"`/`hresult_detail` shape the C extension's
`ole_raise` produces, close enough for the test suite's
`assert_raise(RuntimeError)` (no message-content assertions on this
path in `test_win32ole_event.rb`).

## 7. Testing / CI strategy

- `test/win32ole/test_win32ole_event.rb` needs no changes — its
  top-level `if defined?(WIN32OLE::Event)` guard activates
  automatically once Phase 4 lands.
- Primary CI-verified path: `TestWIN32OLE_EVENT_ADO` (gated on
  `ado_installed`, which itself now depends on `.connect`/
  `.const_load`-adjacent machinery only indirectly — `ado_installed`
  just constructs/opens/closes a `WIN32OLE.new('ADODB.Connection')`).
  Exercises `#advise` via the `IProvideClassInfo2` path,
  `#on_event`/`#off_event`/`#unadvise`/`#handler=`, and both `Hash`-
  and `Array`-based out-argument write-back, against a real
  synchronously-firing event (`WillConnect` fires inline during
  `@db.open`, in-process, so it doesn't depend on the message-queue
  timing that made `SWbemSink` unreliable in CI).
- `TestWIN32OLE_EVENT_SWbemSink` (gated on `swbemsink_available`,
  itself now reachable since `WIN32OLE.new('WbemScripting.SWbemSink')`
  already works pre-Phase-4) continues to run as far as it can; its
  actual event-arrival assertions are accepted as an inherited, known-
  unreliable-in-CI risk (§3), not something this phase's test run is
  expected to turn green.
- New `GC.stress` coverage for the sink's keep-alive discipline (§4.2):
  advise, force a `GC.start`, confirm the connection is still alive
  (e.g. a subsequent event still reaches the callback) before
  `#unadvise`.
- No new skip/guard patterns needed beyond what's already in the test
  file (`ado_installed`, `swbemsink_available`).

## 8. Risks / open questions carried forward

1. **`Fiddle::Closure`-backed vtable has no in-repository precedent.**
   The Phase 1 spike validated the mechanism on both MRI x64-mingw and
   JRuby (design doc §1.2), but that spike code was never committed —
   Phase 4 is the first time this pattern is actually implemented and
   tested in this codebase's own CI.
2. **`SWbemSink` async WMI delivery unresolved in CI** (§3, inherited
   from design doc §8.1) — accepted risk, not attempted here.
3. **Exception-during-callback behavior intentionally diverges from
   the C extension** (§3) — write to `$stderr` and continue, rather
   than terminate the process. Flagged in case a future consumer
   relies on the C extension's terminate-on-exception behavior as a
   safety net.
4. **x86 (32-bit) / ARM64 Windows unverified** — unchanged, inherited
   risk (design doc §8.3).
5. **`Type#implemented_ole_types` and friends stay unimplemented**
   (§3) — Phase 4 builds equivalent traversal logic privately inside
   `event.rb` rather than surfacing it publicly; a future phase could
   factor the shared traversal out into `Type` once there's a second
   caller.
