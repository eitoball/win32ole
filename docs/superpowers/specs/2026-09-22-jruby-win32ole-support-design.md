# JRuby support for win32ole — design

Date: 2026-09-22
Status: proposed

## 1. Background

`win32ole` is a Windows-only gem providing OLE Automation (COM) access
from Ruby. The entire implementation lives in `ext/win32ole/*.c`
(~10,500 lines), a C extension that calls the Win32 COM APIs
(`ole32.dll`, `oleaut32.dll`) directly. Because JRuby does not support
native C extensions, `require 'win32ole'` currently does not work at
all under JRuby.

JRuby itself ships a stub (`lib/ruby/stdlib/win32ole.rb` in the JRuby
distribution) that just tries to load a separate `jruby-win32ole` gem.
That gem (`enebo/jruby-win32ole`) implements the same API on top of
**racob**, a 2011-vintage, unmaintained fork of the JNI-based Jacob
Java/COM bridge, bundling prebuilt `.dll`s for x86/x64. Its last
release was in 2017.

### 1.1 Prior art surveyed

- `ruby/digest`, `ruby/stringio`, `ruby/psych`: ship a `java`-platform
  gem variant (`spec.platform = 'java'` when `RUBY_ENGINE == 'jruby'`)
  containing a Java reimplementation, built via JRuby's own Java
  extension mechanism. This works because these libraries wrap
  algorithms with a pure-Java equivalent (`java.security.MessageDigest`,
  SnakeYAML, etc.). COM Automation has no Java-standard equivalent, so
  this pattern doesn't directly apply — a Java implementation would
  still need a native (JNI) COM bridge, reintroducing exactly the
  prebuilt-binary problem that made `jruby-win32ole` a stale, separate
  gem in the first place.
- `ruby/fiddle`: `lib/fiddle.rb` does
  `RUBY_ENGINE == 'ruby' ? require('fiddle.so') : require('fiddle/ffi_backend')`.
  On JRuby, Fiddle is implemented on top of the `ffi` gem. Fiddle is a
  default gem (part of the standard distribution), so depending on it
  does not add a new external runtime dependency the way depending on
  `ffi` directly would.
- `ruby/io-console` (used by `reline`/`irb`): ships a
  `jruby/lib/io/console/backend/ffi/windows.rb` file that calls
  `kernel32.dll` Windows Console APIs (`GetConsoleMode`,
  `ReadConsoleInputW`, ...) directly via the `ffi` gem, using plain
  packed byte strings (`Array#pack`/`String#unpack`) rather than
  `FFI::Struct` classes for the Win32 structs it marshals. This is the
  closest existing production precedent for "drive a native Windows
  API from JRuby via FFI/Fiddle, no C extension," and its style is
  what this design follows for VARIANT/DISPPARAMS handling.

### 1.2 Spike results

A throwaway spike (branch `jruby-support`, `tmp_spike/`) validated the
core mechanism before this design was written:

- `tmp_spike/fiddle_com_spike.rb`: pure Fiddle (no `ffi` required
  directly — transitively via Fiddle's JRuby backend) successfully
  drives `CLSIDFromProgID` → `CoCreateInstance` →
  `IDispatch::GetIDsOfNames` → `IDispatch::Invoke` for 0-, 1-, and
  2-argument calls, against two different real COM objects
  (`Scripting.FileSystemObject`, `Scripting.Dictionary`), identically
  on MRI (x64 mingw) and JRuby, in GitHub Actions `windows-latest`.
  - Found and fixed a real, easy-to-miss bug: `sizeof(VARIANT)` is
    **16 bytes on x86 but 24 bytes on x64** (the value union's widest
    member on x64 is an anonymous 2-pointer record struct, not the
    8-byte scalars/pointers most people picture). Using a 16-byte
    stride only breaks once a `VARIANTARG` array has 2+ elements —
    single-argument calls work by accident, masking the bug.
  - Found that `Fiddle::Function::STDCALL` is undefined on x64 mingw
    MRI (stdcall and cdecl are ABI-identical on x64, so the constant
    simply isn't compiled in for that target) but *is* defined on
    JRuby's ffi-backed Fiddle. A portable implementation must fall
    back to `Fiddle::Function::DEFAULT` when `STDCALL` isn't defined.
- `tmp_spike/fiddle_com_event_spike.rb`: built a hand-rolled
  IUnknown/IDispatch vtable backed by `Fiddle::Closure::BlockCaller`
  native callbacks (a from-scratch COM *server*, not just a client),
  and handed a pointer to it to `WbemScripting.SWbemServices`'s
  `ExecNotificationQueryAsync` (the same mechanism
  `test/win32ole/test_win32ole_event.rb` itself uses for event tests).
  - Vtable construction, marshaling the sink across the process
    boundary to the out-of-process WMI service (`VT_DISPATCH` arg,
    HRESULT=0), and the Win32 message queue (`PeekMessageW`/
    `DispatchMessageW`) all worked correctly on both engines (verified
    with a self-test: a message posted to our own thread via
    `PostThreadMessageW` was successfully retrieved).
  - The actual asynchronous WMI notification never arrived within 30s
    on the GitHub Actions `windows-latest` runner, even when the
    triggering condition (a new process being created) was caused
    deliberately by the test itself. This looks like an environment/
    permissions constraint of that specific CI runner rather than a
    flaw in the closure/vtable mechanism, but it was **not** resolved
    within the spike, and no other environment was tried.

## 2. Goals

- Make `require 'win32ole'` work on JRuby, implemented in pure Ruby +
  Fiddle, living in the same `win32ole` gem MRI already ships.
- Reach eventual feature parity with the MRI C extension.
- Keep the MRI code path (the existing C extension) completely
  untouched; the JRuby path is purely additive.

## 3. Non-goals (for this document / Phase 1)

- Full feature parity in the first shipped increment — see §5 for
  phasing.
- Verifying `WIN32OLE::Event` actually fires end-to-end in CI — the
  spike left this open (§1.2); Phase 4 inherits this as a known risk,
  not something this document resolves.
- Windows-on-ARM64 or x86 (32-bit) verification — the spike only ran
  on GitHub Actions `windows-latest` (x64). x86 needs the VARIANT
  16-byte-stride path exercised on real 32-bit hardware/CI before it
  can be trusted.
- Performance parity with the C extension.

## 4. Architecture

### 4.1 Engine dispatch

`lib/win32ole.rb` currently just requires the compiled C extension.
It gains a branch:

```ruby
if RUBY_ENGINE == 'jruby'
  require 'win32ole/jruby'
else
  require 'win32ole.so'
end
```

No other MRI-facing file changes. The gemspec's
`spec.extensions = "ext/win32ole/extconf.rb"` stays as-is; on JRuby
that extension task presumably no-ops or is skipped in the standard
JRuby gem-install flow (needs confirming during implementation, but
is out of scope for this design — it's a packaging detail, not an
architectural one).

### 4.2 File layout

```
lib/win32ole/jruby.rb          # entry point; loads everything below
lib/win32ole/jruby/
  win32.rb                     # Fiddle bindings: ole32/oleaut32 functions,
                                # VARIANT/DISPPARAMS/GUID byte-packing helpers
  dispatch.rb                  # GetIDsOfNames/Invoke/QueryInterface helpers
                                # for any raw IDispatch pointer
  win32ole.rb                  # WIN32OLE class                    (Phase 1)
  type.rb, typelib.rb,
  method.rb, param.rb,
  variable.rb                  # WIN32OLE::TYPE family             (Phase 2)
  record.rb                    # WIN32OLE::Record                  (Phase 3)
  variant.rb                   # WIN32OLE::Variant explicit wrapper (Phase 3)
  event.rb                     # WIN32OLE::Event                   (Phase 4)
```

This mirrors the existing C extension's own module boundaries
(`win32ole_type.c`, `win32ole_typelib.c`, `win32ole_method.c`, ...),
which should make cross-referencing behavior during implementation
straightforward.

### 4.3 `win32.rb`: the Fiddle substrate

Everything downstream depends on a small set of primitives, all
validated by the spike:

- `Fiddle.dlopen('ole32')` / `'oleaut32'` function bindings:
  `CoInitialize`, `CoUninitialize`, `CLSIDFromProgID`,
  `CoCreateInstance`, `SysAllocString`, `SysFreeString`, and (Phase 2)
  `LoadTypeLib`/`LoadRegTypeLib` for `ITypeLib`.
- `STDCALL = Fiddle::Function.const_defined?(:STDCALL) ? Fiddle::Function::STDCALL : Fiddle::Function::DEFAULT`
- `VARIANT_SIZE = Fiddle::SIZEOF_VOIDP == 8 ? 24 : 16` — the x86/x64
  VARIANT size gotcha from §1.2, made runtime-conditional instead of
  hardcoded to 24 (the spike hardcoded 24, correct only for x64).
- `variant(vt, value)` / VARIANT byte-packing and unpacking, using
  `Array#pack`/`String#unpack` (io-console style — no `FFI::Struct` or
  `Fiddle::Importer`-generated struct classes).
- `vtable_function(object_addr, index, arg_types, ret_type)` — reads a
  function pointer out of an arbitrary COM object's vtable at a given
  slot and wraps it as a callable `Fiddle::Function`. Used both for
  calling into IDispatch-based objects (Phase 1) and, later, into
  fixed-vtable custom interfaces like `ITypeInfo`/`ITypeLib` (Phase 2).
- `native_address_of(buffer)` — `Fiddle::Pointer.to_ptr(buffer).to_i`,
  used whenever a raw memory address must be embedded as data inside
  another packed buffer (e.g. a VARIANT's `bstrVal`, or
  `DISPPARAMS.rgvarg`). See the GC-safety risk in §7.2 — this pattern
  needs a keep-alive discipline analogous to the spike's `KEEP_ALIVE`
  array, made a first-class part of the design (§4.5), not an
  afterthought.

### 4.4 `dispatch.rb`: IDispatch helpers

A small mixin/module used by `WIN32OLE` (and later `WIN32OLE::Record`,
etc.) wrapping a raw IDispatch pointer:

- `dispid_for(name)` → `GetIDsOfNames`
- `invoke(dispid, args, wflags)` → builds `DISPPARAMS` (args reversed
  per COM convention — confirmed empirically not to matter for the
  *value* correctness in the spike's 2-arg tests, but kept as the
  spec-mandated order since real-world dual interfaces may rely on it
  for named/optional-argument resolution that the spike didn't
  exercise), calls `Invoke`, returns the raw result VARIANT bytes.
- `release` — calls the vtable's `Release`.

### 4.5 Object lifetime / GC safety

Every native address that gets embedded as *data* inside another
buffer (rather than passed as a direct Fiddle::Function argument, which
Fiddle can pin automatically) needs the Ruby object owning that memory
kept alive for as long as native code might dereference it. `WIN32OLE`
instances will hold their own keep-alive state (e.g. an instance
`@__native_buffers__` array) rather than relying on a single process-
wide global list as the spike did — the spike's `KEEP_ALIVE` global was
adequate for a short-lived script but is not adequate for a library
where objects should be independently collectible when the Ruby
program no longer references them (their COM `Release` should also
fire in that case, mirroring the C extension's finalizer behavior).

## 5. Phased roadmap

| Phase | Scope | Depends on |
|---|---|---|
| 1 | `WIN32OLE` core: construct via ProgID/CLSID, dynamic property get/put and method invocation, basic type marshaling (String/Integer/Float/Bool/nil/WIN32OLE), error translation | §6 |
| 2 | `WIN32OLE::TYPE`, `TYPELIB`, `METHOD`, `PARAM`, `VARIABLE` — type-library introspection via `ITypeInfo`/`ITypeLib` (custom vtable interfaces, not IDispatch) | Phase 1's `win32.rb` substrate |
| 3 | `WIN32OLE::Record`, `WIN32OLE::Variant` explicit wrapper | Phase 2 (Record needs `ITypeInfo` to resolve field layouts) |
| 4 | `WIN32OLE::Event` — connection-point-based event sinks via `Fiddle::Closure` | Phase 1; independently resolve the CI/environment question from §1.2 before or during implementation |

Only Phase 1 is designed in implementation-ready detail here (§6).
Phases 2–4 are scoped at the roadmap level; each should get its own
short design pass (or go straight to an implementation plan, at the
implementer's discretion) once Phase 1 has landed and the patterns it
establishes (file layout, keep-alive discipline, error translation)
are proven against the real test suite.

## 6. Phase 1 detailed design

### 6.1 Construction

```ruby
WIN32OLE.new(progid_or_clsid, host = nil)
```

- Resolve a ProgID to a CLSID via `CLSIDFromProgID` (already a plain
  16-byte GUID passthrough if given something that looks like a
  `{...}` CLSID string instead — mirror the C extension's existing
  `ole_initialize`/`create_win32ole` logic for what counts as a valid
  CLSID string).
- `CoCreateInstance(clsid, nil, CLSCTX_SERVER, IID_IDISPATCH, ppv)`.
  `host` (remote OLE) is out of scope for Phase 1 — raise
  `NotImplementedError` if given a non-nil host, rather than silently
  ignoring it.
- `WIN32OLE.connect` and `WIN32OLE.const_load` are out of scope for
  Phase 1.

### 6.2 Dynamic dispatch

`method_missing` intercepts unknown calls:

- **No arguments, method name doesn't end in `=`**: invoke with
  `wflags = DISPATCH_METHOD | DISPATCH_PROPERTYGET`. The spike found
  this combination necessary in practice for at least one real-world
  scripting-oriented dual interface (`Scripting.FileSystemObject`);
  using `DISPATCH_METHOD` alone risked spurious failures on objects
  that only answer property-style dispatch for some members.
- **Method name ends in `=`, exactly one argument**: invoke with
  `wflags = DISPATCH_PROPERTYPUT`, using the `DISPID_PROPERTYPUT`
  (`-3`) named-argument convention (`cNamedArgs = 1`,
  `rgdispidNamedArgs = [DISPID_PROPERTYPUT]`).
- **One or more arguments, otherwise**: `wflags = DISPATCH_METHOD`.

DISPIDs are resolved via `GetIDsOfNames` and intentionally *not*
cached across calls in Phase 1 (the C extension doesn't guarantee
per-instance caching survives object churn either; revisit as a
performance optimization once correctness is established).

### 6.3 Type marshaling (Ruby → VARIANT)

| Ruby | VARIANT |
|---|---|
| `String` | `VT_BSTR` (`SysAllocString` on a UTF-16LE-encoded, NUL-terminated copy) |
| `Integer` (fits in 32 bits) | `VT_I4` |
| `Integer` (needs 64 bits) | `VT_I8` |
| `Float` | `VT_R8` |
| `true` / `false` | `VT_BOOL` (`-1` / `0`) |
| `nil` | `VT_EMPTY` |
| `WIN32OLE` instance | `VT_DISPATCH` (the wrapped IDispatch pointer) |

Arrays, Records, and explicit `WIN32OLE::Variant` wrapping (to force a
specific VARTYPE) are out of scope for Phase 1 — raise
`TypeError`/`NotImplementedError` for argument types not in this
table, rather than silently coercing incorrectly.

### 6.4 Type marshaling (VARIANT → Ruby), return values

Mirror of §6.3 in reverse, plus:

- `VT_DISPATCH` / `VT_UNKNOWN` results wrap into a new `WIN32OLE`
  instance around the returned pointer.
- Any VARTYPE not covered raises `NotImplementedError` with the
  numeric VARTYPE in the message, rather than returning `nil` or
  garbage — Phase 2/3 will extend this table (arrays, records,
  currency, date, decimal) and each addition should have a
  corresponding test before the `NotImplementedError` is removed for
  that type.

### 6.5 Error translation

An `Invoke` call returning a non-zero HRESULT must raise
`WIN32OLERuntimeError` with a message matching the existing C
extension's format closely enough that tests written against MRI's
error messages keep passing where they assert on message content.
This requires reading `EXCEPINFO` (when `DISP_E_EXCEPTION` is
returned) and replicating `win32ole_error.c`'s formatting — an
implementation-time task, not something this design finalizes, since
it's a direct 1:1 port rather than a design decision.

## 7. Testing / CI strategy

- Add a `jruby` entry to `.github/workflows/windows.yml`'s engine
  matrix (currently `engine: cruby` only), running the existing
  `test/win32ole/*.rb` suite unmodified against the new backend.
- Tests exercising Phase 2+ functionality will fail under JRuby until
  those phases land. Rather than skipping whole files, guard
  individual tests/assertions the way
  `test/win32ole/test_win32ole_event.rb` already guards on
  `swbemsink_available` — add an analogous
  `defined?(WIN32OLE::TYPE) && ...`-style guard, or (simpler, and
  preferred if it doesn't fight the test framework) an explicit
  `skip` at the top of tests for not-yet-implemented classes, with a
  comment pointing at the phase that will remove the skip.
- Add a `GC.stress`-enabled test run (at least for Phase 1's core
  invoke path) to probe the keep-alive discipline from §4.5 — this
  class of bug (a buffer collected between being built and being read
  by native code) is exactly the kind that passes every normal test
  run and then crashes intermittently in production.

## 8. Risks / open questions carried forward

1. **`WIN32OLE::Event` live-fire delivery unresolved in CI** (§1.2,
   §5 Phase 4). Needs either a different event source not dependent on
   WMI's async delivery path, or verification on a non-CI Windows
   desktop environment, before Phase 4 can be considered validated.
2. **GC/compaction safety of the `native_address_of` pattern** (§4.5)
   needs explicit `GC.stress` testing (§7) before this is trusted for
   anything beyond a spike.
3. **x86 (32-bit) Windows is unverified.** The VARIANT-size branch
   (§4.3) is written to be correct on paper for both, but only the x64
   path has actually run.
4. **Performance** is expected to be worse than the C extension
   (Fiddle/FFI call overhead vs. direct C calls). Treated as an
   accepted trade-off for portability, not something Phase 1 needs to
   optimize away.
5. **Encoding/locale fidelity** (CP_ACP handling, etc.) against the C
   extension's exact behavior needs implementation-time verification
   against the existing test suite rather than being fully specified
   here.
