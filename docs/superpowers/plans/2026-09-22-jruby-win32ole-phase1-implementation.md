# JRuby win32ole Phase 1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `require 'win32ole'; WIN32OLE.new(...)` work on JRuby (Phase 1 scope only: construction, dynamic dispatch, basic type marshaling, error translation), implemented in pure Ruby + Fiddle, without touching the existing MRI C extension.

**Architecture:** A new `lib/win32ole/jruby/` tree provides a from-scratch `WIN32OLE` class built on Fiddle bindings to `ole32.dll`/`oleaut32.dll`, driving raw `IDispatch` vtables (`GetIDsOfNames`/`Invoke`/`Release`) exactly the way the approved spike (`tmp_spike/fiddle_com_spike.rb`) already proved works identically on MRI and JRuby. `lib/win32ole.rb` gains a `RUBY_ENGINE == 'jruby'` branch that loads this tree instead of the compiled `.so`.

**Tech Stack:** Ruby stdlib `fiddle` only (no direct `ffi` dependency — Fiddle's own JRuby backend pulls it in transitively). Windows-only; runs in GitHub Actions `windows-latest`.

**Spec:** `docs/superpowers/specs/2026-09-22-jruby-win32ole-support-design.md` — this plan implements **§6 (Phase 1 detailed design)** plus the Phase-1-relevant parts of §4 (architecture) and §7 (testing/CI). Phases 2–4 (§5) are explicitly out of scope; every task below raises `NotImplementedError`/`TypeError` rather than guessing when it hits Phase 2+ territory, per §6.3/§6.4.

## Global Constraints

These apply to every task below; not repeated per-task.

- **MRI path stays untouched.** Never edit `ext/win32ole/*.c`. The only shared file touched is `lib/win32ole.rb`, and only to add the engine branch.
- **No direct `ffi` gem dependency.** Only `require 'fiddle'`. (Design §1.1, §4.3.)
- **`STDCALL` fallback:** `Fiddle::Function.const_defined?(:STDCALL) ? Fiddle::Function::STDCALL : Fiddle::Function::DEFAULT` — `STDCALL` is undefined on x64 mingw MRI. (Design §4.3, §1.2.)
- **`VARIANT_SIZE` is runtime-conditional:** `Fiddle::SIZEOF_VOIDP == 8 ? 24 : 16`, never hardcoded. (Design §4.3, §1.2.)
- **No `FFI::Struct` / `Fiddle::Importer` struct classes.** Byte-pack/unpack with `Array#pack`/`String#unpack`, io-console style. (Design §1.1, §4.3.)
- **Every native address embedded as *data* inside another buffer** (a VARIANT's pointer slot, a DISPPARAMS array element) **must be kept alive** via the owning `WIN32OLE` instance's own state, not a process-global list. (Design §4.5.)
- **All `Fiddle.dlopen` / `Fiddle::Function.new` calls for OS entry points must be lazy (memoized methods), never top-level constants.** This is an implementation-time decision this plan adds on top of the design: it's what makes `lib/win32ole/jruby/win32.rb`'s pure byte-marshaling logic (Task 2) locally unit-testable on any OS — requiring the file must never itself call `Fiddle.dlopen('ole32')`, which would raise on non-Windows. Real Windows API calls only happen when a binding method is actually invoked.
- **Argument/return types outside the Phase 1 tables (§6.3/§6.4) raise, never silently coerce:** `TypeError` for an unsupported Ruby argument type, `NotImplementedError` (numeric VARTYPE in the message) for an unsupported return VARTYPE.
- **Verification reality check:** this gem's Windows COM calls cannot run on non-Windows hardware. This repo's own `Rakefile` only wires `rake test` to the native `:compile` task on `mswin|mingw|cygwin` hosts — on every other OS the existing suite silently no-ops (every test file does `begin; require 'win32ole'; rescue LoadError; end` then guards on `if defined?(WIN32OLE)`). Tasks below are split so that **pure logic (no OS calls) gets a real local Red-Green test cycle**, and **native-call tasks are verified by pushing to the branch and reading the GitHub Actions `windows` workflow run** — call this out explicitly in those steps instead of pretending a local `Expected: PASS` is possible.

---

## Task 1: Engine dispatch wiring + file skeleton

**Files:**
- Modify: `lib/win32ole.rb` (all 33 lines)
- Create: `lib/win32ole/jruby.rb`
- Create: `lib/win32ole/jruby/win32.rb` (empty shell for now — filled in Task 2/3)
- Create: `lib/win32ole/jruby/dispatch.rb` (empty shell — filled in Task 4)
- Create: `lib/win32ole/jruby/win32ole.rb` (empty shell — filled in Task 5/6)
- Test: `test/win32ole/jruby/test_jruby_require.rb`

**Interfaces:**
- Produces: `require 'win32ole/jruby'` — the single entry point the branch in `lib/win32ole.rb` calls. After this task, `require 'win32ole'` on JRuby defines an empty `class WIN32OLE; end` (no methods yet — those land in later tasks) without raising.

Today's `lib/win32ole.rb` wraps the native require in `begin/rescue LoadError` so that on a platform without the compiled extension, `require 'win32ole'` silently no-ops instead of raising, and then conditionally reopens `WIN32OLE#methods` (for `did_you_mean` support) *only if* `WIN32OLE` ended up defined. The design doc's §4.1 sketch (`if RUBY_ENGINE == 'jruby' ... else require 'win32ole.so' end`) is a simplification that would change MRI's LoadError-swallowing behavior — keep that behavior; only add the JRuby branch alongside it. Also: the existing `methods` monkeypatch calls `ole_methods`, which won't exist until Phase 2 — its rescue clause must be widened so Phase 1 JRuby doesn't blow up on `some_win32ole_instance.methods`.

- [ ] **Step 1: Write the failing test**

```ruby
# test/win32ole/jruby/test_jruby_require.rb
require 'test/unit'

class TestJRubyRequire < Test::Unit::TestCase
  def test_require_win32ole_defines_the_class_on_jruby
    omit('JRuby-only') unless RUBY_ENGINE == 'jruby'

    require 'win32ole'
    assert(defined?(WIN32OLE), 'WIN32OLE should be defined after require "win32ole" on JRuby')
    assert_kind_of(Class, WIN32OLE)
  end

  def test_methods_override_does_not_raise_before_ole_methods_exists
    omit('JRuby-only') unless RUBY_ENGINE == 'jruby'
    require 'win32ole'

    obj = WIN32OLE.allocate
    assert_nothing_raised { obj.methods }
  end
end
```

- [ ] **Step 2: Run test to verify it fails (or omits) on this machine**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_jruby_require.rb`
Expected on this (non-JRuby) machine: both tests **OMIT** (`RUBY_ENGINE` is `ruby`, not `jruby`) — that's a pass-with-omission, not a failure. This confirms the test file loads and the guard works; the real red/green cycle for this test only happens under JRuby in CI (Task 7).

- [ ] **Step 3: Create the four skeleton files**

```ruby
# lib/win32ole/jruby/win32.rb
class WIN32OLE
  module Win32
  end
end
```

```ruby
# lib/win32ole/jruby/dispatch.rb
class WIN32OLE
  module Dispatch
  end
end
```

```ruby
# lib/win32ole/jruby/win32ole.rb
require 'fiddle'
require 'win32ole/jruby/win32'
require 'win32ole/jruby/dispatch'

class WIN32OLE
  include Dispatch
end
```

```ruby
# lib/win32ole/jruby.rb
require 'win32ole/jruby/win32ole'
```

- [ ] **Step 4: Update `lib/win32ole.rb`**

```ruby
begin
  if RUBY_ENGINE == 'jruby'
    require 'win32ole/jruby'
  else
    require 'win32ole.so'
  end
rescue LoadError
  # do nothing
end

if defined?(WIN32OLE)
  class WIN32OLE

    #
    # By overriding Object#methods, WIN32OLE might
    # work well with did_you_mean gem.
    # This is experimental.
    #
    #  require 'win32ole'
    #  dict = WIN32OLE.new('Scripting.Dictionary')
    #  dict.Ade('a', 1)
    #  #=> Did you mean?  Add
    #
    def methods(*args)
      super + ole_methods_safely.map(&:name).map(&:to_sym)
    end

    private

    def ole_methods_safely
      ole_methods
    rescue WIN32OLE::QueryInterfaceError, NoMethodError
      []
    end
  end
end
```

The only changes from the current file: the `require` line is now engine-conditional, and the rescue clause on line `rescue WIN32OLE::QueryInterfaceError` gained `, NoMethodError` — needed because Phase 1's `WIN32OLE` has no `ole_methods` at all (that's Phase 2), so calling `.methods` on any Phase-1 JRuby `WIN32OLE` instance must not blow up.

- [ ] **Step 5: Run test to verify it passes (JRuby) / still omits (MRI)**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_jruby_require.rb`
Expected on this machine: still 2 omissions, 0 failures (MRI's own `begin require 'win32ole.so' rescue LoadError` path is unchanged — verify no regression by also running `ruby -Ilib -Itest test/win32ole/test_win32ole.rb` and confirming it behaves exactly as before your edit, i.e. `defined?(WIN32OLE)` is false and every test in that file is skipped the same way it was pre-change).
Expected under JRuby (checked in Task 7's CI run): both tests PASS.

- [ ] **Step 6: Commit**

```bash
git add lib/win32ole.rb lib/win32ole/jruby.rb lib/win32ole/jruby/win32.rb lib/win32ole/jruby/dispatch.rb lib/win32ole/jruby/win32ole.rb test/win32ole/jruby/test_jruby_require.rb
git commit -m "jruby: wire up engine dispatch and skeleton files"
```

---

## Task 2: `win32.rb` — pure marshaling & decision logic (local TDD)

**Files:**
- Modify: `lib/win32ole/jruby/win32.rb`
- Test: `test/win32ole/jruby/test_win32.rb`

**Interfaces:**
- Consumes: nothing (no dependency on Task 1's other files).
- Produces (all under `WIN32OLE::Win32`, all `module_function`, all pure — no OS calls):
  - Constants: `STDCALL`, `VOIDP`, `LONG`, `DWORD`, `WORD`, `VOID`, `VARIANT_SIZE`, `PTR_SIZE`, `VT_EMPTY`/`VT_I4`/`VT_I8`/`VT_R8`/`VT_BOOL`/`VT_BSTR`/`VT_DISPATCH`/`VT_UNKNOWN`, `DISPATCH_METHOD`/`DISPATCH_PROPERTYGET`/`DISPATCH_PROPERTYPUT`, `DISPID_PROPERTYPUT`, `CLSCTX_INPROC_SERVER`/`CLSCTX_LOCAL_SERVER`, `IID_NULL`, `IID_IDISPATCH`, `EXCEPINFO_SIZE`, `EXCEPINFO_OFFSETS`.
  - `wstr(str) -> String` (UTF-16LE, NUL-terminated, binary-tagged)
  - `pack_variant(vt, payload8) -> String` / `unpack_variant(bytes) -> [vt, payload8]`
  - `pack_i4/i8/r8/bool/pointer(value) -> String(8 bytes)` and matching `unpack_i4/i8/r8/bool/pointer(payload8) -> value`
  - `pack_empty -> String(8 zero bytes)`
  - `ruby_to_variant_type(value) -> Symbol` (`:bstr`/`:i4`/`:i8`/`:r8`/`:bool`/`:empty`/`:dispatch`, raises `TypeError` otherwise)
  - `variant_ruby_type(vt) -> Symbol` (raises `NotImplementedError` with the numeric vt in the message otherwise)
  - `VT_FOR_TYPE` hash: `{i4:, i8:, r8:, bool:, empty:, bstr:, dispatch:}` → VT constant
  - `dispatch_plan(name, args) -> {name:, wflags:, named_put:}` (the §6.2 decision table)
  - `failed?(hr) -> bool`, `hr_hex(hr) -> String` (e.g. `"0x800401f3"`)
  - `method_error_message(method_name, detail) -> String`, `property_put_error_message(property_name, detail) -> String`, `unknown_server_error_message(server_name) -> String`
  - `parse_excepinfo(bytes) -> {w_code:, scode:, bstr_source_ptr:, bstr_description_ptr:}`

This is deliberately the file's *only* content for this task — no `Fiddle.dlopen` anywhere yet, so every test below runs on this machine right now with plain `ruby`, no Windows required.

- [ ] **Step 1: Write the failing tests**

```ruby
# test/win32ole/jruby/test_win32.rb
require 'test/unit'
require 'win32ole/jruby/win32'

class TestWin32 < Test::Unit::TestCase
  W = WIN32OLE::Win32

  def test_stdcall_falls_back_to_default_when_undefined
    expected = Fiddle::Function.const_defined?(:STDCALL) ? Fiddle::Function::STDCALL : Fiddle::Function::DEFAULT
    assert_equal(expected, W::STDCALL)
  end

  def test_variant_size_is_24_on_64bit_16_on_32bit
    expected = Fiddle::SIZEOF_VOIDP == 8 ? 24 : 16
    assert_equal(expected, W::VARIANT_SIZE)
  end

  def test_wstr_is_utf16le_nul_terminated_binary
    bytes = W.wstr('AB')
    assert_equal(Encoding::ASCII_8BIT, bytes.encoding)
    assert_equal("A\x00B\x00\x00\x00".b, bytes)
  end

  def test_pack_variant_round_trips_with_i4_payload
    packed = W.pack_variant(W::VT_I4, W.pack_i4(42))
    assert_equal(W::VARIANT_SIZE, packed.bytesize)
    vt, payload = W.unpack_variant(packed)
    assert_equal(W::VT_I4, vt)
    assert_equal(42, W.unpack_i4(payload))
  end

  def test_pack_variant_rejects_wrong_size_payload
    assert_raise(ArgumentError) { W.pack_variant(W::VT_I4, "\x00\x00\x00") }
  end

  def test_i8_round_trip
    big = 5_000_000_000
    assert_equal(big, W.unpack_i8(W.pack_i8(big)))
  end

  def test_r8_round_trip
    assert_in_delta(3.5, W.unpack_r8(W.pack_r8(3.5)), 0.0001)
  end

  def test_bool_round_trip
    assert_equal(true, W.unpack_bool(W.pack_bool(true)))
    assert_equal(false, W.unpack_bool(W.pack_bool(false)))
  end

  def test_pointer_round_trip
    addr = 0x00007ff6_12345678
    assert_equal(addr, W.unpack_pointer(W.pack_pointer(addr)))
  end

  def test_iid_idispatch_matches_known_bytes
    expected = [0x00020400, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46].pack('LSSC8')
    assert_equal(expected, W::IID_IDISPATCH)
  end

  def test_ruby_to_variant_type_mapping
    assert_equal(:bstr, W.ruby_to_variant_type('x'))
    assert_equal(:i4, W.ruby_to_variant_type(42))
    assert_equal(:i8, W.ruby_to_variant_type(5_000_000_000))
    assert_equal(:i4, W.ruby_to_variant_type(-(2**31)))
    assert_equal(:i8, W.ruby_to_variant_type(-(2**31) - 1))
    assert_equal(:r8, W.ruby_to_variant_type(1.5))
    assert_equal(:bool, W.ruby_to_variant_type(true))
    assert_equal(:bool, W.ruby_to_variant_type(false))
    assert_equal(:empty, W.ruby_to_variant_type(nil))
  end

  def test_ruby_to_variant_type_rejects_unsupported
    assert_raise(TypeError) { W.ruby_to_variant_type([1, 2]) }
    assert_raise(TypeError) { W.ruby_to_variant_type({}) }
  end

  def test_variant_ruby_type_mapping
    assert_equal(:i4, W.variant_ruby_type(W::VT_I4))
    assert_equal(:bstr, W.variant_ruby_type(W::VT_BSTR))
    assert_equal(:dispatch, W.variant_ruby_type(W::VT_DISPATCH))
    assert_equal(:dispatch, W.variant_ruby_type(W::VT_UNKNOWN))
  end

  def test_variant_ruby_type_rejects_unsupported_vartype
    err = assert_raise(NotImplementedError) { W.variant_ruby_type(99) }
    assert_match(/99/, err.message)
  end

  def test_dispatch_plan_zero_args_uses_method_or_propertyget
    plan = W.dispatch_plan('Count', [])
    assert_equal('Count', plan[:name])
    assert_equal(W::DISPATCH_METHOD | W::DISPATCH_PROPERTYGET, plan[:wflags])
    assert_equal(false, plan[:named_put])
  end

  def test_dispatch_plan_setter_uses_propertyput
    plan = W.dispatch_plan('compareMode=', [1])
    assert_equal('compareMode', plan[:name])
    assert_equal(W::DISPATCH_PROPERTYPUT, plan[:wflags])
    assert_equal(true, plan[:named_put])
  end

  def test_dispatch_plan_setter_requires_exactly_one_arg
    assert_raise(ArgumentError) { W.dispatch_plan('compareMode=', [1, 2]) }
  end

  def test_dispatch_plan_with_args_uses_method
    plan = W.dispatch_plan('Add', ['k', 'v'])
    assert_equal(W::DISPATCH_METHOD, plan[:wflags])
    assert_equal(false, plan[:named_put])
  end

  def test_failed_predicate
    assert_equal(false, W.failed?(0))
    assert_equal(true, W.failed?(-2147024809)) # E_INVALIDARG
  end

  def test_hr_hex_formats_unsigned
    assert_equal('0x800401f3', W.hr_hex(-2147221005))
  end

  def test_method_error_message_matches_existing_c_ext_format
    msg = W.method_error_message('add', 'boom')
    assert_match(/\A\(in OLE method `add': \)boom\z/, msg) #`
  end

  def test_property_put_error_message_matches_existing_c_ext_format
    msg = W.property_put_error_message('compareMode', 'boom')
    assert_match(/\A\(in setting property `compareMode': \)boom\z/, msg) #`
  end

  def test_unknown_server_error_message_matches_existing_c_ext_format
    assert_equal("unknown OLE server: `NonExistProgID'", W.unknown_server_error_message('NonExistProgID')) #`
  end

  def test_parse_excepinfo_reads_known_offsets
    o = W::EXCEPINFO_OFFSETS
    buf = ("\x00" * W::EXCEPINFO_SIZE).b
    buf[o[:wCode], 2] = [7].pack('S')
    buf[o[:bstrDescription], W::PTR_SIZE] = [0x1234].pack(W::PTR_SIZE == 8 ? 'Q' : 'L')
    buf[o[:scode], 4] = [-2147024809].pack('l')

    info = W.parse_excepinfo(buf)
    assert_equal(7, info[:w_code])
    assert_equal(0x1234, info[:bstr_description_ptr])
    assert_equal(-2147024809, info[:scode])
  end
end
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_win32.rb`
Expected: every test FAILS or errors with `NoMethodError`/`NameError` (`WIN32OLE::Win32` is still an empty module).

- [ ] **Step 3: Implement `win32.rb`**

```ruby
# lib/win32ole/jruby/win32.rb
require 'fiddle'

class WIN32OLE
  module Win32
    STDCALL = if Fiddle::Function.const_defined?(:STDCALL)
                Fiddle::Function::STDCALL
              else
                Fiddle::Function::DEFAULT
              end

    VOIDP = Fiddle::TYPE_VOIDP
    LONG  = Fiddle::TYPE_LONG
    DWORD = -Fiddle::TYPE_INT
    WORD  = -Fiddle::TYPE_SHORT
    VOID  = Fiddle::TYPE_VOID

    PTR_SIZE     = Fiddle::SIZEOF_VOIDP
    PACK_PTR     = PTR_SIZE == 8 ? 'Q' : 'L' # native pointer width, for packing real structs (DISPPARAMS, pointer arrays) — NOT for a VARIANT's 8-byte value slot, which always uses 'Q' regardless of platform (see pack_pointer)
    VARIANT_SIZE = PTR_SIZE == 8 ? 24 : 16

    VT_EMPTY    = 0
    VT_I4       = 3
    VT_R8       = 5
    VT_BSTR     = 8
    VT_DISPATCH = 9
    VT_BOOL     = 11
    VT_UNKNOWN  = 13
    VT_I8       = 20

    DISPATCH_METHOD      = 1
    DISPATCH_PROPERTYGET = 2
    DISPATCH_PROPERTYPUT = 4
    DISPID_PROPERTYPUT   = -3

    CLSCTX_INPROC_SERVER = 0x1
    CLSCTX_LOCAL_SERVER  = 0x4

    IID_NULL      = ("\x00" * 16).b
    IID_IDISPATCH = [0x00020400, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46].pack('LSSC8')

    VT_FOR_TYPE = {
      i4: VT_I4, i8: VT_I8, r8: VT_R8, bool: VT_BOOL,
      empty: VT_EMPTY, bstr: VT_BSTR, dispatch: VT_DISPATCH
    }.freeze

    INT32_RANGE = (-(2**31))..(2**31 - 1)

    # EXCEPINFO (oaidl.h): WORD wCode; WORD wReserved; BSTR bstrSource;
    # BSTR bstrDescription; BSTR bstrHelpFile; DWORD dwHelpContext;
    # PVOID pvReserved; HRESULT(*pfnDeferredFillIn)(...); SCODE scode;
    if PTR_SIZE == 8
      EXCEPINFO_SIZE = 64
      EXCEPINFO_OFFSETS = {
        wCode: 0, bstrSource: 8, bstrDescription: 16, bstrHelpFile: 24,
        dwHelpContext: 32, pvReserved: 40, pfnDeferredFillIn: 48, scode: 56
      }.freeze
    else
      EXCEPINFO_SIZE = 32
      EXCEPINFO_OFFSETS = {
        wCode: 0, bstrSource: 4, bstrDescription: 8, bstrHelpFile: 12,
        dwHelpContext: 16, pvReserved: 20, pfnDeferredFillIn: 24, scode: 28
      }.freeze
    end

    module_function

    def wstr(str)
      "#{str}\x00".encode('UTF-16LE').b
    end

    def pack_variant(vt, payload)
      payload = payload.b
      unless payload.bytesize == 8
        raise ArgumentError, "payload must be 8 bytes, got #{payload.bytesize}"
      end

      [vt, 0, 0, 0].pack('S4') + payload + ("\x00".b * (VARIANT_SIZE - 16))
    end

    def unpack_variant(bytes)
      vt, = bytes.unpack1('S')
      [vt, bytes[8, 8]]
    end

    def pack_i4(value)    = [value].pack('l') + ("\x00".b * 4)
    def pack_i8(value)    = [value].pack('q')
    def pack_r8(value)    = [value].pack('d')
    def pack_bool(value)  = [value ? -1 : 0].pack('s') + ("\x00".b * 6)
    def pack_pointer(addr) = [addr].pack('Q')
    def pack_empty        = "\x00".b * 8

    def unpack_i4(payload)    = payload.unpack1('l')
    def unpack_i8(payload)    = payload.unpack1('q')
    def unpack_r8(payload)    = payload.unpack1('d')
    def unpack_bool(payload)  = payload.unpack1('s') != 0
    def unpack_pointer(payload) = payload.unpack1('Q')

    def ruby_to_variant_type(value)
      case value
      when String then :bstr
      when Integer then INT32_RANGE.cover?(value) ? :i4 : :i8
      when Float then :r8
      when true, false then :bool
      when nil then :empty
      when ::WIN32OLE then :dispatch
      else
        raise TypeError, "unsupported argument type for OLE call: #{value.class}"
      end
    end

    def variant_ruby_type(vt)
      case vt
      when VT_EMPTY then :empty
      when VT_I4 then :i4
      when VT_I8 then :i8
      when VT_R8 then :r8
      when VT_BOOL then :bool
      when VT_BSTR then :bstr
      when VT_DISPATCH, VT_UNKNOWN then :dispatch
      else
        raise NotImplementedError, "VARTYPE #{vt} is not supported yet"
      end
    end

    def dispatch_plan(name, args)
      if name.end_with?('=')
        unless args.size == 1
          raise ArgumentError, "property put takes exactly one argument, got #{args.size}"
        end

        { name: name[0..-2], wflags: DISPATCH_PROPERTYPUT, named_put: true }
      elsif args.empty?
        { name: name, wflags: DISPATCH_METHOD | DISPATCH_PROPERTYGET, named_put: false }
      else
        { name: name, wflags: DISPATCH_METHOD, named_put: false }
      end
    end

    def failed?(hr)
      hr.negative?
    end

    def hr_hex(hr)
      format('0x%08x', hr & 0xFFFFFFFF)
    end

    def method_error_message(method_name, detail)
      "(in OLE method `#{method_name}': )#{detail}"
    end

    def property_put_error_message(property_name, detail)
      "(in setting property `#{property_name}': )#{detail}"
    end

    def unknown_server_error_message(server_name)
      "unknown OLE server: `#{server_name}'"
    end

    def parse_excepinfo(bytes)
      o = EXCEPINFO_OFFSETS
      ptr_fmt = PTR_SIZE == 8 ? 'Q' : 'L'
      {
        w_code: bytes[o[:wCode], 2].unpack1('S'),
        bstr_source_ptr: bytes[o[:bstrSource], PTR_SIZE].unpack1(ptr_fmt),
        bstr_description_ptr: bytes[o[:bstrDescription], PTR_SIZE].unpack1(ptr_fmt),
        scode: bytes[o[:scode], 4].unpack1('l')
      }
    end
  end
end
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `ruby -Ilib -Itest test/win32ole/jruby/test_win32.rb`
Expected: all tests PASS. This runs for real right now, on this machine, with system Ruby — no Windows, no JRuby needed, because nothing here calls `Fiddle.dlopen`.

- [ ] **Step 5: Commit**

```bash
git add lib/win32ole/jruby/win32.rb test/win32ole/jruby/test_win32.rb
git commit -m "jruby: pure VARIANT/dispatch-decision/error-message logic for win32.rb"
```

---

## Task 3: `win32.rb` — native Win32 API substrate

**Files:**
- Modify: `lib/win32ole/jruby/win32.rb`

**Interfaces:**
- Consumes: `STDCALL`, `VOIDP`/`LONG`/`DWORD`/`WORD`/`VOID`, `IID_NULL`/`IID_IDISPATCH` from Task 2.
- Produces (all lazy/memoized `module_function` methods, per Global Constraints — none of these run at require time):
  - `co_initialize`, `co_uninitialize`, `clsid_from_progid`, `clsid_from_string`, `co_create_instance`, `sys_alloc_string`, `sys_free_string`, `format_message` — each a memoized `Fiddle::Function`.
  - `native_address_of(buffer) -> Integer`
  - `vtable_function(object_addr, index, arg_types, ret_type) -> Fiddle::Function`
  - `bstr_to_s(addr) -> String` (reads a native BSTR into a Ruby UTF-8 string; `nil` for a null/zero address)
  - `hresult_system_message(hr) -> String` (calls `FormatMessage`; returns `""` if the OS has no message for that code — mirrors `ole_hresult2msg`'s "IGNORE_INSERTS" text-only behavior from `ext/win32ole/win32ole_error.c:5-43`, without the bilingual English/`cWIN32OLE_lcid` retry, which is out of scope for Phase 1 per the design's §8 item 5 encoding caveat)

This is the point where the file gains real Windows dependencies. Nothing here is runnable on this machine — verification is CI-only, via Task 7's workflow, but written now because Task 4/5/6 need it.

- [ ] **Step 1: Implement the native substrate, appended to `lib/win32ole/jruby/win32.rb`'s `module Win32` block**

Insert after the `module_function` line's existing pure methods (i.e. these become additional `module_function` methods in the same module):

```ruby
    def ole32
      @ole32 ||= Fiddle.dlopen('ole32')
    end

    def oleaut32
      @oleaut32 ||= Fiddle.dlopen('oleaut32')
    end

    def kernel32
      @kernel32 ||= Fiddle.dlopen('kernel32')
    end

    def co_initialize
      @co_initialize ||= Fiddle::Function.new(ole32['CoInitialize'], [VOIDP], LONG, STDCALL)
    end

    def co_uninitialize
      @co_uninitialize ||= Fiddle::Function.new(ole32['CoUninitialize'], [], VOID, STDCALL)
    end

    def clsid_from_progid
      @clsid_from_progid ||= Fiddle::Function.new(ole32['CLSIDFromProgID'], [VOIDP, VOIDP], LONG, STDCALL)
    end

    def clsid_from_string
      @clsid_from_string ||= Fiddle::Function.new(ole32['CLSIDFromString'], [VOIDP, VOIDP], LONG, STDCALL)
    end

    def co_create_instance
      @co_create_instance ||= Fiddle::Function.new(
        ole32['CoCreateInstance'], [VOIDP, VOIDP, DWORD, VOIDP, VOIDP], LONG, STDCALL
      )
    end

    def sys_alloc_string
      @sys_alloc_string ||= Fiddle::Function.new(oleaut32['SysAllocString'], [VOIDP], VOIDP, STDCALL)
    end

    def sys_free_string
      @sys_free_string ||= Fiddle::Function.new(oleaut32['SysFreeString'], [VOIDP], VOID, STDCALL)
    end

    FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100
    FORMAT_MESSAGE_FROM_SYSTEM     = 0x00001000
    FORMAT_MESSAGE_IGNORE_INSERTS  = 0x00000200

    def format_message
      @format_message ||= Fiddle::Function.new(
        kernel32['FormatMessageW'], [DWORD, VOIDP, DWORD, DWORD, VOIDP, DWORD, VOIDP], DWORD, STDCALL
      )
    end

    def native_address_of(buffer)
      Fiddle::Pointer.to_ptr(buffer).to_i
    end

    def vtable_function(object_addr, index, arg_types, ret_type)
      vtable_addr = Fiddle::Pointer.new(object_addr)[0, PTR_SIZE].unpack1(PTR_SIZE == 8 ? 'Q' : 'L')
      func_addr = Fiddle::Pointer.new(vtable_addr)[index * PTR_SIZE, PTR_SIZE].unpack1(PTR_SIZE == 8 ? 'Q' : 'L')
      Fiddle::Function.new(func_addr, arg_types, ret_type, STDCALL)
    end

    def bstr_to_s(addr)
      return nil if addr.nil? || addr.zero?

      ptr = Fiddle::Pointer.new(addr)
      units = []
      offset = 0
      loop do
        unit = ptr[offset, 2].unpack1('S')
        break if unit.zero?

        units << unit
        offset += 2
      end
      units.pack('U*')
    end

    def hresult_system_message(hr)
      buf_ptr = ("\x00" * PTR_SIZE).b
      flags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS
      count = format_message.call(flags, nil, hr, 0, buf_ptr, 0, nil)
      return '' if count.zero?

      addr = buf_ptr.unpack1(PTR_SIZE == 8 ? 'Q' : 'L')
      msg = bstr_free_local_string(addr, count)
      msg.chomp
    end

    # count is the number of UTF-16LE *characters* FormatMessageW wrote,
    # not bytes — this exact kind of factor-of-2 slip is what the spike's
    # own retrospective (design §1.2) warns about: easy to get subtly
    # wrong, and it only shows up once you actually run it on Windows,
    # which is why this task's real verification is the CI run in Task 7,
    # not this write-up.
    def bstr_free_local_string(addr, count)
      ptr = Fiddle::Pointer.new(addr)
      msg = ptr[0, count * 2].dup.force_encoding('UTF-16LE').encode('UTF-8')
      local_free.call(addr)
      msg
    end

    def local_free
      @local_free ||= Fiddle::Function.new(kernel32['LocalFree'], [VOIDP], VOIDP, STDCALL)
    end
```

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/win32.rb
git commit -m "jruby: native Win32/COM API bindings for win32.rb"
```

No local run here — see Task 7 for the first point this can actually execute. Re-run `ruby -Ilib -Itest test/win32ole/jruby/test_win32.rb` now as a smoke check that nothing above broke the pure logic from Task 2 (it must still pass unchanged, proving the file is still safely requireable without touching Windows).

---

## Task 4: `dispatch.rb` — IDispatch mixin

**Files:**
- Modify: `lib/win32ole/jruby/dispatch.rb`

**Interfaces:**
- Consumes: `WIN32OLE::Win32` (all of it — Tasks 2 and 3).
- Produces (instance methods mixed into `WIN32OLE`, assuming the including class exposes a `@ptr` integer ivar holding the raw `IDispatch*`):
  - `dispid_for(name) -> Integer` (raises `WIN32OLE::RuntimeError` on failure, formatted per `method_error_message`/caller context — see below)
  - `ole_invoke(dispid, arg_values, wflags, named_put: false) -> [hr, result_variant_bytes, excepinfo_bytes]` — does *not* raise; returns raw data so `win32ole.rb` (Task 6) can build the exact "(in OLE method ...)" vs "(in setting property ...)" message depending on call site.
  - `release` — calls the vtable's `Release`.
  - Private: `get_ids_of_names_fn`, `invoke_fn`, `release_fn` (memoized per-instance `Fiddle::Function` wrappers over this object's own vtable, via `Win32.vtable_function`).

This mirrors `tmp_spike/fiddle_com_spike.rb`'s already-validated `dispid_for`/`invoke_bstr_method`, generalized from "BSTR args, BSTR/nil result" to the Phase 1 type table (§6.3/§6.4) and rewired to keep native buffers alive on `self` instead of a spike-only global.

- [ ] **Step 1: Implement `dispatch.rb`**

```ruby
# lib/win32ole/jruby/dispatch.rb
require 'win32ole/jruby/win32'

class WIN32OLE
  module Dispatch
    W = WIN32OLE::Win32

    def dispid_for(name)
      name_buf = W.wstr(name)
      names = [W.native_address_of(name_buf)].pack(W::PACK_PTR)
      dispids = ("\x00" * 4).b
      hr = get_ids_of_names_fn.call(@ptr, W::IID_NULL, names, 1, 0, dispids)
      return nil if W.failed?(hr)

      dispids.unpack1('l')
    end

    # Returns [hr, result_variant_bytes, excepinfo_bytes]. Never raises —
    # callers build the user-facing error message (they know whether this
    # was a method call or a property-put, which changes the message).
    def ole_invoke(dispid, arg_values, wflags, named_put: false)
      keep_alive = native_buffers

      arg_variants = arg_values.reverse.map { |v| ruby_value_to_variant_bytes(v, keep_alive) }
      args_blob = arg_variants.join
      keep_alive << args_blob unless args_blob.empty?

      named_blob = named_put ? [W::DISPID_PROPERTYPUT].pack('l') : ''
      keep_alive << named_blob unless named_blob.empty?

      # DISPPARAMS is a real native struct (VARIANTARG *rgvarg; DISPID
      # *rgdispidNamedArgs; UINT cArgs; UINT cNamedArgs;) read positionally
      # by the callee, so — unlike a VARIANT's self-designed padding — the
      # two pointer fields MUST be packed at native pointer width (PACK_PTR),
      # not a hardcoded 8 bytes, or this silently misaligns cArgs/cNamedArgs
      # on x86.
      dispparams = [
        arg_variants.empty? ? 0 : W.native_address_of(args_blob),
        named_blob.empty? ? 0 : W.native_address_of(named_blob),
        arg_variants.size,
        named_put ? 1 : 0
      ].pack("#{W::PACK_PTR}#{W::PACK_PTR}LL")
      keep_alive << dispparams

      excepinfo = ("\x00" * W::EXCEPINFO_SIZE).b
      result = ("\x00" * W::VARIANT_SIZE).b

      hr = invoke_fn.call(@ptr, dispid, W::IID_NULL, 0, wflags, dispparams, result, excepinfo, nil)
      [hr, result, excepinfo]
    end

    def release
      return unless @ptr && !@ptr.zero?

      release_fn.call(@ptr)
      @ptr = nil
    end

    private

    def get_ids_of_names_fn
      @get_ids_of_names_fn ||= W.vtable_function(
        @ptr, 5, [W::VOIDP, W::VOIDP, W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::LONG
      )
    end

    def invoke_fn
      @invoke_fn ||= W.vtable_function(
        @ptr, 6,
        [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::WORD, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP],
        W::LONG
      )
    end

    def release_fn
      @release_fn ||= W.vtable_function(@ptr, 2, [W::VOIDP], W::DWORD)
    end
  end
end
```

`ruby_value_to_variant_bytes` and `native_buffers` are defined on `WIN32OLE` itself in Task 5/6 (construction owns the keep-alive array; type marshaling is the invoke-path's job) — `dispatch.rb` only calls them, it doesn't define them, since both are specific to the class's own state, not generic IDispatch plumbing.

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/dispatch.rb
git commit -m "jruby: IDispatch GetIDsOfNames/Invoke/Release mixin"
```

Not independently runnable yet — `ruby_value_to_variant_bytes`/`native_buffers` don't exist until Task 5/6, and there's no `WIN32OLE.new` yet either. Verified together with Task 6 in CI (Task 7).

---

## Task 5: `win32ole.rb` — construction, keep-alive, finalizer

**Files:**
- Modify: `lib/win32ole/jruby/win32ole.rb`

**Interfaces:**
- Consumes: `WIN32OLE::Win32` (Task 2/3).
- Produces:
  - `WIN32OLE::RuntimeError < ::RuntimeError`
  - `WIN32OLE.new(server, host = nil)`
  - `WIN32OLE#native_buffers -> Array` (the `@__native_buffers__` keep-alive array from design §4.5)
  - A `WIN32OLE.new` that installs an `ObjectSpace.define_finalizer` calling `Release` on the raw pointer — capturing only the pointer and the vtable `Release` `Fiddle::Function`, never `self`, so the finalizer doesn't keep the object alive.

Per §6.1: try `CLSIDFromProgID` first; on failure fall back to `CLSIDFromString` (this exact fallback order, and the exact `CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER` context, come straight from `ext/win32ole/win32ole.c:2422-2447` — matched here for fidelity with the C extension's actual behavior, not the design doc's looser "`CLSCTX_SERVER`" paraphrase). `host` is Phase-1 out of scope: raise `NotImplementedError` if non-nil, per §6.1.

- [ ] **Step 1: Implement `win32ole.rb`**

```ruby
# lib/win32ole/jruby/win32ole.rb
require 'fiddle'
require 'win32ole/jruby/win32'
require 'win32ole/jruby/dispatch'

class WIN32OLE
  RuntimeError = Class.new(::RuntimeError)

  include Dispatch

  W = Win32
  private_constant :W

  def initialize(server, host = nil)
    raise NotImplementedError, 'remote OLE (host) is not supported yet' unless host.nil?

    clsid = resolve_clsid(server)
    ppv = ("\x00" * W::PTR_SIZE).b
    hr = W.co_create_instance.call(
      clsid, nil, W::CLSCTX_INPROC_SERVER | W::CLSCTX_LOCAL_SERVER, W::IID_IDISPATCH, ppv
    )
    if W.failed?(hr)
      raise WIN32OLE::RuntimeError, "#{W.unknown_server_error_message(server)}\n#{hresult_detail(hr)}"
    end

    @ptr = ppv.unpack1(W::PTR_SIZE == 8 ? 'Q' : 'L')
    install_finalizer
  end

  def native_buffers
    @native_buffers ||= []
  end

  private

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

  def install_finalizer
    ptr = @ptr
    release_fn = W.vtable_function(ptr, 2, [W::VOIDP], W::DWORD)
    ObjectSpace.define_finalizer(self, self.class.finalizer(ptr, release_fn))
  end

  def self.finalizer(ptr, release_fn)
    proc { release_fn.call(ptr) unless ptr.zero? }
  end
end
```

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/win32ole.rb
git commit -m "jruby: WIN32OLE construction, keep-alive array, GC finalizer"
```

CI-only verification — folded into Task 8's integration tests (a constructor with nothing else to call is not independently meaningful to test in isolation).

---

## Task 6: `win32ole.rb` — dynamic dispatch, type marshaling, error translation

**Files:**
- Modify: `lib/win32ole/jruby/win32ole.rb`

**Interfaces:**
- Consumes: `dispid_for`/`ole_invoke`/`release` (Task 4, via `Dispatch` mixin), `native_buffers` (Task 5), all of `WIN32OLE::Win32` (Task 2/3).
- Produces:
  - `WIN32OLE#method_missing(name, *args)` implementing §6.2's three-way branch via `Win32.dispatch_plan`.
  - `WIN32OLE#ruby_value_to_variant_bytes(value, keep_alive) -> String(VARIANT_SIZE bytes)` — the method `dispatch.rb`'s `ole_invoke` calls for each argument.
  - `WIN32OLE#variant_bytes_to_ruby_value(bytes) -> Object` — return-value marshaling per §6.4.
  - Full error translation per §6.5: `DISP_E_EXCEPTION` (`hr == -2147352567`) reads `EXCEPINFO` and formats via `Win32.parse_excepinfo`; anything else falls back to the plain HRESULT message. Wrapped in `method_error_message` (method calls) or `property_put_error_message` (property puts).

- [ ] **Step 1: Add to `lib/win32ole/jruby/win32ole.rb`** (inside the `class WIN32OLE` body, after the constructor/private section from Task 5 — these become additional public/private methods on the same class)

```ruby
  DISP_E_EXCEPTION = -2147352567 # 0x80020009

  def method_missing(name, *args)
    plan = W.dispatch_plan(name.to_s, args)
    dispid = dispid_for(plan[:name])
    if dispid.nil?
      return super(name, *args)
    end

    hr, result_bytes, excepinfo_bytes = ole_invoke(dispid, args, plan[:wflags], named_put: plan[:named_put])

    if W.failed?(hr)
      detail = error_detail(hr, excepinfo_bytes)
      message = plan[:named_put] ? W.property_put_error_message(plan[:name], detail)
                                  : W.method_error_message(plan[:name], detail)
      raise WIN32OLE::RuntimeError, message
    end

    return nil if plan[:named_put]

    variant_bytes_to_ruby_value(result_bytes)
  end

  def respond_to_missing?(name, include_private = false)
    !dispid_for(name.to_s.sub(/=\z/, '')).nil? || super
  end

  def ruby_value_to_variant_bytes(value, keep_alive)
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
        keep_alive << bstr
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
      wrap_dispatch_pointer(W.unpack_pointer(payload))
    end
  end

  private

  def wrap_dispatch_pointer(ptr)
    obj = allocate
    obj.instance_variable_set(:@ptr, ptr)
    obj.send(:install_finalizer)
    obj
  end

  def error_detail(hr, excepinfo_bytes)
    if hr == DISP_E_EXCEPTION
      info = W.parse_excepinfo(excepinfo_bytes)
      source = W.bstr_to_s(info[:bstr_source_ptr]) || '<Unknown>'
      description = W.bstr_to_s(info[:bstr_description_ptr]) || '<No Description>'
      code = info[:w_code].zero? ? info[:scode].to_s(16) : info[:w_code].to_s
      "\n    OLE error code:#{code} in #{source}\n      #{description}\n#{hresult_detail(hr)}"
    else
      "\n#{hresult_detail(hr)}"
    end
  end
```

`wrap_dispatch_pointer` uses `allocate` (bypassing `initialize`, which requires a server name) plus `install_finalizer` (already `private` from Task 5, called here via `send` since it's a different instance) — this is the "`VT_DISPATCH`/`VT_UNKNOWN` results wrap into a new `WIN32OLE` instance" requirement from §6.4, without needing a second public constructor path. No `AddRef` call is needed: COM's own convention is that an `[out]` interface pointer from `Invoke` has already been `AddRef`'d by the callee, so ownership transfers to the new wrapper as-is, and its own finalizer's eventual `Release` is exactly the one release this reference is owed.

- [ ] **Step 2: Commit**

```bash
git add lib/win32ole/jruby/win32ole.rb
git commit -m "jruby: method_missing dynamic dispatch, type marshaling, error translation"
```

Verified in Task 8 (this is the method that makes the class actually usable end-to-end; there's nothing meaningful to assert until a real COM object is involved, which needs Windows).

---

## Task 7: CI wiring — Rakefile fix + `windows.yml` jruby job

**Files:**
- Modify: `Rakefile:18-19`
- Modify: `.github/workflows/windows.yml`

**Interfaces:** none (build/CI configuration only).

Today's `Rakefile` line 18-19:
```ruby
if /mswin|mingw|cygwin/ =~ RbConfig::CONFIG['host_os']
  task :test => :compile
end
```
On a JRuby-on-Windows CI runner, `RbConfig::CONFIG['host_os']` still matches `mingw`/`mswin`, so `rake test` would pull in `:compile` — which tries to build the C extension. JRuby cannot build C extensions, so this must be skipped specifically for JRuby regardless of host OS (this is exactly the packaging gap the design doc flagged as needing confirmation in §4.1).

- [ ] **Step 1: Fix the Rakefile**

```ruby
if RUBY_ENGINE == 'ruby' && /mswin|mingw|cygwin/ =~ RbConfig::CONFIG['host_os']
  task :test => :compile
end
```

- [ ] **Step 2: Add a `test-jruby` job to `.github/workflows/windows.yml`**

Appended as a new top-level job (kept separate from the existing `ruby-versions`/`test` cruby matrix rather than folded into it, since this repo has no prior evidence the shared `ruby/actions/ruby_versions.yml` reusable workflow enumerates a `jruby` engine the same way, and the spike's own CI (`tmp_spike_ffi_com.yml`) already proved a direct `ruby/setup-ruby@v1` + `ruby-version: 'jruby'` step works standalone):

```yaml
  test-jruby:
    name: build (jruby)
    runs-on: windows-latest
    steps:
    - name: git config
      run: |
        git config --global core.autocrlf false
        git config --global core.eol lf
        git config --global advice.detachedHead 0
    - uses: actions/checkout@v7
    - uses: ruby/setup-ruby@v1
      with:
        ruby-version: 'jruby'
        bundler-cache: true
    - name: Install ffi gem (Fiddle's JRuby backend needs it)
      run: gem install ffi
    - name: Run test
      run: bundle exec rake
```

- [ ] **Step 3: Push and confirm the workflow starts**

Run: `git push -u origin jruby-support` (or push to whatever branch is active), then check the Actions tab for the `windows` workflow's new `build (jruby)` job.
Expected at this point: the job **fails** at `Run test`, but on real assertion failures/errors from Tasks 1–6's tests — not on a build-system or missing-dependency error. That failure signal is exactly what Task 8 exists to turn into passing tests; if instead it fails on something like "cannot find gem ffi" or "rake aborted: :compile", that's this task's own bug (dependency install order, or the Rakefile guard didn't take) and must be fixed here before moving on.

- [ ] **Step 4: Commit**

```bash
git add Rakefile .github/workflows/windows.yml
git commit -m "ci: skip native :compile task and add a jruby CI job"
```

---

## Task 8: Phase 1 integration tests (CI-verified) + `GC.stress` probe

**Files:**
- Create: `test/win32ole/jruby/test_win32ole_phase1.rb`

**Interfaces:** none produced — this is the task that finally exercises Tasks 1–6 against real COM objects.

Per design §7: exercise this against `Scripting.Dictionary`/`Scripting.FileSystemObject` (preinstalled on every Windows runner, same objects the spike used), covering exactly Phase 1's scope — construction, get/put dispatch, basic type marshaling, error message format, `NoMethodError` on unknown members, `VT_DISPATCH` wrapping, and a `GC.stress` pass on the invoke path (the design's §7 third bullet and §8 risk #2, since a keep-alive bug is exactly the kind that "passes every normal test run and then crashes intermittently").

- [ ] **Step 1: Write the tests**

```ruby
# test/win32ole/jruby/test_win32ole_phase1.rb
begin
  require 'win32ole'
rescue LoadError
end
require 'test/unit'

if defined?(WIN32OLE) && RUBY_ENGINE == 'jruby'
  class TestWin32OLEPhase1 < Test::Unit::TestCase
    def setup
      @dict = WIN32OLE.new('Scripting.Dictionary')
    end

    def test_new_by_progid
      assert_kind_of(WIN32OLE, @dict)
    end

    def test_new_unknown_progid_raises
      exc = assert_raise(WIN32OLE::RuntimeError) { WIN32OLE.new('NonExistProgID999') }
      assert_match(/^unknown OLE server: `NonExistProgID999'/, exc.message) #`
    end

    def test_two_arg_method_and_one_arg_method
      @dict.add('a', 1000)
      assert_equal(1000, @dict.item('a'))
    end

    def test_property_put_and_get
      @dict.compareMode = 1
      @dict.add('one', 1)
      assert_equal(1, @dict.item('ONE'))
    end

    def test_raise_message_on_wrong_arg_count
      exc = assert_raise(WIN32OLE::RuntimeError) { @dict.add }
      assert_match(/^\(in OLE method `add': \)/, exc.message) #`
    end

    def test_raise_message_on_bad_property_put
      exc = assert_raise(WIN32OLE::RuntimeError) { @dict.compareMode = -1 }
      assert_match(/^\(in setting property `compareMode': \)/, exc.message) #`
    end

    def test_no_method_error
      exc = assert_raise(NoMethodError) { @dict.non_exist_method }
      assert_match(/non_exist_method/, exc.message)
      assert_kind_of(WIN32OLE, exc.receiver)
    end

    def test_dispatch_return_value_wraps_as_win32ole
      fso = WIN32OLE.new('Scripting.FileSystemObject')
      drives = fso.Drives
      assert_kind_of(WIN32OLE, drives)
    end

    def test_string_bool_and_nil_marshaling
      fso = WIN32OLE.new('Scripting.FileSystemObject')
      assert_equal('foo', fso.GetBaseName('C:\\Temp\\foo.txt'))
      assert_equal(true, fso.FileExists('C:\\Windows\\System32\\drivers\\etc\\hosts'))
    end

    def test_gc_stress_survives_repeated_invoke
      GC.stress = true
      100.times do |i|
        @dict.add("key#{i}", i)
        assert_equal(i, @dict.item("key#{i}"))
      end
    ensure
      GC.stress = false
    end
  end
end
```

- [ ] **Step 2: Push and confirm CI passes**

Run: `git push` (add-and-commit first, see Step 3), then check the `build (jruby)` job in GitHub Actions.
Expected: PASS. If any test fails, use `superpowers:systematic-debugging` on the CI log rather than guessing — the earlier tasks' spike-derived code has already been validated for the 0/1/2-arg cases the spike covered, so a failure here most likely means a Phase-1-specific addition (property-put, error formatting, `GC.stress`, `VT_DISPATCH` wrapping) has a real bug worth root-causing, not a re-run-and-hope situation.

- [ ] **Step 3: Commit**

```bash
git add test/win32ole/jruby/test_win32ole_phase1.rb
git commit -m "test: Phase 1 integration coverage for jruby win32ole backend"
```

---

## Task 9: Cleanup — remove the spike

**Files:**
- Delete: `tmp_spike/ffi_com_spike.rb`, `tmp_spike/fiddle_com_spike.rb`, `tmp_spike/fiddle_com_event_spike.rb`
- Delete: `.github/workflows/tmp_spike_ffi_com.yml`

**Interfaces:** none.

Every one of these files' own header comment says "THROWAWAY" and, in the workflow file's case, explicitly "Delete this file once the spike is done." The spike's findings are now encoded as real, tested library code (Tasks 1–8) and as the design doc (§1.2), so nothing is lost by deleting them. `fiddle_com_event_spike.rb` covers Phase 4 (`WIN32OLE::Event`) groundwork that this plan doesn't touch — leaving it would be dead weight against a plan that's already landed the mechanism it explored (vtable/closure basics are superseded by this plan's real `vtable_function`; the event-sink-specific closure work stays relevant for a future Phase 4 plan, but the file itself is a script, not reusable library code, so there's nothing to port forward mechanically).

- [ ] **Step 1: Delete the files**

```bash
git rm -r tmp_spike/ .github/workflows/tmp_spike_ffi_com.yml
```

- [ ] **Step 2: Confirm nothing else references them**

Run: `grep -rn "tmp_spike" --include="*.rb" --include="*.yml" --include="*.md" .`
Expected: no matches outside `docs/superpowers/specs/2026-09-22-jruby-win32ole-support-design.md`'s historical §1.2 narrative (which should stay — it's documenting what was learned, not a live reference).

- [ ] **Step 3: Commit**

```bash
git commit -m "chore: remove throwaway jruby COM spike now superseded by lib/win32ole/jruby"
```

---

## Explicit scope cut (surfaced, not silently made)

Design §7's first bullet says to run "the existing `test/win32ole/*.rb` suite unmodified" against the new backend, with per-test `skip`/guard additions for Phase 2+ functionality. This plan does **not** do that retrofit — it adds a new, Phase-1-scoped test file (Task 8) instead of touching the ~20 existing test files (most of which exercise `ole_methods`/`ole_typelib`/`WIN32OLE::TYPE` and friends — all Phase 2+, none of it built by this plan). Reasons:

1. Most of those files mix Phase 1 and Phase 2+ assertions in the same test methods (e.g. `test_win32ole.rb`'s `TestCaseForDict` module), so "guard individual tests" would mean rewriting a large fraction of them method-by-method — a mechanical but sizable task disproportionate to "Phase 1 only," and one that would need re-touching again once Phase 2 lands anyway.
2. Since `defined?(WIN32OLE)` is false today for JRuby, adding the new `jruby` CI job (Task 7) means those existing files will, for the first time, actually attempt to run under JRuby — and fail loudly on Phase 2+ APIs that don't exist yet, which is a regression in CI signal (a previously-green job going red) even though no Phase 1 code is at fault.

Recommendation: retrofit the legacy suite with skip-guards as its own follow-up pass once this plan has landed — at that point there's working Phase 1 code to check each test against, so it's a mechanical "does this test need `WIN32OLE::TYPE`? guard it" pass rather than a speculative one done ahead of time.
