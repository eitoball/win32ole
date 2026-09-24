# JRuby support for win32ole — Phase 3 (Record / Variant) design

Date: 2026-09-24
Status: proposed

## 1. Background

Phase 1 (`docs/superpowers/specs/2026-09-22-jruby-win32ole-support-design.md`)
shipped `WIN32OLE` core dispatch with a Ruby↔VARIANT type table limited to
`String`/`Integer`/`Float`/`true`/`false`/`nil`/`WIN32OLE` (its own §6.3
explicitly deferred "Arrays, Records, and explicit `WIN32OLE::Variant`
wrapping" as out of scope). Phase 2
(`docs/superpowers/specs/2026-09-22-jruby-win32ole-phase2-typelib-design.md`)
shipped `WIN32OLE::TYPE`/`TYPELIB`/`METHOD`/`PARAM`/`VARIABLE` — type-library
introspection via `ITypeInfo`/`ITypeLib` — and closed clean (final review,
9 findings fixed, CI-confirmed). Both phases are complete and merged into
this branch's history.

Phase 3 is `WIN32OLE::Record` and `WIN32OLE::Variant`, the two remaining
`VARIANT`-adjacent classes from the original roadmap's Phase 3 row. Reading
the real C extension (`ext/win32ole/win32ole_record.c`, 609 lines;
`win32ole_variant.c`, 738 lines; `win32ole_variant_m.c`, 153 lines) resolved
what would otherwise be design guesswork:

### 1.1 What `WIN32OLE::Record` actually needs (verified against MRI source)

- Construction (`Record.new(typename, oleobj)`) resolves an `ITypeLib*` from
  the second argument — either `WIN32OLE#ole_typelib`-equivalent logic (if
  given a live `WIN32OLE` instance) or a `WIN32OLE::TypeLib`'s own pointer
  directly — then linearly scans the typelib's members
  (`GetTypeInfoCount`/`GetDocumentation`/`GetTypeInfo`, one call per member
  until a name match) and calls `GetRecordInfoFromTypeInfo` (a standalone
  `oleaut32.dll` export, not a vtable member) on the matching `ITypeInfo*`
  to obtain an `IRecordInfo*`. This is *not* the registry-tree walk Phase
  2's §3 excluded — it walks one already-open typelib's own member list,
  the same operation Phase 2's `TypeLib#ole_types` already performs, just
  without going through the `WIN32OLE::Type` wrapper (name-matching stops
  as soon as a `GetDocumentation` name hits, before a `Type` object would
  ever be built).
- `IRecordInfo` is a **third custom (non-`IDispatch`) COM interface** with
  its own fixed vtable, alongside Phase 2's `ITypeInfo`/`ITypeLib`. Needed
  members: `RecordInit`, `GetName`, `GetSize`, `GetFieldNoCopy`, `PutField`,
  `GetFieldNames` (§4.4 has the full slot table).
- A `WIN32OLE::Record` is a thin Ruby object: a live `IRecordInfo*`, a raw
  native buffer holding the record's field data (`RecordInit`-initialized,
  sized by `GetSize`), and a Ruby `Hash` snapshot of field name → value
  populated once at construction time (via `GetFieldNoCopy` per field) and
  used for both reads (`to_h`, `method_missing` getters) and writes
  (`method_missing` setters mutate the Hash; the record is only re-marshaled
  into the native buffer, field by field via `PutField`, when the Record is
  *converted back* into a `VARIANT` — see §1.1's `ole_rec2variant`, called
  from `ole_val2variant` whenever a `WIN32OLE::Record` is passed as a method
  argument). This get/set-via-Hash, marshal-on-demand design is a direct
  port of MRI's own `folerecord_ivar_get`/`_set` + `hash2olerec`, not a new
  decision.
- MRI's own `olerecord_free` frees the native buffer with a bare `free()` —
  it does **not** call `IRecordInfo::RecordClear` first to release any
  BSTR/nested-record fields the buffer might hold. This looks like it could
  leak such fields, but it's MRI's own established (if debatable) behavior;
  Phase 3 ports it as-is for parity rather than "fixing" it unasked (§4.7
  flags the ported behavior explicitly, not silently).

### 1.2 What `WIN32OLE::Variant` actually needs (verified against MRI source)

- The overwhelming majority of `WIN32OLE::Variant`'s real value, and real
  implementation complexity, is **`SAFEARRAY` marshaling** (`VT_ARRAY`) —
  `Variant.array(dims, vt)`, `#[]`/`#[]=` (N-dimensional element
  access/mutation), and passing a plain Ruby `Array` as an OLE method
  argument/return value generically (`ole_val2variant`'s `T_ARRAY` case,
  used by *every* dispatch call, not just explicit `Variant` construction —
  confirmed by reading `win32ole.c:1279-1282`). A `WIN32OLE::Variant`
  wrapping a scalar type is comparatively simple (mostly reuses Phase 1's
  `ruby_to_variant_type`/`variant_ruby_type` machinery, extended to a wider
  VARTYPE table — §4.2).
- `VT_BYREF` is the second real piece of complexity: MRI's
  `struct olevariantdata { VARIANT realvar; VARIANT var; }` — `realvar`
  owns the actual storage, `var` is what's actually passed to Invoke, and
  for a BYREF request `var`'s scalar slot holds a *pointer into `realvar`'s
  own scalar slot* rather than a copy of the value. This is the same
  "native address embedded as data" keep-alive discipline Phase 1's design
  §4.5 already established a pattern for — `realvar` must outlive `var`.
- MRI's SAFEARRAY↔Ruby-Array conversion (`ole_val_ary2variant_ary`,
  `ole_variant2val`'s array branch) is fully N-dimensional and recursive
  (`dimension`, `ary_len_of_dim`, `ary_new_dim`, `ary_store_dim` in
  `win32ole.c`). This design ports that algorithm directly rather than
  simplifying to 1-D-only — the recursion is mechanical (walk nested Ruby
  `Array`s to find the max depth and each dimension's max length, then
  iterate `SAFEARRAY` indices in the same nested-increment pattern MRI
  uses), not a novel design risk, and 1-D-only would silently break any
  2-D range value (e.g. `Excel::Range#value`, a very common real use).
- `WIN32OLE::VariantType` (aliased `WIN32OLE::VARIANT`) is a constants-only
  module exposing every `VT_*` integer MRI defines (`win32ole_variant_m.c`),
  **regardless of which VARTYPEs this phase actually implements marshaling
  for** — existing code (including the legacy test suite) references
  constants like `WIN32OLE::VARIANT::VT_UI1` even where the corresponding
  conversion isn't exercised. Phase 3 defines the full constant table;
  §3 governs which VARTYPEs the *marshaling* code actually supports.

## 2. Goals

- Implement `WIN32OLE::Record`: construction from a `WIN32OLE`/
  `WIN32OLE::TypeLib` + type name, field get/set, `to_h`, `typename`,
  `ole_instance_variable_get`/`_set`, `inspect`, and round-tripping through
  `VT_RECORD` in both directions of ordinary dispatch (a method returning a
  struct; a method argument built from a `WIN32OLE::Record`).
- Implement `WIN32OLE::Variant`: `new(val, vartype)`, `value`, `value=`,
  `vartype`, `.array(dims, vt)`, `#[]`/`#[]=`, the `Empty`/`Null`/`Nothing`/
  `NoParam` constants, and the full `WIN32OLE::VariantType` constant module.
- Implement `VT_ARRAY` (`SAFEARRAY`) marshaling as a shared, N-dimensional
  capability used by **both** `WIN32OLE::Variant` and ordinary
  `WIN32OLE` method dispatch (a plain Ruby `Array` argument/return value),
  matching MRI's own scope — not gated behind explicit `Variant` use.
- Implement `VT_BYREF` for the scalar VARTYPE family (§4.2), enabling
  explicit out-parameter passing the way real ADO/Office APIs require it.
- Extend the scalar VARTYPE table beyond Phase 1's six types to the full
  integer family (`VT_I1`/`UI1`/`I2`/`UI2`/`I4`/`UI4`/`INT`/`UINT`/`I8`/
  `UI8`), `VT_R4`, and `VT_ERROR` — all trivial, same-shape extensions of
  Phase 1's existing 8-byte-scalar-payload pattern (different byte width,
  no new allocator or COM call), needed for `WIN32OLE::Variant.new(val, vt)`
  to support the same range of explicit VARTYPEs MRI does.

## 3. Non-goals (for this document / Phase 3)

- **`VT_CY` (currency) and `VT_DATE`.** Both need a dedicated binary
  representation (`VT_CY` is a scaled 64-bit fixed-point integer; `VT_DATE`
  is an OLE Automation date — a `double` with its own epoch/format) and
  their own Ruby-side conversion (MRI: `rbtime2vtdate`/`vtdate2rbtime` for
  dates; currency has no natural Ruby type at all). Per Phase 1's own §6.4
  policy (raise `NotImplementedError` for anything not explicitly
  supported, rather than MRI's own fallback of stringifying via
  `VariantChangeTypeEx(..., VT_BSTR)` for any unhandled VARTYPE — see
  `win32ole.c`'s `ole_variant2val` default case) — Phase 3 keeps that
  policy: `VT_CY`/`VT_DATE` values raise `NotImplementedError`, they do
  **not** silently come back as a `String`. Implementers should not port
  MRI's string-fallback `default:` case verbatim — doing so would silently
  contradict Phase 1's established, deliberate design choice.
- **`WIN32OLE::Record.new`'s registry-adjacent scan is in scope (§1.1), but
  `WIN32OLE::TypeLib.typelibs`-style registry *enumeration* remains out of
  scope**, unchanged from Phase 2 §3.
- **`SafeArrayGetRecordInfo`-based record-typed array elements** (a
  `SAFEARRAY` whose element type is `VT_RECORD`) are not specially
  exercised by this design's own tests, though the general N-dimensional
  array algorithm (§1.2, §4.3) handles them via the same code path MRI
  uses (`ole_variant2val`'s `vt_base == VT_RECORD` branch) — this falls out
  of the ported algorithm rather than needing dedicated design, but isn't a
  claimed, verified capability the way scalar-element arrays are.
- **`IRecordInfo::RecordClear`/`RecordCopy`/`RecordCreate`/
  `RecordCreateCopy`/`RecordDestroy`/`IsMatchingType`/`GetTypeInfo`/`GetGuid`
  /`PutFieldNoCopy`/`GetField`** (as opposed to `GetFieldNoCopy`) — not
  needed by any in-scope entry point (§4.4's vtable table lists them for
  documentation completeness, unbound).
- **Encoding/locale fidelity and x86 (32-bit) verification** — inherited,
  unchanged, from Phase 1 §8 risks #5/#3 and Phase 2 §8 risk #5.

## 4. Architecture

### 4.1 File layout

```
lib/win32ole/jruby/
  array.rb       # NEW: SAFEARRAY struct + oleaut32 SafeArray* Fiddle bindings,
                 #      N-dimensional Ruby Array <-> SAFEARRAY conversion
  record.rb      # NEW: WIN32OLE::Record; IRecordInfo vtable slots + the one
                 #      standalone GetRecordInfoFromTypeInfo binding, inline
                 #      (small enough not to warrant its own substrate file,
                 #      unlike typeinfo.rb's TYPEATTR/FUNCDESC/VARDESC family)
  variant.rb     # NEW: WIN32OLE::Variant, WIN32OLE::VariantType constant
                 #      module
  win32.rb       # EXTENDED: pack_variant/unpack_variant generalized from a
                 #      fixed 8-byte scalar payload to a variable-width body
                 #      (needed for VT_RECORD's 16-byte BRECORD and for
                 #      BYREF pointer arithmetic into an existing VARIANT
                 #      buffer); full scalar VT_* pack/unpack family;
                 #      VariantChangeTypeEx binding (used by BYREF's
                 #      type-coercion path, §4.2)
  win32ole.rb    # EXTENDED: ruby_value_to_variant_bytes /
                 #      variant_bytes_to_ruby_value (Phase 1) gain branches
                 #      for Ruby Array (-> array.rb), WIN32OLE::Record
                 #      (-> record.rb), and WIN32OLE::Variant/VT_RECORD/
                 #      VT_ARRAY results (-> variant.rb / array.rb)
```

### 4.2 `win32.rb` extensions: the wider VARIANT substrate

- **`pack_variant`/`unpack_variant` generalized.** Today's `pack_variant`
  requires an exactly-8-byte `payload` and zero-pads the rest of the
  buffer; `unpack_variant` always reads exactly the 8 bytes at offset 8.
  Both become body-length-aware: `pack_variant(vt, body)` accepts any
  `body.bytesize <= VARIANT_SIZE - 8` (zero-padding the remainder, as
  today); `unpack_variant(bytes, body_size: 8)` reads `body_size` bytes
  from offset 8. Every existing Phase 1/2 call site passes 8 (unchanged
  behavior); `VT_RECORD`'s `pack_record`/`unpack_record` (§4.4) pass 16 (2
  pointers: `pvRecord`, `pRecInfo`) — this is why VARIANT_SIZE is 24 on x64
  in the first place (Phase 1 §1.2: the union's widest member, `DECIMAL`,
  is 16 bytes past the 8-byte header; `BRECORD` is exactly the same width).
- **Full scalar VT_* pack/unpack family**, same 8-byte-slot shape as
  Phase 1's existing `pack_i4`/`pack_r8`/etc., added for `VT_I1`/`VT_UI1`/
  `VT_I2`/`VT_UI2`/`VT_UI4`/`VT_INT`/`VT_UINT`/`VT_UI8`/`VT_R4`/`VT_ERROR`
  (`VT_I4`/`VT_I8`/`VT_R8`/`VT_BOOL` already exist). Each is a one-line
  `Array#pack`/`String#unpack1` with a different format character
  (`c`/`C`/`s`/`S`/`L`/native-int/`Q`/`f`/`l` respectively — `VT_ERROR` is a
  plain `LONG`, packed like `VT_I4`), zero-padded to 8 bytes exactly like
  the existing ones. `VT_FOR_TYPE`/`ruby_to_variant_type`/
  `variant_ruby_type` (Phase 1) are not changed by this — those stay
  scoped to *automatic* Ruby→VARIANT inference for plain dispatch calls
  (§4.5 lists what's newly auto-inferred: `Array`, `WIN32OLE::Record`,
  `WIN32OLE::Variant`); the new scalar pack/unpack functions are for
  `WIN32OLE::Variant.new(val, explicit_vt)`'s explicit-VARTYPE path only,
  which bypasses type inference entirely (mirrors MRI: `ole_val2variant_ex`
  is a distinct function from `ole_val2variant`).
- **`VT_BYREF` construction**: given an already-packed `realvar` byte
  buffer (a Ruby `String`, kept alive by the owning `WIN32OLE::Variant`
  instance — see §4.7) and its VARTYPE, build `var`'s 8-byte body as a
  *pointer to `realvar`'s own body offset* (`native_pointer_for(realvar) +
  8`, i.e. one `PACK_PTR`-width value) with `var`'s own vt set to
  `vt | VT_BYREF`. `VT_VARIANT|VT_BYREF` is the one exception (mirrors
  MRI's `ole_set_byref`): `var`'s body holds a pointer to the *entire*
  `realvar` buffer's start (offset 0, not offset 8), since a
  `VARIANT*`-typed BYREF points at a whole VARIANT, not at one scalar
  field within it.
- **`variant_change_type` binding**: `VariantChangeTypeEx` (`oleaut32`),
  used when a `WIN32OLE::Variant` is constructed with an explicit VARTYPE
  that doesn't match the Ruby value's natural type (e.g.
  `Variant.new("2e3", VT_R4)` — MRI does this via `ole_val2variant_ex` +
  `VariantChangeTypeEx` fallback for the mismatched-type case; ports
  directly, same as Phase 1's error-path `EXCEPINFO` reads already use raw
  Fiddle bindings for other `oleaut32` calls).

### 4.3 `array.rb`: the SAFEARRAY substrate

- **`SAFEARRAY`/`SAFEARRAYBOUND` byte layout** (`oaidl.h`, hand-packed —
  flat enough for Phase 1-style `Array#pack`, no `Fiddle::Importer.struct`
  needed; unlike Phase 2's `TYPEATTR`/`FUNCDESC`/`VARDESC` (nested unions,
  self-referential `TYPEDESC`s), `SAFEARRAY` is two flat structs with no
  nesting or unions, squarely in the "tractable by hand" category Phase 2
  §2.1 itself carved out):

  ```
  SAFEARRAYBOUND: ULONG cElements; LONG lLbound;              # 8 bytes
  SAFEARRAY:      USHORT cDims; USHORT fFeatures;
                   ULONG cbElements; ULONG cLocks;
                   PVOID pvData;
                   SAFEARRAYBOUND rgsabound[cDims];  # variable-length tail
  ```

  This design only ever *reads* a `SAFEARRAY*` returned by
  `SafeArrayCreate`/out of a VARIANT — it never hand-constructs the struct
  bytes directly (`SafeArrayCreate`/`SafeArrayCreateVector` do that
  allocation natively); the layout above is documentation for `GetDim`/
  `GetLBound`/`GetUBound`-free introspection if ever needed, not a
  construction path — the design instead calls those three `oleaut32`
  functions like MRI does, rather than reading `cDims`/`rgsabound`
  directly, since the packed layout is easy to get subtly wrong for the
  same reason Phase 1 §1.2/Phase 2 §1.2 already learned twice.
- **`oleaut32` Fiddle bindings**: `SafeArrayCreate(vt, cDims, psab*) ->
  SAFEARRAY*`, `SafeArrayCreateVector(vt, lLbound, cElements) ->
  SAFEARRAY*` (used by `Variant`'s `VT_UI1|VT_ARRAY`-from-`String` fast
  path, §4.5), `SafeArrayDestroy`, `SafeArrayLock`/`Unlock`,
  `SafeArrayGetDim -> UINT`, `SafeArrayGetLBound`/`GetUBound(psa, dim,
  out LONG*) -> HRESULT`, `SafeArrayPtrOfIndex(psa, indices*, out void**)
  -> HRESULT`, `SafeArrayPutElement(psa, indices*, void*) -> HRESULT`,
  `SafeArrayAccessData`/`UnaccessData` (used by the `String`-to-`VT_UI1`
  fast path only, mirroring MRI's `ole_val2olevariantdata`'s first branch).
- **`ruby_array_to_safearray(ary, elem_vt)`**: direct port of
  `dimension`/`ary_len_of_dim`/`ole_val_ary2variant_ary`'s
  `SafeArrayCreate` + nested-index-fill loop (§1.2) — walks nested Ruby
  `Array`s once to compute `(dim_count, [size_per_dim])`, creates the
  `SAFEARRAY` with those bounds, locks it, then walks the Ruby array again
  writing each leaf value (itself one recursive VARIANT-typed pack, so
  arrays of records/arrays of arrays fall out of the recursion rather than
  needing a special case) via `SafeArrayPutElement` at each computed index
  tuple, unlocking on completion or `SafeArrayDestroy` + re-raise on
  failure (MRI: leaves a partially-filled but structurally valid array on
  a mid-loop failure; this design matches that rather than trying to
  "improve" atomicity MRI itself doesn't provide).
- **`safearray_to_ruby_array(psa, elem_vt)`**: direct port of
  `ole_variant2val`'s array branch (§1.2) — reads `GetDim`, each
  dimension's `GetLBound`/`GetUBound`, locks, then walks the same
  nested-index-increment pattern MRI's `ary_new_dim`/`ary_store_dim` use to
  build (and correctly nest) the resulting Ruby `Array`, converting each
  leaf `SAFEARRAY` slot back through the general VARIANT→Ruby path
  (`variant_bytes_to_ruby_value`, so a `SAFEARRAY` of `WIN32OLE::Record`s
  or of `VT_DISPATCH` pointers works via the same recursion, no special
  case needed beyond what §1.2's `vt_base == VT_RECORD` note already
  covers for the record-info-lookup step).
- **`VT_UI1|VT_ARRAY` ↔ `String` fast path**: when the caller passes a
  plain Ruby `String` where a `VT_UI1|VT_ARRAY` is expected (or reads one
  back), skip the generic per-element `SafeArrayPutElement`/nested-Array
  path entirely and bulk-copy via `SafeArrayAccessData`/`memcpy`-equivalent
  (`Fiddle::Pointer#[]=` on the accessed data pointer) in one call, mirror
  of `ole_val2olevariantdata`'s first branch and `folevariant_value`'s
  `dim == 1` `Array#pack('C*')` reverse path. This is a real, separate code
  path in MRI (not an optimization detail this design can skip) because
  without it, a binary blob argument would round-trip through a Ruby
  `Array` of small `Integer`s instead of staying a `String`.

### 4.4 `record.rb`: `WIN32OLE::Record` and `IRecordInfo`

- **`IRecordInfo` vtable slots** (transcribed from `oaidl.h`, same
  implementation-time-verification caveat as Phase 2 §4.2/§8's struct
  transcriptions — not yet confirmed against a live Windows build):
  `RecordInit`=3, `RecordClear`=4, `RecordCopy`=5, `GetGuid`=6, `GetName`=7,
  `GetSize`=8, `GetTypeInfo`=9, `GetField`=10, `GetFieldNoCopy`=11,
  `PutField`=12, `PutFieldNoCopy`=13, `GetFieldNames`=14,
  `IsMatchingType`=15, `RecordCreate`=16, `RecordCreateCopy`=17,
  `RecordDestroy`=18. In scope (bound to `Fiddle::Function`s): 3/7/8/11/12/
  14 (§3 lists the rest as out of scope, table kept for documentation).
- **`GetRecordInfoFromTypeInfo(ITypeInfo*, out IRecordInfo**) -> HRESULT`**:
  a standalone `oleaut32` export (not a vtable member — same shape as
  Phase 1's `CLSIDFromProgID`/`CoCreateInstance` top-level bindings).
- **Construction** (`Record.new(typename, oleobj)`): resolve an
  `ITypeLib*` — reuses Phase 2's `WIN32OLE#ole_typelib`/
  `WIN32OLE::TypeLib`'s own `@ptr` directly, no new typelib-resolution code
  — then linear-scan for the matching member name (§1.1) and call
  `GetRecordInfoFromTypeInfo`. Raises `WIN32OLE::RuntimeError` on failure
  at either step (§6 — **not** `QueryInterfaceError`; MRI's own
  `folerecord_initialize` uses `eWIN32OLERuntimeError` throughout, unlike
  Phase 2's `ITypeInfo`-family failures which use
  `eWIN32OLEQueryInterfaceError` — this is a real, easy-to-miss distinction
  if an implementer pattern-matches on Phase 2's choice instead of
  checking MRI's actual exception class per call site).
- **Field storage**: `RecordInit`-sized native buffer, allocated via
  `Fiddle::Pointer.malloc(size)` (portable across MRI/JRuby per Fiddle's
  own stdlib contract — flagged in §8 as needing the same kind of
  empirical JRuby-backend confirmation Phase 2 §8 risk #1 required for
  `Fiddle::Importer.struct`), freed with a bare `Fiddle::Pointer#free`/
  `Fiddle.free` call (no `RecordClear` first — §1.1's parity note) in the
  same finalizer that releases the `IRecordInfo*`.
- **Field access**: a Ruby `Hash` (`name => value`) populated once at
  construction via `GetFieldNames` + `GetFieldNoCopy` per field (§1.1);
  `to_h`, `ole_instance_variable_get`, and `method_missing`-as-getter all
  read this Hash directly (`Hash#fetch`, raising Ruby's own `KeyError` for
  an unknown name — matches MRI's `rb_hash_fetch` choice of exception).
  `ole_instance_variable_set`/`method_missing`-as-setter write into the
  same Hash (`Hash#[]=` after an existence-checking `fetch`, matching
  MRI's `olerecord_ivar_set`). The Hash is never re-synced from the native
  buffer after construction — matches MRI, since nothing else mutates the
  buffer between construction and the record being marshaled back out
  (next point).
- **Marshaling a `Record` into a `VARIANT`** (`WIN32OLE::Record` used as an
  OLE method argument — hooked from `win32ole.rb`, §4.5): call
  `PutField(INVOKE_PROPERTYPUT, buffer, field_name_wstr, &var)` once per
  Hash entry (skipping `nil` values, matching MRI's `hash2olerec`'s `val !=
  Qnil` guard) to write the Hash's current values into the native buffer,
  then build a `VT_RECORD` `VARIANT` whose body is `(buffer_ptr,
  irecordinfo_ptr)` (§4.2's `pack_record`).

### 4.5 `variant.rb`: `WIN32OLE::Variant`

- **`WIN32OLE::VariantType`** (aliased `WIN32OLE::VARIANT`, matching MRI's
  own backward-compatible alias — Phase 1/2 haven't needed this pattern
  yet, but it's a direct, mechanical constant-table port): every `VT_*`
  integer MRI's `win32ole_variant_m.c` defines, independent of which ones
  this phase's marshaling code actually implements (§1.2's closing point).
- **`new(val, vartype = nil)`**: one argument → reuses Phase 1's
  `ruby_value_to_variant_bytes`-equivalent *inference* path unchanged
  (matches MRI: `ole_val2variant`, the same function ordinary dispatch
  argument-building already calls — so a bare `WIN32OLE::Variant.new(x)`
  behaves identically to passing `x` directly, which is correct per MRI's
  own `folevariant_initialize`'s `len == 1` branch). Two/three arguments
  (`val`, `vartype`, optional but unused-beyond-arity-checking third
  arg — matches MRI's `rb_check_arity(len, 1, 3)`, which accepts a 3rd
  argument without using it) → explicit path: `VT_RECORD` explicitly
  raises `ArgumentError` immediately (matches MRI: "WIN32OLE::Variant does
  not support VT_RECORD... use WIN32OLE::Record instead" — the two classes
  are deliberately disjoint for record types), `VT_ARRAY`-flagged VARTYPEs
  route to `array.rb`, everything else routes to §4.2's explicit
  scalar-pack + optional `VariantChangeTypeEx` + optional `VT_BYREF`
  wrapping, mirroring `ole_val2olevariantdata`'s branch structure exactly
  (§1.2).
- **`.array(dims, vt)`**: `SafeArrayCreate` with `dims.size` dimensions
  sized from the `dims` Array, no initial values (matches MRI:
  `folevariant_s_array`, an empty/zero-filled array the caller then fills
  via `#[]=`).
- **`#[]`/`#[]=`**: `SafeArrayPtrOfIndex` (read) / `SafeArrayPutElement`
  (write) at the given index tuple, locking/unlocking around each call —
  direct port of `folevariant_ary_aref`/`_aset` (§1.2), reusing
  `array.rb`'s lock/index-pointer helpers rather than duplicating them.
- **`value`/`value=`/`vartype`**: read/replace `var`'s current value via
  the general VARIANT→Ruby/Ruby→VARIANT paths (§4.2/§4.3), with the
  `VT_UI1|VT_ARRAY`-is-really-a-`String` special case from §4.3 applied in
  `value` (matches MRI's own `folevariant_value` special case) and
  `value=` rejecting non-`String`/non-matching-VARTYPE array writes with
  `WIN32OLE::RuntimeError` (matches `folevariant_set_value`'s guard).
- **`Empty`/`Null`/`Nothing`/`NoParam` constants**: built the same way MRI
  builds them (`Variant.new(nil, VT_EMPTY)`, `Variant.new(nil, VT_NULL)`,
  `Variant.new(nil, VT_DISPATCH)`, `Variant.new(DISP_E_PARAMNOTFOUND,
  VT_ERROR)`), each exercising the explicit-VARTYPE path above — a natural
  self-test that the explicit path itself works before any external COM
  call touches it, matching MRI's own choice to build these eagerly at
  class-init time rather than lazily.

### 4.6 `win32ole.rb` (Phase 1) dispatch extension

`ruby_value_to_variant_bytes`/`variant_bytes_to_ruby_value` (Phase 1's
existing argument/return-value marshaling methods, `lib/win32ole/jruby/
win32ole.rb`) gain branches, checked before the existing
`W.ruby_to_variant_type`/`W.variant_ruby_type` dispatch:

- Ruby `Array` argument → `Array.ruby_array_to_safearray` (§4.3), wrapped
  `VT_VARIANT|VT_ARRAY` (matches MRI's `ole_val2variant`'s `T_ARRAY` case,
  which always uses `VT_VARIANT|VT_ARRAY` for an *implicit* array, as
  opposed to `Variant.array`'s caller-chosen element VARTYPE).
- `WIN32OLE::Record` argument → §4.4's `PutField`-based marshaling.
- `WIN32OLE::Variant` argument → §4.5's already-packed `var` bytes, copied
  through directly (matches MRI's `ole_variant2variant` — a plain
  `VariantCopy`, no re-inference).
- A `VT_ARRAY`-flagged or `VT_RECORD` result value → routes to
  `Array.safearray_to_ruby_array`/`record.rb`'s `create_win32ole_record`-
  equivalent instead of raising `NotImplementedError` — this is the one
  place Phase 1's own §6.4 "any VARTYPE not covered raises
  NotImplementedError" table actually grows two new covered entries
  (`VT_ARRAY`, `VT_RECORD`), consistent with that section's own framing
  ("Phase 2/3 will extend this table").

### 4.7 Resource lifetime

- **`Record`**: `IRecordInfo*` (finalizer-released, Phase 1/2's established
  pattern) + the malloc'd field buffer (freed alongside it, §4.4). A
  `Record` passed as a dispatch argument must be kept reachable by the
  caller's own keep-alive array (Phase 1 §4.5's per-call `@__native_
  buffers__`-equivalent mechanism) for the duration of the native `Invoke`
  call, exactly like a `WIN32OLE` instance's underlying pointer already is
  — the record's native buffer address is embedded as *data* inside the
  outgoing `DISPPARAMS`/`VARIANT`, not passed as a direct Fiddle::Function
  argument Fiddle could pin on its own.
- **`Variant`**: owns two packed VARIANT byte buffers (`realvar`, `var`,
  mirroring MRI's `struct olevariantdata`) as plain Ruby `String`s held in
  instance variables — ordinary Ruby object-graph reachability keeps them
  alive exactly as long as the `WIN32OLE::Variant` instance itself is
  reachable, which is sufficient since `var`'s BYREF body (§4.2) only ever
  points *into* `realvar`, never past the `Variant` object's own lifetime.
  A `SAFEARRAY*` held by either buffer's pointer body is released via
  `SafeArrayDestroy` in the `Variant`'s own finalizer (matches MRI's
  `olevariant_free`: `VariantClear` on both, which for an array-typed
  VARIANT internally calls `SafeArrayDestroy`).
- **`array.rb`'s N-dimensional helpers** allocate no long-lived state of
  their own — the `SAFEARRAY*` they build is owned by whichever `VARIANT`
  buffer embeds it (a `Variant` instance, per above, or a transient
  dispatch-argument buffer covered by §4.6's existing per-call keep-alive
  discipline).

## 5. Per-class API (Phase 3 scope)

| Class | Implemented | Notes |
|---|---|---|
| `WIN32OLE::Record` | `new`, `to_h`, `typename`, `method_missing` (get/set), `ole_instance_variable_get`, `ole_instance_variable_set`, `inspect` | Matches MRI's full public surface (§4.4) — this class has no non-goal members, unlike Type/TypeLib/Method/Param/Variable in Phase 2. |
| `WIN32OLE::Variant` | `new`, `.array`, `value`, `value=`, `vartype`, `[]`, `[]=` | Matches MRI's full public surface (§4.5) — likewise no non-goal members. |
| `WIN32OLE::VariantType` (`WIN32OLE::VARIANT`) | Full `VT_*` constant table | Constants only; §3 governs which are actually convertible. |
| `WIN32OLE` (dispatch, additions) | `Array` argument/return marshaling, `WIN32OLE::Record` argument marshaling, `WIN32OLE::Variant` argument passthrough, `VT_ARRAY`/`VT_RECORD` result unmarshaling | §4.6 — extends Phase 1's existing type table, not a new public method. |

## 6. Error translation

- `WIN32OLE::Record` construction failures (typelib resolution, member
  name not found, `GetRecordInfoFromTypeInfo` failure): `WIN32OLE::
  RuntimeError` (§4.4 — explicitly not `QueryInterfaceError`, despite the
  superficial resemblance to Phase 2's `ITypeInfo`-family failures).
- `PutField` failure (writing a `Record`'s Hash back into its native
  buffer): `WIN32OLE::RuntimeError`, message format matches Phase 1/2's
  `Win32.method_error_message`-style helpers (new
  `Win32.runtime_error_message`-shaped helper if one doesn't already exist
  in a reusable form — implementation-time detail).
- `SafeArrayCreate`/`SafeArrayLock`/`SafeArrayPutElement`/
  `SafeArrayPtrOfIndex` HRESULT failures: `WIN32OLE::RuntimeError` (matches
  MRI's `ole_raise(hr, eWIN32OLERuntimeError, ...)` call sites in
  `win32ole.c`/`win32ole_variant.c`). A `NULL` return from
  `SafeArrayCreate`/`SafeArrayCreateVector` itself (no HRESULT, an
  out-of-memory-shaped failure with no error code) raises a plain Ruby
  `RuntimeError`, **not** `WIN32OLE::RuntimeError` — matches MRI's own
  distinction (`rb_raise(rb_eRuntimeError, "memory allocation error")` vs.
  `ole_raise(hr, eWIN32OLERuntimeError, ...)`) and is easy to miss if an
  implementer defaults every failure in this area to the WIN32OLE-specific
  class.
- `Variant.new(val, VT_RECORD)`: `ArgumentError` (§4.5, matches MRI's
  explicit `rb_raise(rb_eArgError, ...)` guard, not a generic
  `TypeError`/`NotImplementedError`).
- Unknown/mismatched member name on `Record`/`ole_instance_variable_get`/
  `_set`: Ruby's own `KeyError` (via `Hash#fetch`, §4.4 — matches MRI's
  `rb_hash_fetch`), not a `WIN32OLE`-namespaced exception.

## 7. Testing / CI strategy

The legacy suite already covers this phase's scope directly:
`test/win32ole/test_win32ole_record.rb` (212 lines), `test_win32ole_
variant.rb` (722 lines), `test_win32ole_variant_m.rb` (constant-table
coverage), `test_win32ole_variant_outarg.rb` (`VT_BYREF`-specific
coverage — confirms BYREF is a real, exercised feature, not
speculative scope). Run unmodified against the new backend and triage
failures with the same discipline Phase 1/2 established: each failure is
either (a) exactly a §3 non-goal (`VT_CY`/`VT_DATE`, the unbound
`IRecordInfo` members) — expected, not actionable — or (b) a real Phase 3
bug, fixed and re-verified.

Additional tests this phase should add, beyond running the legacy suite:

- A local, OS-independent unit test for `array.rb`'s N-dimensional
  index-walk logic (`ruby_array_to_safearray`'s dimension/size inference,
  `safearray_to_ruby_array`'s nested-index reconstruction) against
  synthetic nested-Array fixtures — this is genuinely novel Ruby logic
  (not a native call), the same rationale Phase 2 §7 used for testing
  `Fiddle::Importer.struct` layouts locally before anything depends on
  them.
- A `GC.stress`-enabled test for `Record`'s and `Variant`'s finalizer
  paths (native buffer + `IRecordInfo*`/`SAFEARRAY*`), matching Phase 1/2's
  own rationale (§7 of each): this is exactly the class of bug that
  "passes every normal test run and then crashes intermittently."
- A `VT_BYREF` round-trip test exercising `Variant.new(val, VT_I4 |
  VT_BYREF)` followed by mutating the referenced value through a live OLE
  out-parameter call, confirming `realvar` and `var` actually stay
  correctly linked (this is the one piece of §4.2/§4.7 with no
  Phase-1/2-established precedent to lean on — the keep-alive discipline
  there is new, not just reused).

## 8. Risks / open questions carried forward

1. **`IRecordInfo`'s vtable slot transcription (§4.4) is unverified against
   a live Windows build**, same caveat as Phase 2 §8 risk #2 for
   `TYPEATTR`/`FUNCDESC`/`VARDESC`. The implementation plan's first task
   must validate it (e.g. call `GetSize` on a known real record type and
   sanity-check the returned size) before anything else depends on it.
2. **`Fiddle::Pointer.malloc`/`#free` portability across MRI's and JRuby's
   Fiddle backends is unverified** (§4.4) — the same kind of empirical bet
   Phase 2 §8 risk #1 flagged for `Fiddle::Importer.struct`, needs
   confirming early rather than assumed.
3. **The N-dimensional `SAFEARRAY` algorithm (§4.3) is a direct port of
   MRI's own recursive logic, not independently re-derived or simplified**
   — correct by construction only to the extent the port is faithful;
   the implementation plan should diff its own index-walk against MRI's
   `ary_new_dim`/`ary_store_dim`/`dimension`/`ary_len_of_dim` line-by-line
   rather than reimplementing from a written description of the algorithm
   alone.
4. **MRI's own `olerecord_free` not calling `RecordClear` (§1.1, §4.4) is
   ported as-is for parity**, not fixed — if this turns out to be a real
   leak worth closing, that's a decision for MRI's own maintainers (a
   behavior change to the reference implementation, out of scope for a
   JRuby-parity shim) rather than something this phase should quietly
   diverge on.
5. **x86 (32-bit) unverified**, inherited unchanged from Phase 1 §8 risk #3
   / Phase 2 §8 risk #5. `VT_RECORD`'s `BRECORD` body is 2 native pointers
   regardless of width (8 bytes on x86, 16 on x64) — §4.2's generalized
   `pack_variant`/`unpack_variant` is written to be width-correct on paper
   for both, but, per both prior phases' own history, "on paper" is
   exactly the part that needs an actual run to trust.
6. **Performance**: `SAFEARRAY` marshaling adds one `Fiddle::Function` call
   per element for both directions (`SafeArrayPutElement`/
   `SafeArrayPtrOfIndex`), on top of Phase 1/2's already-accepted
   Fiddle-call-overhead trade-off (Phase 1 §8 risk #4). A large 2-D range
   read (e.g. a big Excel sheet) could be meaningfully slower than the C
   extension's direct-memory-access equivalent. Accepted as a known cost,
   not addressed by this design, consistent with both prior phases'
   stance on the same trade-off.
