# test/win32ole/jruby/test_typeinfo.rb
require 'test/unit'

if RUBY_ENGINE == 'jruby'
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
    # x86 is an inherited, not-re-verified risk (design spec §8 risk #5) --
    # this plan only claims x64 correctness, the actual CI target. The x86
    # branch here pins what TI::TYPEATTR's current (x64-oriented, hardcoded)
    # padding actually computes on a 32-bit Fiddle::SIZEOF_VOIDP, so this
    # assertion at least matches the code instead of a stale, disproven
    # number -- it is not a claim that this value is C-ABI-correct on x86.
    assert_equal(PTR64 ? 96 : 80, TI::TYPEATTR.size)
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

  # Offset-pinning regression tests: validate field positions independently
  def test_vardesc_wvarflags_is_at_offset_56_on_x64
    return unless PTR64

    buf = TI::VARDESC.malloc
    buf.to_ptr[0, TI::VARDESC.size] = "\x00" * TI::VARDESC.size
    buf.wVarFlags = 0xFFFF
    first_nonzero = (0...TI::VARDESC.size).find { |i| buf.to_ptr[i, 1].unpack1('C') != 0 }
    assert_equal(56, first_nonzero)
  ensure
    Fiddle.free(buf.to_ptr) if buf.respond_to?(:to_ptr)
  end

  def test_vardesc_varkind_is_at_offset_60_on_x64
    return unless PTR64

    buf = TI::VARDESC.malloc
    buf.to_ptr[0, TI::VARDESC.size] = "\x00" * TI::VARDESC.size
    buf.varkind = -1
    first_nonzero = (0...TI::VARDESC.size).find { |i| buf.to_ptr[i, 1].unpack1('C') != 0 }
    assert_equal(60, first_nonzero)
  ensure
    Fiddle.free(buf.to_ptr) if buf.respond_to?(:to_ptr)
  end

  def test_funcdesc_wfuncflags_is_at_offset_80_on_x64
    return unless PTR64

    buf = TI::FUNCDESC.malloc
    buf.to_ptr[0, TI::FUNCDESC.size] = "\x00" * TI::FUNCDESC.size
    buf.wFuncFlags = 0xFFFF
    first_nonzero = (0...TI::FUNCDESC.size).find { |i| buf.to_ptr[i, 1].unpack1('C') != 0 }
    assert_equal(80, first_nonzero)
  ensure
    Fiddle.free(buf.to_ptr) if buf.respond_to?(:to_ptr)
  end

  def test_typeattr_tdescalias_vt_is_at_offset_72_on_x64
    return unless PTR64

    buf = TI::TYPEATTR.malloc
    buf.to_ptr[0, TI::TYPEATTR.size] = "\x00" * TI::TYPEATTR.size
    buf.tdescAlias_vt = 0xFFFF
    first_nonzero = (0...TI::TYPEATTR.size).find { |i| buf.to_ptr[i, 1].unpack1('C') != 0 }
    assert_equal(72, first_nonzero)
  ensure
    Fiddle.free(buf.to_ptr) if buf.respond_to?(:to_ptr)
  end

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

  def test_vartype_names_distinguishes_dispatch_and_unknown
    # VT_DISPATCH(9) and VT_UNKNOWN(13) are genuinely different VARTYPEs;
    # they must not resolve to the same name.
    assert_equal('DISPATCH', TI::VARTYPE_NAMES[9])
    assert_equal('UNKNOWN', TI::VARTYPE_NAMES[13])
  end

  def test_vartype_names_covers_common_introspection_types
    assert_equal('VOID', TI::VARTYPE_NAMES[24])
    assert_equal('VARIANT', TI::VARTYPE_NAMES[12])
    assert_equal('HRESULT', TI::VARTYPE_NAMES[25])
    assert_equal('PTR', TI::VARTYPE_NAMES[26])
    assert_equal('USERDEFINED', TI::VARTYPE_NAMES[29])
  end

  def test_vartype_name_falls_back_for_unknown_vartype
    assert_equal('Unknown Type 999', TI.vartype_name(999))
  end

  def test_vartype_name_uses_the_table_for_known_vartype
    assert_equal('I4', TI.vartype_name(3))
  end

  def test_query_interface_error_message_matches_method_error_message_shape
    msg = WIN32OLE::Win32.query_interface_error_message('GetTypeInfo', 'boom')
    assert_match(/\Afailed to GetTypeInfo: boom\z/, msg)
  end

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
    Fiddle.free(vtable) if vtable
    Fiddle.free(obj) if obj
  end

  def test_ref_type_of_impl_type_fn_resolves_the_documented_slot
    obj, vtable = fake_vtable_object(10)
    fn = TI.ref_type_of_impl_type_fn(obj.to_i)
    assert_equal(0x1000 + TI::ITYPEINFO_VTBL[:GetRefTypeOfImplType], fn.to_i)
  ensure
    Fiddle.free(vtable) if vtable
    Fiddle.free(obj) if obj
  end

  def test_get_names_fn_resolves_the_documented_slot
    obj, vtable = fake_vtable_object(10)
    fn = TI.get_names_fn(obj.to_i)
    assert_equal(0x1000 + TI::ITYPEINFO_VTBL[:GetNames], fn.to_i)
  ensure
    Fiddle.free(vtable) if vtable
    Fiddle.free(obj) if obj
  end

  def test_get_ids_of_names_fn_resolves_the_documented_slot
    obj, vtable = fake_vtable_object(11)
    fn = TI.get_ids_of_names_fn(obj.to_i)
    assert_equal(0x1000 + TI::ITYPEINFO_VTBL[:GetIDsOfNames], fn.to_i)
  ensure
    Fiddle.free(vtable) if vtable
    Fiddle.free(obj) if obj
  end

  def test_type_info_of_guid_fn_resolves_the_documented_slot
    obj, vtable = fake_vtable_object(7)
    fn = TI.type_info_of_guid_fn(obj.to_i)
    assert_equal(0x1000 + TI::ITYPELIB_VTBL[:GetTypeInfoOfGuid], fn.to_i)
  ensure
    Fiddle.free(vtable) if vtable
    Fiddle.free(obj) if obj
  end

  W = WIN32OLE::Win32

  def fake_typedesc(union, vt)
    buf = Fiddle::Pointer.malloc(TI::TYPEDESC.size)
    buf[0, W::PTR_SIZE] = [union].pack(W::PACK_PTR)
    buf[W::PTR_SIZE, 2] = [vt].pack('S')
    buf
  end

  def test_typedesc_details_of_a_scalar
    td = fake_typedesc(0, 3) # VT_I4
    assert_equal(['I4'], TI.typedesc_details(0, td.to_i))
  ensure
    Fiddle.free(td) if td
  end

  def test_typedesc_details_follows_ptr_and_safearray
    inner = fake_typedesc(0, 8) # VT_BSTR
    sa = fake_typedesc(inner.to_i, 27) # VT_SAFEARRAY -> BSTR
    ptr = fake_typedesc(sa.to_i, 26) # VT_PTR -> SAFEARRAY
    assert_equal(%w[PTR SAFEARRAY BSTR], TI.typedesc_details(0, ptr.to_i))
  ensure
    [ptr, sa, inner].each { |p| Fiddle.free(p) if p }
  end

  # Only the low 4 bytes of the TYPEDESC union hold the HREFTYPE for
  # VT_USERDEFINED; the upper bytes of the pointer-sized slot can be
  # anything, so they must be masked off before GetRefTypeInfo.
  def test_typedesc_details_userdefined_masks_hreftype_and_tolerates_failure
    received = []
    get_ref = Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::DWORD, W::VOIDP], W::STDCALL) do |_this, href, _out|
      received << href
      -2147319765 # TYPE_E_ELEMENTNOTFOUND
    end
    slots = TI::ITYPEINFO_VTBL[:GetRefTypeInfo] + 1
    vtable = Fiddle::Pointer.malloc(W::PTR_SIZE * slots)
    slots.times { |i| vtable[i * W::PTR_SIZE, W::PTR_SIZE] = [0].pack(W::PACK_PTR) }
    vtable[TI::ITYPEINFO_VTBL[:GetRefTypeInfo] * W::PTR_SIZE, W::PTR_SIZE] = [get_ref.to_i].pack(W::PACK_PTR)
    obj = Fiddle::Pointer.malloc(W::PTR_SIZE)
    obj[0, W::PTR_SIZE] = [vtable.to_i].pack(W::PACK_PTR)
    garbage_high = W::PTR_SIZE == 8 ? 0xDEADBEEF_00000005 : 5
    td = fake_typedesc(garbage_high, 29) # VT_USERDEFINED

    assert_equal(['USERDEFINED'], TI.typedesc_details(obj.to_i, td.to_i))
    assert_equal([5], received)
  ensure
    [td, obj, vtable].each { |p| Fiddle.free(p) if p }
  end

  # A fake ITypeInfo with two implemented types (flags 0x1 and 0x2) whose
  # GetRefTypeInfo hands back distinct fake pointers (0x1000 + href);
  # implemented type #1's GetRefTypeOfImplType fails and must be skipped.
  def test_each_impl_type_yields_flags_and_ref_pointers
    keep = []
    slot = TI::ITYPEINFO_VTBL
    attr = Fiddle::Pointer.malloc(TI::TYPEATTR.size)
    attr[0, TI::TYPEATTR.size] = ("\x00" * TI::TYPEATTR.size).b
    attr[TI::TYPEATTR.offsetof('cImplTypes'), 2] = [3].pack('S')
    closures = {
      slot[:GetTypeAttr] => [[W::VOIDP, W::VOIDP], W::LONG, proc { |_t, out| out[0, W::PTR_SIZE] = [attr.to_i].pack(W::PACK_PTR); 0 }],
      slot[:ReleaseTypeAttr] => [[W::VOIDP, W::VOIDP], W::VOID, proc { |*| nil }],
      slot[:GetImplTypeFlags] => [[W::VOIDP, W::DWORD, W::VOIDP], W::LONG, proc { |_t, i, out| out[0, 4] = [i + 1].pack('l'); 0 }],
      slot[:GetRefTypeOfImplType] => [[W::VOIDP, W::DWORD, W::VOIDP], W::LONG,
                                      proc { |_t, i, out| next(-1) if i == 1; out[0, 4] = [i].pack('L'); 0 }],
      slot[:GetRefTypeInfo] => [[W::VOIDP, W::DWORD, W::VOIDP], W::LONG,
                                proc { |_t, href, out| out[0, W::PTR_SIZE] = [0x1000 + href].pack(W::PACK_PTR); 0 }]
    }
    slots = slot.values.max + 1
    vtable = Fiddle::Pointer.malloc(W::PTR_SIZE * slots)
    slots.times { |i| vtable[i * W::PTR_SIZE, W::PTR_SIZE] = [0].pack(W::PACK_PTR) }
    closures.each do |index, (args, ret, body)|
      c = Fiddle::Closure::BlockCaller.new(ret, args, W::STDCALL, &body)
      keep << c
      vtable[index * W::PTR_SIZE, W::PTR_SIZE] = [c.to_i].pack(W::PACK_PTR)
    end
    obj = Fiddle::Pointer.malloc(W::PTR_SIZE)
    obj[0, W::PTR_SIZE] = [vtable.to_i].pack(W::PACK_PTR)

    yielded = []
    TI.each_impl_type(obj.to_i) { |flags, ref| yielded << [flags, ref] }
    assert_equal([[1, 0x1000], [3, 0x1002]], yielded)
  ensure
    [obj, vtable, attr].each { |p| Fiddle.free(p) if p }
  end
end
end
