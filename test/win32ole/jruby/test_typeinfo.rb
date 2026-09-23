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
    assert_equal(PTR64 ? 96 : 76, TI::TYPEATTR.size)
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

  def test_query_interface_error_message_matches_method_error_message_shape
    msg = WIN32OLE::Win32.query_interface_error_message('GetTypeInfo', 'boom')
    assert_match(/\Afailed to GetTypeInfo: boom\z/, msg)
  end
end
