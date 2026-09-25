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

class TestVariantByRefRoundTrip < Test::Unit::TestCase
  def test_mutating_realvar_bytes_is_visible_through_the_byref_pointer
    v = WIN32OLE::Variant.new(42, WIN32OLE::VariantType::VT_I4 | WIN32OLE::VariantType::VT_BYREF)
    realvar_ptr = v.instance_variable_get(:@realvar)
    var = v.instance_variable_get(:@var)

    _vt, payload = WIN32OLE::Win32.unpack_variant(var)
    ptr_into_realvar = WIN32OLE::Win32.unpack_pointer(payload)
    assert_equal(realvar_ptr.to_i + 8, ptr_into_realvar)

    # Simulate an out-parameter callee overwriting *ptr_into_realvar in place
    Fiddle::Pointer.new(ptr_into_realvar)[0, 4] = [99].pack('l')
    assert_equal(99, WIN32OLE::Win32.unpack_i4(realvar_ptr[8, 4]))
  end
end
end
