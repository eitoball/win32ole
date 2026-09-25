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

class TestVariantArrayIndexValidation < Test::Unit::TestCase
  # array_state's psa is real native memory only on Windows; here we stub
  # SafeArray.safe_array_get_dim's underlying Fiddle::Function so the
  # dimension-count check under test runs (and raises) before any real
  # SafeArray* native call is reached, so this is runnable off-Windows.
  def setup
    @original_dim_fn = WIN32OLE::SafeArray.instance_variable_get(:@safe_array_get_dim)
    fake = Object.new
    def fake.call(_psa)
      2
    end
    WIN32OLE::SafeArray.instance_variable_set(:@safe_array_get_dim, fake)
  end

  def teardown
    WIN32OLE::SafeArray.instance_variable_set(:@safe_array_get_dim, @original_dim_fn)
  end

  def build_array_variant
    v = WIN32OLE::Variant.allocate
    vt = WIN32OLE::VariantType::VT_I4 | WIN32OLE::VariantType::VT_ARRAY
    v.instance_variable_set(:@var, WIN32OLE::Win32.pack_variant(vt, WIN32OLE::Win32.pack_pointer(0x10000)))
    v
  end

  def test_bracket_read_raises_argument_error_on_mismatched_index_count
    v = build_array_variant
    assert_raise(ArgumentError) { v[0] }
  end

  def test_bracket_write_raises_argument_error_on_mismatched_index_count
    v = build_array_variant
    assert_raise(ArgumentError) { v[0] = 1 }
  end
end
end
