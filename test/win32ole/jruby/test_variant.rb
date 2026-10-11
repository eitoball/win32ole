require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/win32ole'

class TestVariantType < Test::Unit::TestCase
  def test_vt_i4_matches_win32_constant
    assert_equal(WIN32OLE::Win32::VT_I4, WIN32OLE::VariantType::VT_I4)
  end

  def test_vt_array_and_byref_flags
    assert_equal(0x2000, WIN32OLE::VariantType::VT_ARRAY)
    assert_equal(0x4000, WIN32OLE::VariantType::VT_BYREF)
  end

  def test_vt_cy_and_vt_date_constants
    assert_equal(6, WIN32OLE::VariantType::VT_CY)
    assert_equal(7, WIN32OLE::VariantType::VT_DATE)
  end

  def test_variant_is_aliased_to_varianttype
    assert_same(WIN32OLE::VariantType, WIN32OLE::VARIANT)
  end
end

# ole_variant2variant: a Variant converts to a VariantCopy of its VARIANT,
# which the temporary's VariantClear frees -- the Variant keeps its own.
class TestVariantArgumentCopy < Test::Unit::TestCase
  W = WIN32OLE::Win32

  class FakeFn
    def initialize(&body)
      @body = body
    end

    def call(*args)
      @body.call(*args)
    end
  end

  def setup
    @events = []
    events = @events
    @originals = %i[variant_copy_fn variant_clear_fn].to_h { |name| [name, W.method(name)] }
    W.define_singleton_method(:variant_copy_fn) do
      FakeFn.new { |dest, src| events << [:copy, src.dup]; dest[0, W::VARIANT_SIZE] = 'C' * W::VARIANT_SIZE; 0 }
    end
    W.define_singleton_method(:variant_clear_fn) { FakeFn.new { |var| events << [:clear, var.dup]; 0 } }
  end

  def teardown
    @originals.each { |name, m| W.define_singleton_method(name, m) }
  end

  def test_a_variant_converts_to_a_copy_that_to_free_clears
    v = WIN32OLE::Variant.new(42)
    to_free = []
    bytes = WIN32OLE.ruby_value_to_variant_bytes(v, to_free)

    assert_equal('C' * W::VARIANT_SIZE, bytes)
    assert_equal([[:copy, v.instance_variable_get(:@var)]], @events)
    to_free.each(&:call)
    assert_equal([:clear, 'C' * W::VARIANT_SIZE], @events.last)
  end
end

# olevariant_free: a Variant owns what its realvar holds and VariantClears
# it at GC (VT_BYREF: whatever the buffer holds by then, a callee's
# replacement included); value= clears the value it replaces. What the
# conversion allocated besides is freed at once.
class TestVariantOwnership < Test::Unit::TestCase
  W = WIN32OLE::Win32
  V = WIN32OLE::Variant
  VT = WIN32OLE::VariantType
  FakeFn = TestVariantArgumentCopy::FakeFn

  def setup
    @events = []
    events = @events
    bstrs = [0x1000, 0x2000]
    stubs = {
      sys_alloc_string: FakeFn.new { |_| bstrs.shift },
      sys_free_string: FakeFn.new { |bstr| events << [:free, bstr] },
      # a copy holds the original's pointer + 1
      variant_copy_fn: FakeFn.new do |dest, src|
        vt, payload = W.unpack_variant(src)
        dest[0, W::VARIANT_SIZE] = W.pack_variant(vt, W.pack_pointer(W.unpack_pointer(payload) + 1))
        0
      end,
      variant_clear_fn: FakeFn.new do |var|
        bytes = var.is_a?(::String) ? var : Fiddle::Pointer.new(var.to_i)[0, W::VARIANT_SIZE]
        events << [:clear, W.unpack_pointer(W.unpack_variant(bytes)[1])]
        0
      end
    }
    @originals = stubs.keys.to_h { |name| [name, W.method(name)] }
    stubs.each { |name, fn| W.define_singleton_method(name) { fn } }
  end

  def teardown
    @originals.each { |name, m| W.define_singleton_method(name, m) }
  end

  def finalize(variant)
    V.realvar_finalizer(variant.instance_variable_get(:@realvar_finalizer_state)).call
  end

  def test_a_variant_owns_a_copy_and_clears_it_at_gc
    v = V.new('abc')
    assert_equal([[:free, 0x1000]], @events) # the conversion's own BSTR, at once
    finalize(v)
    assert_equal([[:free, 0x1000], [:clear, 0x1001]], @events)
  end

  def test_a_byref_variant_clears_what_its_buffer_holds_at_gc
    v = V.new('abc', VT::VT_BSTR | VT::VT_BYREF)
    replacement = W.pack_variant(VT::VT_BSTR, W.pack_pointer(0x9000)) # a callee's in/out replacement
    v.instance_variable_get(:@realvar)[0, W::VARIANT_SIZE] = replacement
    finalize(v)
    assert_equal([:clear, 0x9000], @events.last)
  end

  def test_value_assignment_clears_the_replaced_value
    v = V.new('abc', VT::VT_BSTR)
    v.value = 'def'
    assert_include(@events, [:clear, 0x1001])
    @events.clear
    finalize(v)
    assert_equal([[:clear, 0x2001]], @events)
  end

  def test_finalizing_twice_clears_once
    v = V.new('abc')
    finalize(v)
    finalize(v)
    assert_equal(1, @events.count { |e| e.first == :clear })
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

# ole_val2olevariantdata / ole_variant2val semantics that need no OLE
# Automation call (no VariantChangeTypeEx, no SysAllocString).
class TestVariantScalarSemantics < Test::Unit::TestCase
  VT = WIN32OLE::VariantType
  V = WIN32OLE::Variant

  def test_nil_with_a_numeric_vartype_is_that_types_zero
    v = V.new(nil, VT::VT_I4)
    assert_equal([0, VT::VT_I4], [v.value, v.vartype])
    assert_equal(0.0, V.new(nil, VT::VT_R8).value)
    assert_equal(false, V.new(nil, VT::VT_BOOL).value)
  end

  def test_null_and_empty_read_back_as_nil
    assert_nil(V::Null.value)
    assert_nil(V.new(nil, VT::VT_EMPTY).value)
  end

  def test_nil_bstr_reads_back_as_an_empty_string
    v = V.new(nil, VT::VT_BSTR)
    assert_equal(['', VT::VT_BSTR], [v.value, v.vartype])
  end

  def test_nil_variant_keeps_vt_variant_and_reads_back_nil
    v = V.new(nil, VT::VT_VARIANT)
    assert_equal([nil, VT::VT_VARIANT], [v.value, v.vartype])
  end

  def test_nil_dispatch_reads_back_nil
    assert_nil(V.new(nil, VT::VT_DISPATCH).value)
  end

  def test_nil_date_is_the_ole_epoch
    assert_equal(Time.new(1899, 12, 30), V.new(nil, VT::VT_DATE).value)
  end

  def test_vt_error_holds_the_scode
    v = V.new(-1, VT::VT_ERROR)
    assert_equal([-1, VT::VT_ERROR], [v.value, v.vartype])
    assert_equal(-2147352572, V::NoParam.value)
  end

  def test_vt_error_rejects_a_non_integer
    assert_raise(WIN32OLE::RuntimeError) { V.new('x', VT::VT_ERROR) }
  end

  def test_a_variant_value_is_converted_from_its_own_variant
    assert_equal(3, V.new(V.new(3, VT::VT_I4), VT::VT_I4).value)
  end

  def test_a_time_becomes_vt_date_and_round_trips
    t = Time.new(2014, 8, 27, 12, 34, 56) + 0.789
    v = V.new(t)
    assert_equal(VT::VT_DATE, v.vartype)
    assert_in_delta(t, v.value, 0.001)
    assert_equal(Time.new(2004, 12, 24, 12, 24, 45), V.new(Time.new(2004, 12, 24, 12, 24, 45), VT::VT_DATE).value)
  end
end

# Reading a VARIANT a caller doesn't own (a Variant's own value, a
# SAFEARRAY element, anything reached through VT_BYREF) must leave it
# intact, as ole_variant2val does: no SysFreeString, an AddRef before
# wrapping an IDispatch.
class TestBorrowedVariantRead < Test::Unit::TestCase
  W = WIN32OLE::Win32
  VT = WIN32OLE::VariantType

  def fake_bstr(str)
    body = str.encode('UTF-16LE').b
    buf = Fiddle::Pointer.malloc(4 + body.bytesize + 2)
    buf[0, 4] = [body.bytesize].pack('L')
    buf[4, body.bytesize + 2] = body + "\x00\x00".b
    buf
  end

  def test_a_borrowed_bstr_is_read_without_being_freed
    buf = fake_bstr('abc')
    var = W.pack_variant(VT::VT_BSTR, W.pack_pointer(buf.to_i + 4))
    assert_equal('abc', WIN32OLE.variant_bytes_to_ruby_value(var, owned: false))
  ensure
    Fiddle.free(buf) if buf
  end

  def test_byref_values_are_dereferenced
    cell = Fiddle::Pointer.malloc(8)
    cell[0, 8] = [42].pack('l') + "\x00".b * 4
    var = W.pack_variant(VT::VT_I4 | VT::VT_BYREF, W.pack_pointer(cell.to_i))
    assert_equal(42, WIN32OLE.variant_bytes_to_ruby_value(var))
  ensure
    Fiddle.free(cell) if cell
  end

  def test_a_null_array_variant_reads_back_nil
    [VT::VT_ARRAY | VT::VT_DISPATCH, VT::VT_ARRAY | VT::VT_DISPATCH | VT::VT_BYREF].each do |vt|
      v = WIN32OLE::Variant.new(nil, vt)
      assert_equal([nil, vt], [v.value, v.vartype])
    end
  end
end

# An OLE DATE counts days from 1899-12-30; for negative dates the
# fraction is the time of day itself, not an offset toward zero.
class TestOleDateConversion < Test::Unit::TestCase
  W = WIN32OLE::Win32

  def test_fractional_days_carry_milliseconds
    t = W.vtdate_to_time(41878.524268391200167) # 2014-08-27 12:34:56.789
    assert_equal('2014-08-27 12:34:56', t.strftime('%Y-%m-%d %H:%M:%S'))
    assert_in_delta(0.789, t.nsec / 1_000_000_000.0, 0.001)
  end

  def test_negative_dates_keep_a_positive_time_of_day
    assert_equal(Time.new(1899, 12, 29, 6, 0, 0), W.vtdate_to_time(-1.25))
    assert_equal(Time.new(100, 1, 1), W.vtdate_to_time(-657434.0))
  end

  def test_time_to_vtdate_is_the_inverse
    assert_equal(0.0, W.time_to_vtdate(Time.new(1899, 12, 30)))
    assert_equal(-1.25, W.time_to_vtdate(Time.new(1899, 12, 29, 6, 0, 0)))
  end
end
end
