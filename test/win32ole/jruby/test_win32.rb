# test/win32ole/jruby/test_win32.rb
require 'test/unit'

if RUBY_ENGINE == 'jruby'
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

  def test_pack_variant_rejects_oversized_payload
    # After generalization, only reject payloads that exceed VARIANT_SIZE - 8
    assert_raise(ArgumentError) { W.pack_variant(W::VT_I4, "\x00".b * (W::VARIANT_SIZE - 7)) }
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

  def test_dispatch_plan_with_args_uses_method_and_propertyget
    plan = W.dispatch_plan('Add', ['k', 'v'])
    assert_equal(W::DISPATCH_METHOD | W::DISPATCH_PROPERTYGET, plan[:wflags])
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

  def test_vtable_function_rejects_small_integer_addresses
    # A real crash was found via CI: passing a small-but-valid Integer (not
    # a real pointer) through to the first native dereference caused a JVM
    # EXCEPTION_ACCESS_VIOLATION rather than a catchable Ruby exception.
    err = assert_raise(TypeError) { W.vtable_function(100, 2, [], Fiddle::TYPE_VOID) }
    assert_match(/100/, err.message)
  end

  def test_vtable_function_rejects_non_integer_addresses
    assert_raise(TypeError) { W.vtable_function('not a pointer', 2, [], Fiddle::TYPE_VOID) }
    assert_raise(TypeError) { W.vtable_function(nil, 2, [], Fiddle::TYPE_VOID) }
  end

  def test_vtable_address_shares_the_same_guard
    assert_raise(TypeError) { W.vtable_address(1) }
  end

  def test_vtable_address_reads_the_first_pointer_sized_field
    fake_vtable_addr = 0x123456
    object_buf = [fake_vtable_addr].pack(W::PACK_PTR)
    object_ptr = Fiddle::Pointer.to_ptr(object_buf)
    assert_equal(fake_vtable_addr, W.vtable_address(object_ptr.to_i))
  end

  def test_pack_variant_accepts_a_variable_length_body
    bytes = W.pack_variant(W::VT_RECORD, "\x01\x02")
    assert_equal(W::VARIANT_SIZE, bytes.bytesize)
  end

  def test_pack_variant_zero_pads_a_short_body
    bytes = W.pack_variant(W::VT_I4, W.pack_i4(7))
    assert_equal("\x00".b * (W::VARIANT_SIZE - 12), bytes[12, W::VARIANT_SIZE - 12])
  end

  def test_pack_variant_rejects_a_body_longer_than_the_slot
    assert_raise(ArgumentError) { W.pack_variant(W::VT_I4, "\x00".b * (W::VARIANT_SIZE - 7)) }
  end

  def test_unpack_variant_reads_the_requested_body_size
    body = [1, 2].pack(W::PACK_PTR * 2)
    bytes = W.pack_variant(W::VT_RECORD, body)
    vt, read_body = W.unpack_variant(bytes, body_size: body.bytesize)
    assert_equal(W::VT_RECORD, vt)
    assert_equal(body, read_body)
  end

  def test_unpack_variant_default_body_size_is_unchanged
    bytes = W.pack_variant(W::VT_I4, W.pack_i4(42))
    vt, payload = W.unpack_variant(bytes)
    assert_equal(W::VT_I4, vt)
    assert_equal(42, W.unpack_i4(payload))
  end

  def test_scalar_family_round_trips
    assert_equal(-5, W.unpack_i1(W.pack_i1(-5)))
    assert_equal(200, W.unpack_ui1(W.pack_ui1(200)))
    assert_equal(-1000, W.unpack_i2(W.pack_i2(-1000)))
    assert_equal(40_000, W.unpack_ui2(W.pack_ui2(40_000)))
    assert_equal(4_000_000_000, W.unpack_ui4(W.pack_ui4(4_000_000_000)))
    assert_equal(-7, W.unpack_int(W.pack_int(-7)))
    assert_equal(7, W.unpack_uint(W.pack_uint(7)))
    assert_equal(18_000_000_000, W.unpack_ui8(W.pack_ui8(18_000_000_000)))
    assert_in_delta(1.5, W.unpack_r4(W.pack_r4(1.5)), 0.0001)
    assert_equal(-2_147_024_809, W.unpack_error(W.pack_error(-2_147_024_809)))
  end

  def test_scalar_family_payloads_are_eight_bytes_zero_padded
    assert_equal(8, W.pack_i1(1).bytesize)
    assert_equal(8, W.pack_ui8(1).bytesize) # the one already-8-byte-wide case: no padding needed
    assert_equal("\x00".b * 7, W.pack_i1(1)[1, 7])
  end

  def test_new_vt_constants_do_not_collide_with_existing_ones
    existing = [W::VT_EMPTY, W::VT_I4, W::VT_R8, W::VT_BSTR, W::VT_DISPATCH, W::VT_BOOL, W::VT_UNKNOWN, W::VT_I8]
    new_scalars = [W::VT_NULL, W::VT_I2, W::VT_R4, W::VT_ERROR, W::VT_VARIANT,
                   W::VT_I1, W::VT_UI1, W::VT_UI2, W::VT_UI4, W::VT_UI8, W::VT_INT, W::VT_UINT, W::VT_RECORD]
    assert_empty(existing & new_scalars)
    assert_equal(0x2000, W::VT_ARRAY)
    assert_equal(0x4000, W::VT_BYREF)
  end
end
end
