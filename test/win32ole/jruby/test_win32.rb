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
