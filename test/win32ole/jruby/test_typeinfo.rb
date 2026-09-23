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
    assert_equal(PTR64 ? 88 : 76, TI::TYPEATTR.size)
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
end
