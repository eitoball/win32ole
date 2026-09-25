require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/win32ole'
require 'win32ole/jruby/typelib'
require 'win32ole/jruby/record'

class TestRecordConstruction < Test::Unit::TestCase
  def test_new_rejects_an_oleobj_that_is_neither_win32ole_nor_typelib
    assert_raise(TypeError) { WIN32OLE::Record.new('Book', Object.new) }
  end
end

class TestRecordFieldAccess < Test::Unit::TestCase
  def setup
    @record = WIN32OLE::Record.allocate
    @record.instance_variable_set(:@typename, 'Book')
    @record.instance_variable_set(:@fields, { 'title' => 'The Ruby Book', 'cost' => 20 })
  end

  def test_typename
    assert_equal('Book', @record.typename)
  end

  def test_to_h_returns_the_fields_hash
    assert_equal({ 'title' => 'The Ruby Book', 'cost' => 20 }, @record.to_h)
  end

  def test_to_h_is_not_a_defensive_copy
    @record.to_h['cost'] = 99
    assert_equal(99, @record.to_h['cost'])
  end

  def test_method_missing_getter
    assert_equal('The Ruby Book', @record.title)
  end

  def test_method_missing_setter
    @record.title = 'Ruby'
    assert_equal('Ruby', @record.title)
  end

  def test_method_missing_getter_raises_key_error_for_unknown_field
    assert_raise(KeyError) { @record.no_such_field }
  end

  def test_method_missing_setter_raises_key_error_for_unknown_field
    assert_raise(KeyError) { @record.no_such_field = 1 }
  end

  def test_ole_instance_variable_get
    assert_equal(20, @record.ole_instance_variable_get(:cost))
  end

  def test_ole_instance_variable_set
    @record.ole_instance_variable_set(:cost, 30)
    assert_equal(30, @record.ole_instance_variable_get(:cost))
  end

  def test_inspect
    assert_equal('#<WIN32OLE::Record:Book>', @record.inspect)
  end
end
end
