require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/record'

class TestRecordConstruction < Test::Unit::TestCase
  def test_new_rejects_an_oleobj_that_is_neither_win32ole_nor_typelib
    assert_raise(TypeError) { WIN32OLE::Record.new('Book', Object.new) }
  end
end
end
