require 'test/unit'

class TestJRubyRequire < Test::Unit::TestCase
  def test_require_win32ole_defines_the_class_on_jruby
    omit('JRuby-only') unless RUBY_ENGINE == 'jruby'

    require 'win32ole'
    assert(defined?(WIN32OLE), 'WIN32OLE should be defined after require "win32ole" on JRuby')
    assert_kind_of(Class, WIN32OLE)
  end

  def test_methods_override_does_not_raise_before_ole_methods_exists
    omit('JRuby-only') unless RUBY_ENGINE == 'jruby'
    require 'win32ole'

    obj = WIN32OLE.allocate
    assert_nothing_raised { obj.methods }
  end
end
