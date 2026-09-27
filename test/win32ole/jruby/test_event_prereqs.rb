begin
  require 'win32ole'
rescue LoadError
end
require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/win32ole'

class TestEventPrereqs < Test::Unit::TestCase
  def test_connect_raises_not_implemented_for_non_nil_host
    assert_raise(NotImplementedError) { WIN32OLE.connect('Scripting.Dictionary', 'remotehost') }
  end
end
end
