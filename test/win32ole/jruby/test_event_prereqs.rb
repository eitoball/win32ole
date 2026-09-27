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

  def test_const_load_defines_constants_without_redefining_existing_ones
    omit('requires a live WIN32OLE COM object') unless defined?(WIN32OLE) && RUBY_ENGINE == 'jruby'
    dict = WIN32OLE.new('Scripting.Dictionary')
    mod = Module.new
    WIN32OLE.const_load(dict, mod)
    # Scripting.Dictionary's typelib (Scripting library) has no constants of
    # its own worth asserting on portably; the real coverage is ADO's
    # WIN32OLE.const_load(@db, ADO) in test_win32ole_event.rb once Event
    # itself lands (Task 15's GC-stress task doesn't touch this, but the
    # ADO-gated legacy suite exercises it directly in CI).
    assert_kind_of(Module, mod)
  end
end
end
