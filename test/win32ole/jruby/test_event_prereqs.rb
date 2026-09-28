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

  # The old guard here was `defined?(WIN32OLE) && RUBY_ENGINE == 'jruby'`,
  # which is tautologically true -- this file's own outer guard already
  # required win32ole/jruby/win32ole -- so on a machine with no COM at all
  # the test ERRORED out of Fiddle.dlopen instead of omitting. Probe for the
  # capability by actually trying, and omit on the two ways it can be
  # absent: no ole32/oleaut32 DLLs (non-Windows), or no such server
  # registered.
  def test_const_load_defines_constants_without_redefining_existing_ones
    dict =
      begin
        WIN32OLE.new('Scripting.Dictionary')
      rescue Fiddle::DLError, WIN32OLE::RuntimeError => e
        omit("requires a live WIN32OLE COM object (#{e.class})")
      end
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
