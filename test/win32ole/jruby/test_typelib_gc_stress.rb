begin
  require 'win32ole'
rescue LoadError
end
require 'test/unit'

if defined?(WIN32OLE) && RUBY_ENGINE == 'jruby'
  class TestTypeLibGCStress < Test::Unit::TestCase
    def test_gc_stress_survives_repeated_type_and_typelib_construction
      dict = WIN32OLE.new('Scripting.Dictionary')
      GC.stress = true
      50.times do
        type = dict.ole_type
        tlib = type.ole_typelib
        assert_kind_of(WIN32OLE::Type, type)
        assert_kind_of(WIN32OLE::TypeLib, tlib)
      end
    ensure
      GC.stress = false
    end
  end
end
