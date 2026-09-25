begin
  require 'win32ole'
rescue LoadError
end
require 'test/unit'

if defined?(WIN32OLE) && RUBY_ENGINE == 'jruby'
  class TestRecordVariantGCStress < Test::Unit::TestCase
    def test_gc_stress_survives_repeated_variant_array_construction
      dict = WIN32OLE.new('Scripting.Dictionary')
      dict.add('a', 1)
      dict.add('b', 2)
      GC.stress = true
      50.times do
        keys = dict.Keys
        items = dict.Items
        assert_kind_of(Array, keys)
        assert_kind_of(Array, items)
      end
    ensure
      GC.stress = false
    end
  end
end
