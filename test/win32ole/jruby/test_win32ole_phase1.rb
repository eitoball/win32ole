begin
  require 'win32ole'
rescue LoadError
end
require 'test/unit'

if defined?(WIN32OLE) && RUBY_ENGINE == 'jruby'
  class TestWin32OLEPhase1 < Test::Unit::TestCase
    def setup
      @dict = WIN32OLE.new('Scripting.Dictionary')
    end

    def test_new_by_progid
      assert_kind_of(WIN32OLE, @dict)
    end

    def test_new_unknown_progid_raises
      exc = assert_raise(WIN32OLE::RuntimeError) { WIN32OLE.new('NonExistProgID999') }
      assert_match(/^unknown OLE server: `NonExistProgID999'/, exc.message) #`
    end

    def test_two_arg_method_and_one_arg_method
      @dict.add('a', 1000)
      assert_equal(1000, @dict.item('a'))
    end

    def test_property_put_and_get
      @dict.compareMode = 1
      @dict.add('one', 1)
      assert_equal(1, @dict.item('ONE'))
    end

    def test_raise_message_on_wrong_arg_count
      exc = assert_raise(WIN32OLE::RuntimeError) { @dict.add }
      assert_match(/^\(in OLE method `add': \)/, exc.message) #`
    end

    def test_raise_message_on_bad_property_put
      exc = assert_raise(WIN32OLE::RuntimeError) { @dict.compareMode = -1 }
      assert_match(/^\(in setting property `compareMode': \)/, exc.message) #`
    end

    def test_no_method_error
      exc = assert_raise(NoMethodError) { @dict.non_exist_method }
      assert_match(/non_exist_method/, exc.message)
      assert_kind_of(WIN32OLE, exc.receiver)
    end

    def test_dispatch_return_value_wraps_as_win32ole
      fso = WIN32OLE.new('Scripting.FileSystemObject')
      drives = fso.Drives
      assert_kind_of(WIN32OLE, drives)
    end

    def test_string_bool_and_nil_marshaling
      fso = WIN32OLE.new('Scripting.FileSystemObject')
      assert_equal('foo', fso.GetBaseName('C:\\Temp\\foo.txt'))
      assert_equal(true, fso.FileExists('C:\\Windows\\System32\\drivers\\etc\\hosts'))
    end

    def test_gc_stress_survives_repeated_invoke
      GC.stress = true
      100.times do |i|
        @dict.add("key#{i}", i)
        assert_equal(i, @dict.item("key#{i}"))
      end
    ensure
      GC.stress = false
    end
  end
end
