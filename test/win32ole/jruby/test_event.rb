require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/event'

class TestEvent < Test::Unit::TestCase
  W = WIN32OLE::Win32

  def test_new_raises_type_error_for_non_win32ole_argument
    assert_raise(TypeError) { WIN32OLE::Event.new('A') }
  end

  def test_message_loop_is_a_class_method
    assert_respond_to(WIN32OLE::Event, :message_loop)
  end
end
end
