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

  def advised_event
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@events, [])
    ev.instance_variable_set(:@handler, nil)
    ev.instance_variable_set(:@finalizer_state, { cp_ptr: 0x1234 }) # truthy stand-in for "advised"
    ev
  end

  def test_on_event_registers_a_named_callback
    ev = advised_event
    ev.on_event('Foo') { |*| }
    events = ev.instance_variable_get(:@events)
    assert_equal(1, events.size)
    assert_equal('Foo', events.first[:name])
    assert_equal(false, events.first[:with_outargs])
  end

  def test_on_event_accepts_a_symbol_event_name
    ev = advised_event
    ev.on_event(:Foo) { |*| }
    assert_equal('Foo', ev.instance_variable_get(:@events).first[:name])
  end

  def test_on_event_rejects_non_string_non_symbol_event_name
    ev = advised_event
    assert_raise(TypeError) { ev.on_event(42) { |*| } }
  end

  def test_on_event_with_outargs_sets_the_with_outargs_flag
    ev = advised_event
    ev.on_event_with_outargs('Foo') { |*| }
    assert_equal(true, ev.instance_variable_get(:@events).first[:with_outargs])
  end

  def test_on_event_twice_with_the_same_name_replaces_not_stacks
    ev = advised_event
    first = proc { |*| }
    second = proc { |*| }
    ev.on_event('Foo', &first)
    ev.on_event('Foo', &second)
    events = ev.instance_variable_get(:@events)
    assert_equal(1, events.size)
    assert_same(second, events.first[:proc])
  end

  def test_on_event_twice_with_no_name_replaces_the_catch_all
    ev = advised_event
    first = proc { |*| }
    second = proc { |*| }
    ev.on_event(&first)
    ev.on_event(&second)
    events = ev.instance_variable_get(:@events)
    assert_equal(1, events.size)
    assert_nil(events.first[:name])
    assert_same(second, events.first[:proc])
  end

  def test_on_event_raises_runtime_error_when_not_advised
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@events, [])
    ev.instance_variable_set(:@finalizer_state, nil)
    assert_raise(WIN32OLE::RuntimeError) { ev.on_event('Foo') { |*| } }
  end

  def test_off_event_removes_only_the_named_entry
    ev = advised_event
    ev.on_event('Foo') { |*| }
    ev.on_event('Bar') { |*| }
    ev.off_event('Foo')
    events = ev.instance_variable_get(:@events)
    assert_equal(1, events.size)
    assert_equal('Bar', events.first[:name])
  end

  def test_off_event_with_no_args_removes_only_the_catch_all
    ev = advised_event
    ev.on_event('Foo') { |*| }
    ev.on_event { |*| }
    ev.off_event
    events = ev.instance_variable_get(:@events)
    assert_equal(1, events.size)
    assert_equal('Foo', events.first[:name])
  end

  def test_off_event_accepts_a_symbol
    ev = advised_event
    ev.on_event('Foo') { |*| }
    ev.off_event(:Foo)
    assert_empty(ev.instance_variable_get(:@events))
  end

  def test_handler_accessor_round_trips
    ev = WIN32OLE::Event.allocate
    handler = Object.new
    ev.handler = handler
    assert_same(handler, ev.handler)
  end

  def test_find_iid_by_name_is_private
    assert(WIN32OLE::Event.private_method_defined?(:find_iid_by_name))
  end

  def test_find_iid_by_guid_is_private
    assert(WIN32OLE::Event.private_method_defined?(:find_iid_by_guid))
  end

  def test_resolve_event_source_is_private
    assert(WIN32OLE::Event.private_method_defined?(:resolve_event_source))
  end
end
end
