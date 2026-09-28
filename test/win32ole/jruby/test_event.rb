require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/event'

class TestEvent < Test::Unit::TestCase
  W = WIN32OLE::Win32
  E_NOINTERFACE = -2147467262
  DISP_E_BADINDEX = -2147352565

  def build_test_sink(source_iid_bytes = ("\x01" * 16).b, event_typeinfo_ptr = 0)
    ev = WIN32OLE::Event.allocate
    sink_addr, vtable_addr, closures = ev.send(:build_sink, source_iid_bytes, event_typeinfo_ptr)
    [sink_addr, vtable_addr, closures]
  end

  def test_query_interface_closure_returns_sink_for_known_and_source_iids
    source_iid = ("\x01" * 16).b
    sink_addr, vtable_addr, closures = build_test_sink(source_iid)
    qi_fn = Fiddle::Function.new(closures[0], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG)
    ppv = ("\xFF" * W::PTR_SIZE).b

    assert_equal(0, qi_fn.call(sink_addr, W::IID_IUNKNOWN, ppv))
    assert_equal(sink_addr, ppv.unpack1(W::PACK_PTR))

    assert_equal(0, qi_fn.call(sink_addr, W::IID_IDISPATCH, ppv))
    assert_equal(sink_addr, ppv.unpack1(W::PACK_PTR))

    assert_equal(0, qi_fn.call(sink_addr, source_iid, ppv))
    assert_equal(sink_addr, ppv.unpack1(W::PACK_PTR))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

  def test_query_interface_closure_rejects_unknown_iid
    sink_addr, vtable_addr, closures = build_test_sink
    qi_fn = Fiddle::Function.new(closures[0], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG)
    ppv = ("\xFF" * W::PTR_SIZE).b

    assert_equal(E_NOINTERFACE, qi_fn.call(sink_addr, ("\xFE" * 16).b, ppv))
    assert_equal(0, ppv.unpack1(W::PACK_PTR))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

  def test_add_ref_and_release_closures_share_a_refcount
    sink_addr, vtable_addr, closures = build_test_sink
    add_ref_fn = Fiddle::Function.new(closures[1], [W::VOIDP], W::DWORD)
    release_fn = Fiddle::Function.new(closures[2], [W::VOIDP], W::DWORD)

    assert_equal(1, add_ref_fn.call(sink_addr))
    assert_equal(2, add_ref_fn.call(sink_addr))
    assert_equal(1, release_fn.call(sink_addr))
    assert_equal(0, release_fn.call(sink_addr))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

  def test_get_type_info_count_closure_always_reports_zero
    sink_addr, vtable_addr, closures = build_test_sink
    fn = Fiddle::Function.new(closures[3], [W::VOIDP, W::VOIDP], W::LONG)
    pct = ("\xFF" * 4).b

    assert_equal(0, fn.call(sink_addr, pct))
    assert_equal(0, pct.unpack1('L'))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

  def test_get_type_info_closure_always_fails_with_bad_index
    sink_addr, vtable_addr, closures = build_test_sink
    fn = Fiddle::Function.new(closures[4], [W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::LONG)
    ppti = ("\xFF" * W::PTR_SIZE).b

    assert_equal(DISP_E_BADINDEX, fn.call(sink_addr, 0, 0, ppti))
    assert_equal(0, ppti.unpack1(W::PACK_PTR))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

  def test_build_sink_wires_all_seven_vtable_slots_to_the_closures
    sink_addr, vtable_addr, closures = build_test_sink
    assert_equal(7, closures.size)
    closures.each_with_index do |closure, i|
      slot = Fiddle::Pointer.new(vtable_addr)[i * W::PTR_SIZE, W::PTR_SIZE].unpack1(W::PACK_PTR)
      assert_equal(closure.to_i, slot)
    end
    assert_equal(vtable_addr, Fiddle::Pointer.new(sink_addr)[0, W::PTR_SIZE].unpack1(W::PACK_PTR))
  ensure
    Fiddle.free(vtable_addr) if vtable_addr
    Fiddle.free(sink_addr) if sink_addr
  end

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
