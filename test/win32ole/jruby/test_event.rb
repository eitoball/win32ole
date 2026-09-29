require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'stringio'
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
    Fiddle.free(Fiddle::Pointer.new(vtable_addr)) if vtable_addr
    Fiddle.free(Fiddle::Pointer.new(sink_addr)) if sink_addr
  end

  def test_query_interface_closure_rejects_unknown_iid
    sink_addr, vtable_addr, closures = build_test_sink
    qi_fn = Fiddle::Function.new(closures[0], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG)
    ppv = ("\xFF" * W::PTR_SIZE).b

    assert_equal(E_NOINTERFACE, qi_fn.call(sink_addr, ("\xFE" * 16).b, ppv))
    assert_equal(0, ppv.unpack1(W::PACK_PTR))
  ensure
    Fiddle.free(Fiddle::Pointer.new(vtable_addr)) if vtable_addr
    Fiddle.free(Fiddle::Pointer.new(sink_addr)) if sink_addr
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
    Fiddle.free(Fiddle::Pointer.new(vtable_addr)) if vtable_addr
    Fiddle.free(Fiddle::Pointer.new(sink_addr)) if sink_addr
  end

  def test_get_type_info_count_closure_always_reports_zero
    sink_addr, vtable_addr, closures = build_test_sink
    fn = Fiddle::Function.new(closures[3], [W::VOIDP, W::VOIDP], W::LONG)
    pct = ("\xFF" * 4).b

    assert_equal(0, fn.call(sink_addr, pct))
    assert_equal(0, pct.unpack1('L'))
  ensure
    Fiddle.free(Fiddle::Pointer.new(vtable_addr)) if vtable_addr
    Fiddle.free(Fiddle::Pointer.new(sink_addr)) if sink_addr
  end

  def test_get_type_info_closure_always_fails_with_bad_index
    sink_addr, vtable_addr, closures = build_test_sink
    fn = Fiddle::Function.new(closures[4], [W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::LONG)
    ppti = ("\xFF" * W::PTR_SIZE).b

    assert_equal(DISP_E_BADINDEX, fn.call(sink_addr, 0, 0, ppti))
    assert_equal(0, ppti.unpack1(W::PACK_PTR))
  ensure
    Fiddle.free(Fiddle::Pointer.new(vtable_addr)) if vtable_addr
    Fiddle.free(Fiddle::Pointer.new(sink_addr)) if sink_addr
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
    Fiddle.free(Fiddle::Pointer.new(vtable_addr)) if vtable_addr
    Fiddle.free(Fiddle::Pointer.new(sink_addr)) if sink_addr
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

  def test_find_event_entry_prefers_a_named_match_over_the_catch_all
    ev = WIN32OLE::Event.allocate
    fallback = { name: nil, proc: proc { :fallback }, with_outargs: false }
    named = { name: 'Foo', proc: proc { :named }, with_outargs: false }
    ev.instance_variable_set(:@events, [fallback, named])
    entry, is_default = ev.send(:find_event_entry, 'Foo')
    assert_same(named, entry)
    assert_equal(false, is_default)
  end

  def test_find_event_entry_falls_back_to_the_catch_all
    ev = WIN32OLE::Event.allocate
    fallback = { name: nil, proc: proc { :fallback }, with_outargs: false }
    ev.instance_variable_set(:@events, [fallback])
    entry, is_default = ev.send(:find_event_entry, 'Bar')
    assert_same(fallback, entry)
    assert_equal(true, is_default)
  end

  def test_find_event_entry_returns_nil_when_nothing_matches
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@events, [])
    entry, is_default = ev.send(:find_event_entry, 'Bar')
    assert_nil(entry)
    assert_equal(false, is_default)
  end

  def test_handle_invoke_writes_exception_message_to_stderr_and_does_not_raise
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@events, [{ name: nil, proc: proc { raise 'boom' }, with_outargs: false }])
    ev.instance_variable_set(:@handler, nil)
    ev.instance_variable_set(:@event_typeinfo_ptr, 0)
    ev.define_singleton_method(:resolve_event_name) { |_dispid| 'Whatever' }

    dispparams = [0, 0, 0, 0].pack("#{W::PACK_PTR}#{W::PACK_PTR}LL")
    dispparams_ptr = Fiddle::Pointer.to_ptr(dispparams)

    err = capture_stderr { ev.send(:handle_invoke, 1, dispparams_ptr, nil) }
    assert_match(/boom/, err)
  end

  def test_unadvise_on_a_never_advised_event_is_a_safe_noop
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@finalizer_state, nil)
    assert_nil(ev.unadvise)
  end

  # Builds a fake COM object: a malloc'd single-pointer object whose vtable
  # has `slots` pointer-sized entries, NULL except for the given
  # index => closure pairs. Returns [object_ptr, vtable_ptr]; the caller owns
  # (and must free) both, and must keep the closures themselves referenced
  # for as long as the object is used -- GC'ing a Fiddle::Closure frees its
  # native trampoline. Same technique as test_win32.rb's
  # test_query_interface_* fakes, generalised to more than one slot.
  def build_fake_com_object(slots, closures_by_index)
    vtable = Fiddle::Pointer.malloc(W::PTR_SIZE * slots)
    slots.times { |i| vtable[i * W::PTR_SIZE, W::PTR_SIZE] = [0].pack(W::PACK_PTR) }
    closures_by_index.each do |index, closure|
      vtable[index * W::PTR_SIZE, W::PTR_SIZE] = [closure.to_i].pack(W::PACK_PTR)
    end
    obj = Fiddle::Pointer.malloc(W::PTR_SIZE)
    obj[0, W::PTR_SIZE] = [vtable.to_i].pack(W::PACK_PTR)
    [obj, vtable]
  end

  def counting_release_closure(counter)
    Fiddle::Closure::BlockCaller.new(W::DWORD, [W::VOIDP], W::STDCALL) do |_this|
      counter[0] += 1
      0
    end
  end

  # #unadvise drains the Windows message queue before tearing the sink down;
  # user32 is not loadable off Windows, so stub that one call out.
  def with_stubbed_message_pump
    original = W.method(:pump_windows_messages)
    pumped = [0]
    W.define_singleton_method(:pump_windows_messages) { pumped[0] += 1; nil }
    yield pumped
  ensure
    W.define_singleton_method(:pump_windows_messages, original) if original
  end

  # Exercises the REAL teardown path (native Unadvise/Release/Fiddle.free)
  # against fake-but-real-closure-backed vtables, then asserts the second
  # call is a no-op. The previous version of this test seeded cp_ptr: nil, so
  # the first call short-circuited at the guard and no native call ever ran.
  def test_unadvise_is_idempotent
    unadvise_calls = [0]
    cp_releases = [0]
    ti_releases = [0]
    received_cookies = []
    unadvise_closure = Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::DWORD], W::STDCALL) do |_this, cookie|
      unadvise_calls[0] += 1
      received_cookies << cookie
      0
    end
    cp_release = counting_release_closure(cp_releases)
    ti_release = counting_release_closure(ti_releases)

    cp_obj, cp_vtable = build_fake_com_object(7, 2 => cp_release, 6 => unadvise_closure)
    ti_obj, ti_vtable = build_fake_com_object(3, 2 => ti_release)
    sink_buf = Fiddle::Pointer.malloc(W::PTR_SIZE)
    vtable_buf = Fiddle::Pointer.malloc(W::PTR_SIZE)

    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@finalizer_state, {
                               cp_ptr: cp_obj.to_i, cookie: 0x2A, ti_ptr: ti_obj.to_i,
                               sink_addr: sink_buf.to_i, vtable_addr: vtable_buf.to_i
                             })
    ev.instance_variable_set(:@sink_closures, [])

    with_stubbed_message_pump do |pumped|
      assert_nil(ev.unadvise)
      assert_nil(ev.unadvise) # must not raise, and must not Unadvise/Release/free twice
      assert_equal(1, pumped[0])
    end

    assert_equal(1, unadvise_calls[0])
    assert_equal([0x2A], received_cookies)
    assert_equal(1, cp_releases[0])
    assert_equal(1, ti_releases[0])
    state = ev.instance_variable_get(:@finalizer_state)
    assert_nil(state[:cp_ptr])
    assert_nil(state[:ti_ptr])
    assert_nil(state[:sink_addr]) # the two malloc'd buffers were freed exactly once
    assert_nil(state[:vtable_addr])
  ensure
    Fiddle.free(cp_vtable) if cp_vtable
    Fiddle.free(cp_obj) if cp_obj
    Fiddle.free(ti_vtable) if ti_vtable
    Fiddle.free(ti_obj) if ti_obj
  end

  # The one test that covers EVERY vtable_function call site on the
  # advise/unadvise path at once: a fake IDispatch whose QueryInterface hands
  # back a fake IConnectionPointContainer, whose FindConnectionPoint hands
  # back a fake IConnectionPoint, all backed by real
  # Fiddle::Closure::BlockCaller trampolines. A wrong arg_types arity on any
  # of those calls raises ArgumentError before a single byte is exchanged, so
  # this fails loudly against arity bugs without needing Windows.
  def test_advise_on_event_unadvise_round_trip_against_fake_com_objects
    cookie_written = 0xABCD
    ti_releases = [0]
    container_releases = [0]
    cp_releases = [0]
    qi_iids = []
    find_cp_iids = []
    advised_sinks = []
    unadvise_calls = [0]

    # Every closure below is bound to a local that stays in scope for the
    # whole test: once its address is written into a malloc'd vtable slot, the
    # GC can no longer see that raw Integer as a reference, so an
    # inline-only closure can be collected -- freeing its native trampoline
    # -- while the test still expects to call through the slot. See
    # build_fake_com_object's own contract.
    ti_release = counting_release_closure(ti_releases)
    ti_obj, ti_vtable = build_fake_com_object(3, 2 => ti_release)

    unadvise_closure = Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::DWORD], W::STDCALL) do |_this, _cookie|
      unadvise_calls[0] += 1
      0
    end
    advise_closure = Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL) do |_this, sink, pcookie|
      advised_sinks << sink.to_i
      pcookie[0, 4] = [cookie_written].pack('L')
      0
    end
    cp_release = counting_release_closure(cp_releases)
    cp_obj, cp_vtable = build_fake_com_object(
      7, 2 => cp_release, 5 => advise_closure, 6 => unadvise_closure
    )

    find_cp_closure = Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL) do |_this, riid, ppcp|
      find_cp_iids << riid[0, 16]
      ppcp[0, W::PTR_SIZE] = [cp_obj.to_i].pack(W::PACK_PTR)
      0
    end
    container_release = counting_release_closure(container_releases)
    container_obj, container_vtable = build_fake_com_object(
      5, 2 => container_release, 4 => find_cp_closure
    )

    qi_closure = Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL) do |_this, riid, ppv|
      iid = riid[0, 16]
      qi_iids << iid
      if iid == W::IID_ICONNECTIONPOINTCONTAINER
        ppv[0, W::PTR_SIZE] = [container_obj.to_i].pack(W::PACK_PTR)
        0
      else
        ppv[0, W::PTR_SIZE] = [0].pack(W::PACK_PTR)
        E_NOINTERFACE
      end
    end
    dispatch_obj, dispatch_vtable = build_fake_com_object(1, 0 => qi_closure)

    # Regression guard for the closure-liveness contract above: if any of the
    # closures wired into the vtables were only reachable through their raw
    # address in malloc'd memory, this collects them and the first call
    # through that slot segfaults the process rather than failing an
    # assertion. Cheap, and it fails loudly the moment someone inlines a
    # closure back into build_fake_com_object.
    GC.start

    source_iid = ("\x07" * 16).b
    ole = Object.new
    ole.define_singleton_method(:dispatch_ptr) { dispatch_obj.to_i }

    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@events, [])
    ev.instance_variable_set(:@handler, nil)
    ev.instance_variable_set(:@finalizer_state, nil)
    ev.instance_variable_set(:@sink_closures, nil)
    ev.instance_variable_set(:@event_typeinfo_ptr, nil)
    ev.instance_variable_set(:@ole, nil)
    ev.define_singleton_method(:resolve_event_source) { |_ole, _itf| [source_iid, ti_obj.to_i] }

    ev.send(:advise, ole, nil)

    assert_equal([W::IID_ICONNECTIONPOINTCONTAINER], qi_iids)
    assert_equal([source_iid], find_cp_iids)
    assert_equal(1, container_releases[0]) # released right after FindConnectionPoint
    state = ev.instance_variable_get(:@finalizer_state)
    assert_equal(cp_obj.to_i, state[:cp_ptr])
    assert_equal(cookie_written, state[:cookie])
    assert_equal(ti_obj.to_i, state[:ti_ptr])
    assert_equal([state[:sink_addr]], advised_sinks) # Advise really got our sink
    assert_equal(7, ev.instance_variable_get(:@sink_closures).size)
    assert_equal(ti_obj.to_i, ev.instance_variable_get(:@event_typeinfo_ptr))
    assert_same(ole, ev.instance_variable_get(:@ole)) # keeps the source object alive

    ev.on_event('Foo') { |*| }
    assert_equal('Foo', ev.instance_variable_get(:@events).first[:name])

    GC.start # again right before the teardown calls -- the exact crash path
    with_stubbed_message_pump do
      assert_nil(ev.unadvise)
      assert_nil(ev.unadvise)
    end

    assert_equal(1, unadvise_calls[0])
    assert_equal(1, cp_releases[0])
    assert_equal(1, ti_releases[0])
    assert_nil(ev.instance_variable_get(:@event_typeinfo_ptr))
    assert_nil(ev.instance_variable_get(:@sink_closures))
  ensure
    [dispatch_vtable, dispatch_obj, container_vtable, container_obj,
     cp_vtable, cp_obj, ti_vtable, ti_obj].each { |p| Fiddle.free(p) if p }
  end

  def fake_byref_variant(vt, ref_bytesize)
    ref_buf = Fiddle::Pointer.malloc(ref_bytesize)
    var = W.pack_variant(vt | W::VT_BYREF, W.pack_pointer(ref_buf.to_i))
    var_ptr = Fiddle::Pointer.to_ptr(var)
    [var_ptr, ref_buf]
  end

  def test_write_byref_variant_writes_a_bool
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_BOOL, 2)
    ev.send(:write_byref_variant, var_ptr, true)
    assert_equal(-1, ref_buf[0, 2].unpack1('s'))
  ensure
    Fiddle.free(ref_buf) if ref_buf
  end

  def test_write_byref_variant_writes_an_i4
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_I4, 4)
    ev.send(:write_byref_variant, var_ptr, 42)
    assert_equal(42, ref_buf[0, 4].unpack1('l'))
  ensure
    Fiddle.free(ref_buf) if ref_buf
  end

  def test_write_byref_variant_writes_an_r8
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_R8, 8)
    ev.send(:write_byref_variant, var_ptr, 1.5)
    assert_in_delta(1.5, ref_buf[0, 8].unpack1('d'), 0.0001)
  ensure
    Fiddle.free(ref_buf) if ref_buf
  end

  def test_write_byref_variant_ignores_a_non_byref_variant
    ev = WIN32OLE::Event.allocate
    var = W.pack_variant(W::VT_I4, W.pack_i4(0))
    var_ptr = Fiddle::Pointer.to_ptr(var)
    assert_nil(ev.send(:write_byref_variant, var_ptr, 99)) # must not raise / must not dereference garbage
  end

  def test_write_byref_variant_ignores_a_type_mismatched_value
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_I4, 4)
    ref_buf[0, 4] = [7].pack('l')
    ev.send(:write_byref_variant, var_ptr, [1, 2, 3]) # Array has no matching case -- silent no-op, matches C
    assert_equal(7, ref_buf[0, 4].unpack1('l'))
  ensure
    Fiddle.free(ref_buf) if ref_buf
  end

  def test_write_array_outargs_writes_positionally_and_stops_at_cargs
    ev = WIN32OLE::Event.allocate
    rgvarg = Fiddle::Pointer.malloc(W::VARIANT_SIZE * 2)
    ref0 = Fiddle::Pointer.malloc(4)
    ref1 = Fiddle::Pointer.malloc(4)
    rgvarg[1 * W::VARIANT_SIZE, W::VARIANT_SIZE] = W.pack_variant(W::VT_I4 | W::VT_BYREF, W.pack_pointer(ref0.to_i)) # arg 0
    rgvarg[0 * W::VARIANT_SIZE, W::VARIANT_SIZE] = W.pack_variant(W::VT_I4 | W::VT_BYREF, W.pack_pointer(ref1.to_i)) # arg 1

    ev.send(:write_array_outargs, [11, 22], 2, rgvarg.to_i)

    assert_equal(11, ref0[0, 4].unpack1('l'))
    assert_equal(22, ref1[0, 4].unpack1('l'))
  ensure
    Fiddle.free(rgvarg) if rgvarg
    Fiddle.free(ref0) if ref0
    Fiddle.free(ref1) if ref1
  end

  def test_write_byref_variant_ignores_a_null_byref_pointer
    ev = WIN32OLE::Event.allocate
    var = W.pack_variant(W::VT_I4 | W::VT_BYREF, W.pack_pointer(0))
    var_ptr = Fiddle::Pointer.to_ptr(var)
    assert_nil(ev.send(:write_byref_variant, var_ptr, 42)) # ref_addr is null -- must not dereference
  end

  def test_write_byref_variant_writes_a_ui1
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_UI1, 1)
    ev.send(:write_byref_variant, var_ptr, 200)
    assert_equal(200, ref_buf[0, 1].unpack1('C'))
  ensure
    Fiddle.free(ref_buf) if ref_buf
  end

  def test_write_byref_variant_writes_a_false_bool
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_BOOL, 2)
    ref_buf[0, 2] = [-1].pack('s') # pre-seed with "true" so a no-op would be caught
    ev.send(:write_byref_variant, var_ptr, false)
    assert_equal(0, ref_buf[0, 2].unpack1('s'))
  ensure
    Fiddle.free(ref_buf) if ref_buf
  end

  def test_write_byref_variant_writes_a_float_into_r4
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_R4, 4)
    ev.send(:write_byref_variant, var_ptr, 2.5)
    assert_in_delta(2.5, ref_buf[0, 4].unpack1('f'), 0.0001)
  ensure
    Fiddle.free(ref_buf) if ref_buf
  end

  # SysAllocString itself lives in oleaut32.dll, which is not loadable on
  # this (non-Windows) machine -- see W.sys_alloc_string. We stub it here
  # purely to exercise write_byref_variant's own String/VT_BSTR branch
  # (that it calls sys_alloc_string with a UTF-16LE-encoded copy of the
  # value and writes the returned pointer into the ref slot) without
  # depending on the real DLL.
  def test_write_byref_variant_writes_a_bstr
    ev = WIN32OLE::Event.allocate
    var_ptr, ref_buf = fake_byref_variant(W::VT_BSTR, W::PTR_SIZE)
    original_sys_alloc_string = W.method(:sys_alloc_string)
    fake_bstr_addr = 0x1234
    received_wstr = nil
    fake_fn = Object.new
    fake_fn.define_singleton_method(:call) do |wstr_bytes|
      received_wstr = wstr_bytes
      fake_bstr_addr
    end
    W.define_singleton_method(:sys_alloc_string) { fake_fn }

    ev.send(:write_byref_variant, var_ptr, 'hello')

    assert_equal(fake_bstr_addr, ref_buf[0, W::PTR_SIZE].unpack1(W::PACK_PTR))
    assert_equal(W.wstr('hello'), received_wstr)
  ensure
    W.define_singleton_method(:sys_alloc_string, original_sys_alloc_string) if original_sys_alloc_string
    Fiddle.free(ref_buf) if ref_buf
  end

  def test_write_array_outargs_stops_when_the_array_is_longer_than_cargs
    ev = WIN32OLE::Event.allocate
    rgvarg = Fiddle::Pointer.malloc(W::VARIANT_SIZE * 2)
    ref0 = Fiddle::Pointer.malloc(4)
    ref1 = Fiddle::Pointer.malloc(4)
    rgvarg[1 * W::VARIANT_SIZE, W::VARIANT_SIZE] = W.pack_variant(W::VT_I4 | W::VT_BYREF, W.pack_pointer(ref0.to_i)) # arg 0
    rgvarg[0 * W::VARIANT_SIZE, W::VARIANT_SIZE] = W.pack_variant(W::VT_I4 | W::VT_BYREF, W.pack_pointer(ref1.to_i)) # arg 1

    # 33 has no matching rgvarg slot (cargs is 2) -- must not overrun rgvarg
    ev.send(:write_array_outargs, [11, 22, 33], 2, rgvarg.to_i)

    assert_equal(11, ref0[0, 4].unpack1('l'))
    assert_equal(22, ref1[0, 4].unpack1('l'))
  ensure
    Fiddle.free(rgvarg) if rgvarg
    Fiddle.free(ref0) if ref0
    Fiddle.free(ref1) if ref1
  end

  # Guards against the priority bug class where the Hash check and the
  # outargs-Array check are written as two independent `if`s instead of
  # `if/elsif` -- with with_outargs: true AND a Hash return, both
  # conditions would be true, so a regression would call both write paths
  # instead of only write_hash_result. write_hash_result itself needs a
  # live ITypeInfo::GetNames call (COM-only), so we spy on both methods
  # rather than let them run for real.
  def test_handle_invoke_prefers_hash_writeback_over_array_outargs
    ev = WIN32OLE::Event.allocate
    ev.instance_variable_set(:@events, [{ name: nil, proc: proc { { 0 => 1 } }, with_outargs: true }])
    ev.instance_variable_set(:@handler, nil)
    ev.instance_variable_set(:@event_typeinfo_ptr, 0)
    ev.define_singleton_method(:resolve_event_name) { |_dispid| 'Whatever' }

    hash_called = false
    array_called = false
    ev.define_singleton_method(:write_hash_result) { |*_args| hash_called = true }
    ev.define_singleton_method(:write_array_outargs) { |*_args| array_called = true }

    dispparams = [0, 0, 0, 0].pack("#{W::PACK_PTR}#{W::PACK_PTR}LL")
    dispparams_ptr = Fiddle::Pointer.to_ptr(dispparams)

    ev.send(:handle_invoke, 1, dispparams_ptr, nil)

    assert(hash_called)
    assert_false(array_called)
  end

  def capture_stderr
    old = $stderr
    $stderr = StringIO.new
    yield
    $stderr.string
  ensure
    $stderr = old
  end
end
end
