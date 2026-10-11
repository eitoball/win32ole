# test/win32ole/jruby/test_connect.rb
require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby/win32ole'

# WIN32OLE.connect falls back to moniker binding (ole_bind_obj) when the
# name is neither a ProgID nor a CLSID -- e.g. "winmgmts:". The ole32 entry
# points are stubbed; the bind context and moniker are fake COM objects.
class TestConnectMoniker < Test::Unit::TestCase
  W = WIN32OLE::Win32
  CO_E_CLASSSTRING = -2147221005

  class FakeFn
    def initialize(&body)
      @body = body
    end

    def call(*args)
      @body.call(*args)
    end
  end

  def fake_object(slots)
    vtable = Fiddle::Pointer.malloc(W::PTR_SIZE * 9)
    9.times { |i| vtable[i * W::PTR_SIZE, W::PTR_SIZE] = [0].pack(W::PACK_PTR) }
    slots.each do |index, (args, ret, body)|
      closure = Fiddle::Closure::BlockCaller.new(ret, args, W::STDCALL, &body)
      @keep << closure
      vtable[index * W::PTR_SIZE, W::PTR_SIZE] = [closure.to_i].pack(W::PACK_PTR)
    end
    obj = Fiddle::Pointer.malloc(W::PTR_SIZE)
    obj[0, W::PTR_SIZE] = [vtable.to_i].pack(W::PACK_PTR)
    @mem.push(obj, vtable)
    obj.to_i
  end

  def stub(name, value)
    @originals[name] = W.method(name)
    W.define_singleton_method(name) { value }
  end

  def setup
    @ole_initialized = Thread.current.thread_variable_get(:win32ole_ole_initialized)
    @keep = []
    @mem = []
    @originals = {}
    @released = []
    @bound = nil
    released = @released
    release = [[W::VOIDP], W::DWORD, proc { |this| released << this.to_i; 0 }]
    @bindctx = fake_object(2 => release)
    @dispatch = 0x12345678
    dispatch = @dispatch
    bound = []
    @bound = bound
    @moniker = fake_object(
      2 => release,
      8 => [[W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::LONG, proc do |_this, pbc, _left, riid, ppv|
        bound << [pbc.to_i, riid[0, 16]]
        ppv[0, W::PTR_SIZE] = [dispatch].pack(W::PACK_PTR)
        0
      end]
    )
    bindctx = @bindctx
    moniker = @moniker
    @parsed = []
    parsed = @parsed
    stub(:ole_initialize_fn, FakeFn.new { |_| 0 })
    stub(:clsid_from_progid, FakeFn.new { |*| CO_E_CLASSSTRING })
    stub(:clsid_from_string, FakeFn.new { |*| CO_E_CLASSSTRING })
    stub(:create_bind_ctx, FakeFn.new { |_reserved, out| out[0, W::PTR_SIZE] = [bindctx].pack(W::PACK_PTR); 0 })
    stub(:mk_parse_display_name, FakeFn.new do |pbc, name, _eaten, out|
      parsed << [pbc, W.wstr_to_s(Fiddle::Pointer.to_ptr(name).to_i)]
      out[0, W::PTR_SIZE] = [moniker].pack(W::PACK_PTR)
      0
    end)
  end

  # The stubbed OleInitialize must not leave the thread marked as
  # initialized for the real COM tests that run after this one.
  def teardown
    @originals.each { |name, m| W.define_singleton_method(name, m) }
    @mem.each { |p| Fiddle.free(p) }
    Thread.current.thread_variable_set(:win32ole_ole_initialized, @ole_initialized)
  end

  def test_an_unresolvable_name_is_bound_as_a_moniker
    obj = WIN32OLE.allocate
    WIN32OLE.define_singleton_method(:wrap_dispatch_pointer) { |ptr| obj.instance_variable_set(:@ptr, ptr); obj }
    begin
      assert_same(obj, WIN32OLE.connect('winmgmts:'))
    ensure
      WIN32OLE.singleton_class.send(:remove_method, :wrap_dispatch_pointer)
    end

    assert_equal(@dispatch, obj.instance_variable_get(:@ptr))
    assert_equal([[@bindctx, 'winmgmts:']], @parsed)
    assert_equal([[@bindctx, W::IID_IDISPATCH]], @bound)
    assert_equal([@moniker, @bindctx], @released) # both released, the IDispatch kept
  end
end

# ole_initialize / ole_uninitialize: OleInitialize once per thread (COM
# initialization is per OS thread), undone by OleUninitialize.
class TestOleInitialize < Test::Unit::TestCase
  W = WIN32OLE::Win32

  def setup
    @calls = []
    calls = @calls
    @originals = %i[ole_initialize_fn ole_uninitialize_fn hresult_system_message].to_h { |name| [name, W.method(name)] }
    W.define_singleton_method(:hresult_system_message) { |_hr| '' } # FormatMessage
    @init_hr = 0
    test = self
    W.define_singleton_method(:ole_initialize_fn) { TestConnectMoniker::FakeFn.new { |_| calls << :init; test.init_hr } }
    W.define_singleton_method(:ole_uninitialize_fn) { TestConnectMoniker::FakeFn.new { calls << :uninit; nil } }
    @ole_initialized = Thread.current.thread_variable_get(:win32ole_ole_initialized)
    Thread.current.thread_variable_set(:win32ole_ole_initialized, nil)
  end

  attr_reader :init_hr

  def teardown
    @originals.each { |name, m| W.define_singleton_method(name, m) }
    Thread.current.thread_variable_set(:win32ole_ole_initialized, @ole_initialized)
  end

  def test_initializes_once_per_thread_and_uninitializes
    WIN32OLE.ole_initialize
    WIN32OLE.ole_initialize
    Thread.new { WIN32OLE.ole_initialize }.join
    WIN32OLE.ole_uninitialize
    WIN32OLE.ole_uninitialize
    assert_equal(%i[init init uninit], @calls)
  end

  def test_a_failed_initialize_raises
    @init_hr = -2147417850 # RPC_E_CHANGED_MODE
    assert_raise(RuntimeError) { WIN32OLE.ole_initialize }
  end
end
end
