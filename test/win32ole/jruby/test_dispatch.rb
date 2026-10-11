# test/win32ole/jruby/test_dispatch.rb
require 'test/unit'

if RUBY_ENGINE == 'jruby'
require 'win32ole/jruby'

class TestDispatch < Test::Unit::TestCase
  W = WIN32OLE::Win32

  # A fake IDispatch: GetIDsOfNames hands out 100, 101, 102... in name
  # order and records the names; Invoke records the DISPPARAMS it got and
  # returns VT_EMPTY.
  def setup
    @names = []
    @calls = []
    names = @names
    calls = @calls
    get_ids = Fiddle::Closure::BlockCaller.new(
      W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::STDCALL
    ) do |_this, _riid, rgsz, count, _lcid, out|
      ptrs = Fiddle::Pointer.new(rgsz.to_i)[0, count * W::PTR_SIZE].unpack("#{W::PACK_PTR}*")
      names << ptrs.map { |p| W.wstr_to_s(p) }
      Fiddle::Pointer.new(out.to_i)[0, count * 4] = (100...(100 + count)).to_a.pack('l*')
      0
    end
    invoke = Fiddle::Closure::BlockCaller.new(
      W::LONG,
      [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::WORD, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP],
      W::STDCALL
    ) do |_this, dispid, _riid, _lcid, _flags, pdp, _result, _excep, _argerr|
      rgvarg, rgnamed, cargs, cnamed = Fiddle::Pointer.new(pdp.to_i)[0, W::PTR_SIZE * 2 + 8]
                                                    .unpack("#{W::PACK_PTR}#{W::PACK_PTR}LL")
      values = Array.new(cargs) do |i|
        _vt, payload = W.unpack_variant(Fiddle::Pointer.new(rgvarg + i * W::VARIANT_SIZE)[0, W::VARIANT_SIZE])
        W.unpack_i4(payload)
      end
      named = cnamed.zero? ? [] : Fiddle::Pointer.new(rgnamed)[0, cnamed * 4].unpack('l*')
      calls << { dispid: dispid, values: values, named: named }
      0
    end
    @closures = [get_ids, invoke]
    @vtable = Fiddle::Pointer.malloc(W::PTR_SIZE * 7)
    7.times { |i| @vtable[i * W::PTR_SIZE, W::PTR_SIZE] = [0].pack(W::PACK_PTR) }
    @vtable[5 * W::PTR_SIZE, W::PTR_SIZE] = [get_ids.to_i].pack(W::PACK_PTR)
    @vtable[6 * W::PTR_SIZE, W::PTR_SIZE] = [invoke.to_i].pack(W::PACK_PTR)
    @obj = Fiddle::Pointer.malloc(W::PTR_SIZE)
    @obj[0, W::PTR_SIZE] = [@vtable.to_i].pack(W::PACK_PTR)
    @ole = WIN32OLE.allocate # no finalizer: the fake has no Release
    @ole.instance_variable_set(:@ptr, @obj.to_i)
  end

  def teardown
    Fiddle.free(@obj) if @obj
    Fiddle.free(@vtable) if @vtable
  end

  def test_trailing_hash_becomes_named_arguments_ahead_of_reversed_positionals
    @ole.getDetailsOf(7, 8, { iColumn: 1, 'vItem' => 2 })

    assert_equal([%w[getDetailsOf], %w[getDetailsOf iColumn vItem]], @names)
    assert_equal({ dispid: 100, values: [1, 2, 8, 7], named: [101, 102] }, @calls.last)
  end

  def test_named_argument_keys_must_be_strings_or_symbols
    assert_raise(TypeError) { @ole.getDetailsOf({ 1 => 2 }) }
  end

  def test_no_hash_means_no_named_arguments
    @ole.add(1, 2)
    assert_equal({ dispid: 100, values: [2, 1], named: [] }, @calls.last)
  end

  def with_refcount(count)
    refs = [count]
    add_ref = Fiddle::Closure::BlockCaller.new(W::DWORD, [W::VOIDP], W::STDCALL) { |_| refs[0] += 1 }
    release = Fiddle::Closure::BlockCaller.new(W::DWORD, [W::VOIDP], W::STDCALL) { |_| refs[0] -= 1 }
    @closures.push(add_ref, release)
    @vtable[1 * W::PTR_SIZE, W::PTR_SIZE] = [add_ref.to_i].pack(W::PACK_PTR)
    @vtable[2 * W::PTR_SIZE, W::PTR_SIZE] = [release.to_i].pack(W::PACK_PTR)
    refs
  end

  # fole_s_reference_count: AddRef+Release, reporting the count without
  # changing it.
  def test_s_ole_reference_count_reports_without_changing_it
    refs = with_refcount(3)
    assert_equal(3, WIN32OLE.ole_reference_count(@ole))
    assert_equal([3], refs)
  end

  # fole_s_free: one Release (if any reference is left), returning the
  # new count; the object itself is left as it is (a debugging aid).
  def test_s_ole_free_releases_once_and_returns_the_count
    refs = with_refcount(2)
    assert_equal(1, WIN32OLE.ole_free(@ole))
    assert_equal([1], refs)
    assert_equal(@obj.to_i, @ole.instance_variable_get(:@ptr))
  end

  # In ext/win32ole these are aliases, not help-file launchers:
  # ole_method_help is ole_method, ole_obj_help is ole_type.
  def test_help_methods_are_aliases
    assert_equal(:ole_method, WIN32OLE.instance_method(:ole_method_help).original_name)
    assert_equal(:ole_type, WIN32OLE.instance_method(:ole_obj_help).original_name)
  end

  # fole_method_help: the WIN32OLE::Method named +name+ on the object's
  # type (Method.new searches it case-insensitively, implemented types too).
  def test_ole_method_builds_the_named_method_on_the_objects_type
    type = Object.new
    @ole.define_singleton_method(:type_via_containing_typelib) { type }
    received = nil
    original = WIN32OLE::Method.method(:new)
    WIN32OLE::Method.define_singleton_method(:new) { |*args| received = args; :method }
    begin
      assert_equal(:method, @ole.ole_method('Add'))
    ensure
      WIN32OLE::Method.define_singleton_method(:new, original)
    end
    assert_equal([type, 'Add'], received)
    assert_raise(TypeError) { @ole.ole_method(1) }
  end

  # fole_free: Release the IDispatch once and forget it.
  def test_ole_free_releases_once_and_forgets_the_pointer
    released = []
    release = Fiddle::Closure::BlockCaller.new(W::DWORD, [W::VOIDP], W::STDCALL) { |this| released << this.to_i; 0 }
    @closures << release
    @vtable[2 * W::PTR_SIZE, W::PTR_SIZE] = [release.to_i].pack(W::PACK_PTR)

    assert_nil(@ole.ole_free)
    assert_nil(@ole.ole_free)
    assert_equal([@obj.to_i], released)
    assert_raise(WIN32OLE::RuntimeError) { @ole.add(1) }
  end

  def fake_object(slots)
    count = [slots.keys.max + 1, 7].max
    vtable = Fiddle::Pointer.malloc(W::PTR_SIZE * count)
    count.times { |i| vtable[i * W::PTR_SIZE, W::PTR_SIZE] = [0].pack(W::PACK_PTR) }
    slots.each do |index, (args, ret, body)|
      closure = Fiddle::Closure::BlockCaller.new(ret, args, W::STDCALL, &body)
      @closures << closure
      vtable[index * W::PTR_SIZE, W::PTR_SIZE] = [closure.to_i].pack(W::PACK_PTR)
    end
    obj = Fiddle::Pointer.malloc(W::PTR_SIZE)
    obj[0, W::PTR_SIZE] = [vtable.to_i].pack(W::PACK_PTR)
    @fakes.push(obj, vtable)
    obj
  end

  IID_IPERSISTMEMORY = [0xBD1AE5E0, 0xA6AE, 0x11CE, 0xBD, 0x37, 0x50, 0x42, 0x00, 0xC1, 0x00, 0x00].pack('LSSC8')

  # fole_activex_initialize: QueryInterface IPersistMemory and InitNew it.
  def activex_fixture(init_hr)
    @fakes = []
    calls = []
    persist = fake_object(
      2 => [[W::VOIDP], W::DWORD, proc { |_| calls << :release; 0 }],
      8 => [[W::VOIDP], W::LONG, proc { |_| calls << :init_new; init_hr }]
    )
    qi = Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL) do |_this, riid, ppv|
      calls << :qi if riid[0, 16] == IID_IPERSISTMEMORY
      ppv[0, W::PTR_SIZE] = [persist.to_i].pack(W::PACK_PTR)
      0
    end
    @closures << qi
    @vtable[0, W::PTR_SIZE] = [qi.to_i].pack(W::PACK_PTR)
    calls
  end

  def test_ole_activex_initialize_inits_ipersistmemory
    calls = activex_fixture(0)
    assert_nil(@ole.ole_activex_initialize)
    assert_equal(%i[qi init_new release], calls)
  ensure
    @fakes&.each { |p| Fiddle.free(p) }
  end

  def test_ole_activex_initialize_raises_when_initnew_fails
    original = W.method(:hresult_system_message)
    W.define_singleton_method(:hresult_system_message) { |_| '' } # FormatMessage
    activex_fixture(-2147467259) # E_FAIL
    assert_raise(WIN32OLE::RuntimeError) { @ole.ole_activex_initialize }
  ensure
    W.define_singleton_method(:hresult_system_message, original) if original
    @fakes&.each { |p| Fiddle.free(p) }
  end

  # _NewEnum (DISPID -4) returns a VT_UNKNOWN whose QueryInterface hands
  # back an IEnumVARIANT yielding the I4s 10 and 20, then S_FALSE.
  def test_each_walks_the_new_enum_ienumvariant
    @fakes = []
    released = []
    items = [10, 20]
    enum = fake_object(
      0 => [[W::VOIDP, W::VOIDP, W::VOIDP], W::LONG, proc { |this, _riid, ppv| ppv[0, W::PTR_SIZE] = [this.to_i].pack(W::PACK_PTR); 0 }],
      2 => [[W::VOIDP], W::DWORD, proc { |this| released << this.to_i; 0 }],
      3 => [[W::VOIDP, W::DWORD, W::VOIDP, W::VOIDP], W::LONG, proc do |_this, _n, var, _fetched|
        next 1 if items.empty? # S_FALSE
        var[0, W::VARIANT_SIZE] = W.pack_variant(W::VT_I4, W.pack_i4(items.shift))
        0
      end]
    )
    newenum = Fiddle::Closure::BlockCaller.new(
      W::LONG, [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::WORD, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL
    ) do |_this, dispid, _riid, _lcid, _flags, _pdp, result, _excep, _argerr|
      next -2147352573 unless dispid == -4 # DISP_E_MEMBERNOTFOUND
      result[0, W::VARIANT_SIZE] = W.pack_variant(W::VT_UNKNOWN, W.pack_pointer(enum.to_i))
      0
    end
    @closures << newenum
    @vtable[6 * W::PTR_SIZE, W::PTR_SIZE] = [newenum.to_i].pack(W::PACK_PTR)

    assert_equal([10, 20], @ole.each.to_a)
    assert_equal([enum.to_i, enum.to_i], released) # the _NewEnum result and the QI'd enumerator
  ensure
    @fakes&.each { |p| Fiddle.free(p) }
  end
end
end
