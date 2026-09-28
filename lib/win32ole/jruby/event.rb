require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'
require 'win32ole/jruby/win32ole'

class WIN32OLE
  class Event
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    def initialize(ole, itf = nil)
      raise TypeError, '1st parameter must be WIN32OLE object' unless ole.is_a?(WIN32OLE)

      @events = []
      @handler = nil
      @finalizer_state = nil
      @sink_closures = nil

      advise(ole, itf)
    end

    def self.message_loop
      W.pump_windows_messages
    end

    def on_event(event = nil, &block)
      register_event(event, block, false)
    end

    def on_event_with_outargs(event = nil, &block)
      register_event(event, block, true)
    end

    def off_event(event = nil)
      name = event.nil? ? nil : normalize_event_name(event)
      @events.reject! { |e| e[:name] == name }
      nil
    end

    def handler=(obj)
      @handler = obj
    end

    def handler
      @handler
    end

    private

    TKIND_COCLASS = 5

    def release_ptr(ptr)
      return if ptr.nil? || ptr.zero?

      W.vtable_function(ptr, 2, [W::VOIDP], W::DWORD).call(ptr)
    end

    def get_type_info0(idispatch_ptr)
      out = ("\x00" * W::PTR_SIZE).b
      hr = TI.get_type_info_fn(idispatch_ptr).call(idispatch_ptr, 0, W::LOCALE_SYSTEM_DEFAULT, out)
      return nil if W.failed?(hr)

      out.unpack1(W::PACK_PTR)
    end

    def containing_typelib(itypeinfo_ptr)
      tlib_out = ("\x00" * W::PTR_SIZE).b
      index_out = ("\x00" * 4).b
      hr = TI.containing_typelib_fn(itypeinfo_ptr).call(itypeinfo_ptr, tlib_out, index_out)
      return nil if W.failed?(hr)

      tlib_out.unpack1(W::PACK_PTR)
    end

    def type_attr_ptr(itypeinfo_ptr)
      out = ("\x00" * W::PTR_SIZE).b
      hr = TI.type_attr_fn(itypeinfo_ptr).call(itypeinfo_ptr, out)
      return nil if W.failed?(hr)

      out.unpack1(W::PACK_PTR)
    end

    def type_name(itypeinfo_ptr)
      name_out = ("\x00" * W::PTR_SIZE).b
      TI.documentation_fn_for_typeinfo(itypeinfo_ptr).call(itypeinfo_ptr, -1, name_out, nil, nil, nil)
      bstr = name_out.unpack1(W::PACK_PTR)
      name = W.bstr_to_s(bstr)
      W.sys_free_string.call(bstr) unless bstr.zero?
      name
    end

    def type_guid(attr_ptr)
      ta = TI::TYPEATTR.new(attr_ptr)
      [ta.guid_Data1, ta.guid_Data2, ta.guid_Data3].pack('LSS') + ta.guid_Data4.pack('C8')
    end

    def impl_type_ref_typeinfo(itypeinfo_ptr, index)
      href_out = ("\x00" * 4).b
      hr = TI.ref_type_of_impl_type_fn(itypeinfo_ptr).call(itypeinfo_ptr, index, href_out)
      return nil if W.failed?(hr)

      ref_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.ref_type_info_fn(itypeinfo_ptr).call(itypeinfo_ptr, href_out.unpack1('L'), ref_out)
      return nil if W.failed?(hr)

      ref_out.unpack1(W::PACK_PTR)
    end

    # ext/win32ole/win32ole_event.c:481-590 (find_iid, pitf given): scans
    # every COCLASS in ole's containing typelib for an implemented type
    # named itf_name; that impl type's own GUID becomes the event source IID.
    def find_iid_by_name(ole, itf_name)
      itypeinfo_ptr = get_type_info0(ole.dispatch_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if itypeinfo_ptr.nil?

      tlib_ptr = containing_typelib(itypeinfo_ptr)
      release_ptr(itypeinfo_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if tlib_ptr.nil?

      found_iid = nil
      found_ti_ptr = nil
      count = TI.type_info_count_fn(tlib_ptr).call(tlib_ptr)
      count.times do |index|
        break if found_iid

        ti_out = ("\x00" * W::PTR_SIZE).b
        next if W.failed?(TI.type_info_fn(tlib_ptr).call(tlib_ptr, index, ti_out))

        ti_ptr = ti_out.unpack1(W::PACK_PTR)
        attr_ptr = type_attr_ptr(ti_ptr)
        if attr_ptr
          if TI::TYPEATTR.new(attr_ptr).typekind == TKIND_COCLASS
            impl_count = TI::TYPEATTR.new(attr_ptr).cImplTypes
            impl_count.times do |t|
              impl_ti_ptr = impl_type_ref_typeinfo(ti_ptr, t)
              next if impl_ti_ptr.nil?

              if type_name(impl_ti_ptr) == itf_name
                impl_attr_ptr = type_attr_ptr(impl_ti_ptr)
                if impl_attr_ptr
                  found_iid = type_guid(impl_attr_ptr)
                  TI.release_type_attr_fn(impl_ti_ptr).call(impl_ti_ptr, impl_attr_ptr)
                  found_ti_ptr = impl_ti_ptr
                end
              end
              release_ptr(impl_ti_ptr) unless impl_ti_ptr == found_ti_ptr
              break if found_iid
            end
          end
          TI.release_type_attr_fn(ti_ptr).call(ti_ptr, attr_ptr)
        end
        release_ptr(ti_ptr)
      end
      release_ptr(tlib_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if found_iid.nil?

      [found_iid, found_ti_ptr]
    end

    # ext/win32ole/win32ole_event.c:518-524 (find_iid, pitf NULL): a direct
    # GetTypeInfoOfGuid lookup, used when the caller already knows the IID
    # (the IProvideClassInfo2::GetGUID path in find_default_source).
    def find_iid_by_guid(ole, iid_bytes)
      itypeinfo_ptr = get_type_info0(ole.dispatch_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if itypeinfo_ptr.nil?

      tlib_ptr = containing_typelib(itypeinfo_ptr)
      release_ptr(itypeinfo_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if tlib_ptr.nil?

      ti_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.type_info_of_guid_fn(tlib_ptr).call(tlib_ptr, iid_bytes, ti_out)
      release_ptr(tlib_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if W.failed?(hr)

      ti_out.unpack1(W::PACK_PTR)
    end

    IMPLTYPEFLAG_FDEFAULT = 0x1
    IMPLTYPEFLAG_FSOURCE = 0x2
    GUIDKIND_DEFAULT_SOURCE_DISP_IID = 1

    def impl_type_flags(itypeinfo_ptr, index)
      flags_out = ("\x00" * 4).b
      hr = TI.impl_type_flags_fn(itypeinfo_ptr).call(itypeinfo_ptr, index, flags_out)
      return nil if W.failed?(hr)

      flags_out.unpack1('l')
    end

    # ext/win32ole/win32ole_event.c:667-701
    def find_default_source_from_typeinfo(ti_ptr, attr_ptr)
      count = TI::TYPEATTR.new(attr_ptr).cImplTypes
      count.times do |i|
        flags = impl_type_flags(ti_ptr, i)
        next if flags.nil?
        next if (flags & IMPLTYPEFLAG_FDEFAULT).zero? || (flags & IMPLTYPEFLAG_FSOURCE).zero?

        ref_ti_ptr = impl_type_ref_typeinfo(ti_ptr, i)
        return ref_ti_ptr if ref_ti_ptr
      end
      nil
    end

    # ext/win32ole/win32ole_event.c:592-665: find, in ti_ptr's containing
    # typelib, the COCLASS whose default impl type is ti_ptr itself.
    def find_coclass(ti_ptr)
      tlib_ptr = containing_typelib(ti_ptr)
      return [nil, nil] if tlib_ptr.nil?

      target_attr_ptr = type_attr_ptr(ti_ptr)
      target_guid = target_attr_ptr && type_guid(target_attr_ptr)
      TI.release_type_attr_fn(ti_ptr).call(ti_ptr, target_attr_ptr) if target_attr_ptr

      found_ti_ptr = nil
      found_attr_ptr = nil
      count = TI.type_info_count_fn(tlib_ptr).call(tlib_ptr)
      count.times do |i|
        break if found_ti_ptr

        ti2_out = ("\x00" * W::PTR_SIZE).b
        next if W.failed?(TI.type_info_fn(tlib_ptr).call(tlib_ptr, i, ti2_out))

        ti2_ptr = ti2_out.unpack1(W::PACK_PTR)
        attr2_ptr = type_attr_ptr(ti2_ptr)
        if attr2_ptr.nil?
          release_ptr(ti2_ptr)
          next
        end
        if TI::TYPEATTR.new(attr2_ptr).typekind != TKIND_COCLASS
          TI.release_type_attr_fn(ti2_ptr).call(ti2_ptr, attr2_ptr)
          release_ptr(ti2_ptr)
          next
        end

        matched = TI::TYPEATTR.new(attr2_ptr).cImplTypes.times.any? do |j|
          flags = impl_type_flags(ti2_ptr, j)
          next false if flags.nil? || (flags & IMPLTYPEFLAG_FDEFAULT).zero?

          ref_ti_ptr = impl_type_ref_typeinfo(ti2_ptr, j)
          next false if ref_ti_ptr.nil?

          ref_attr_ptr = type_attr_ptr(ref_ti_ptr)
          ref_guid = ref_attr_ptr && type_guid(ref_attr_ptr)
          TI.release_type_attr_fn(ref_ti_ptr).call(ref_ti_ptr, ref_attr_ptr) if ref_attr_ptr
          release_ptr(ref_ti_ptr)
          ref_guid == target_guid
        end

        if matched
          found_ti_ptr = ti2_ptr
          found_attr_ptr = attr2_ptr
        else
          TI.release_type_attr_fn(ti2_ptr).call(ti2_ptr, attr2_ptr)
          release_ptr(ti2_ptr)
        end
      end
      release_ptr(tlib_ptr)
      [found_ti_ptr, found_attr_ptr]
    end

    def provide_class_info2_iid(idispatch_ptr)
      pci2_ptr = W.query_interface(idispatch_ptr, W::IID_IPROVIDECLASSINFO2)
      return nil if pci2_ptr.nil?

      iid_out = ("\x00" * 16).b
      hr = W.vtable_function(pci2_ptr, 4, [W::DWORD, W::VOIDP], W::LONG).call(
        pci2_ptr, GUIDKIND_DEFAULT_SOURCE_DISP_IID, iid_out
      )
      release_ptr(pci2_ptr)
      return nil if W.failed?(hr)

      iid_out
    end

    def provide_class_info_typeinfo(idispatch_ptr)
      pci_ptr = W.query_interface(idispatch_ptr, W::IID_IPROVIDECLASSINFO)
      return nil if pci_ptr.nil?

      ti_out = ("\x00" * W::PTR_SIZE).b
      hr = W.vtable_function(pci_ptr, 3, [W::VOIDP], W::LONG).call(pci_ptr, ti_out)
      release_ptr(pci_ptr)
      return nil if W.failed?(hr)

      ti_out.unpack1(W::PACK_PTR)
    end

    # ext/win32ole/win32ole_event.c:703-786, minus the GetGUID/find_iid
    # fast path (handled by resolve_event_source below, since it needs an
    # `iid_bytes` result rather than an `itypeinfo_ptr` result).
    def find_default_source(ole)
      itypeinfo_ptr = provide_class_info_typeinfo(ole.dispatch_ptr) || get_type_info0(ole.dispatch_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if itypeinfo_ptr.nil?

      attr_ptr = type_attr_ptr(itypeinfo_ptr)
      if attr_ptr.nil?
        release_ptr(itypeinfo_ptr)
        raise WIN32OLE::RuntimeError, 'interface not found'
      end

      result_ti_ptr = find_default_source_from_typeinfo(itypeinfo_ptr, attr_ptr)
      if result_ti_ptr.nil?
        co_ti_ptr, co_attr_ptr = find_coclass(itypeinfo_ptr)
        if co_ti_ptr
          result_ti_ptr = find_default_source_from_typeinfo(co_ti_ptr, co_attr_ptr)
          TI.release_type_attr_fn(co_ti_ptr).call(co_ti_ptr, co_attr_ptr)
          release_ptr(co_ti_ptr)
        end
      end
      TI.release_type_attr_fn(itypeinfo_ptr).call(itypeinfo_ptr, attr_ptr)
      release_ptr(itypeinfo_ptr)
      raise WIN32OLE::RuntimeError, 'interface not found' if result_ti_ptr.nil?

      result_attr_ptr = type_attr_ptr(result_ti_ptr)
      if result_attr_ptr.nil?
        release_ptr(result_ti_ptr)
        raise WIN32OLE::RuntimeError, 'interface not found'
      end
      guid = type_guid(result_attr_ptr)
      TI.release_type_attr_fn(result_ti_ptr).call(result_ti_ptr, result_attr_ptr)
      [guid, result_ti_ptr]
    end

    # The single entry point #advise (Task 13) calls.
    def resolve_event_source(ole, itf)
      return find_iid_by_name(ole, itf) if itf

      iid_bytes = provide_class_info2_iid(ole.dispatch_ptr)
      if iid_bytes
        begin
          return [iid_bytes, find_iid_by_guid(ole, iid_bytes)]
        rescue WIN32OLE::RuntimeError
          # IProvideClassInfo2 succeeded but the IID it named isn't
          # resolvable in this typelib -- fall through to the
          # IProvideClassInfo/ImplType-traversal path below, matching
          # win32ole_event.c:730-736's own fallthrough.
        end
      end
      find_default_source(ole)
    end

    # Built up across Tasks 9-13; a successful construction isn't
    # exercised by any test until Task 13 wires the real implementation in.
    def advise(ole, itf)
      raise NotImplementedError, 'advise is implemented in Task 13'
    end

    def register_event(event, block, with_outargs)
      if @finalizer_state.nil?
        raise WIN32OLE::RuntimeError, 'IConnectionPoint not found. You must call advise at first.'
      end

      name = event.nil? ? nil : normalize_event_name(event)
      @events.reject! { |e| e[:name] == name }
      @events << { name: name, proc: block, with_outargs: with_outargs }
      nil
    end

    def normalize_event_name(event)
      unless event.is_a?(String) || event.is_a?(Symbol)
        raise TypeError, 'wrong argument type (expected String or Symbol)'
      end

      event.to_s
    end

    SINK_VTBL_SLOTS = 7

    def query_interface_closure(sink_addr, source_iid_bytes, refcount)
      Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL) do |_this, riid_ptr, ppv_ptr|
        riid = riid_ptr[0, 16]
        if riid == W::IID_IUNKNOWN || riid == W::IID_IDISPATCH || riid == source_iid_bytes
          ppv_ptr[0, W::PTR_SIZE] = [sink_addr].pack(W::PACK_PTR)
          refcount[0] += 1
          0
        else
          ppv_ptr[0, W::PTR_SIZE] = [0].pack(W::PACK_PTR)
          E_NOINTERFACE
        end
      rescue StandardError, ScriptError => e
        warn_closure_exception('QueryInterface', e)
        E_NOINTERFACE
      end
    end

    def add_ref_closure(refcount)
      Fiddle::Closure::BlockCaller.new(W::DWORD, [W::VOIDP], W::STDCALL) do |_this|
        refcount[0] += 1
      end
    end

    def release_closure(refcount)
      Fiddle::Closure::BlockCaller.new(W::DWORD, [W::VOIDP], W::STDCALL) do |_this|
        refcount[0] -= 1
        refcount[0]
      end
    end

    def get_type_info_count_closure
      Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::VOIDP], W::STDCALL) do |_this, pct_ptr|
        pct_ptr[0, 4] = [0].pack('L')
        0
      end
    end

    def get_type_info_closure
      Fiddle::Closure::BlockCaller.new(W::LONG, [W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::STDCALL) do |_this, _idx, _lcid, ppti_ptr|
        ppti_ptr[0, W::PTR_SIZE] = [0].pack(W::PACK_PTR)
        DISP_E_BADINDEX
      end
    end

    def get_ids_of_names_closure(event_typeinfo_ptr)
      Fiddle::Closure::BlockCaller.new(
        W::LONG, [W::VOIDP, W::VOIDP, W::VOIDP, W::DWORD, W::DWORD, W::VOIDP], W::STDCALL
      ) do |_this, _riid, names_ptr, cnames, _lcid, dispids_ptr|
        TI.get_ids_of_names_fn(event_typeinfo_ptr).call(event_typeinfo_ptr, names_ptr.to_i, cnames, dispids_ptr.to_i)
      rescue StandardError, ScriptError => e
        warn_closure_exception('GetIDsOfNames', e)
        DISP_E_UNKNOWNNAME
      end
    end

    DISP_E_UNKNOWNNAME = -2147352570
    E_NOINTERFACE = -2147467262
    DISP_E_BADINDEX = -2147352565

    def warn_closure_exception(where, error)
      warn "#{error.backtrace&.first}: #{error.message} (#{error.class}) in WIN32OLE::Event sink's #{where}"
    end

    def invoke_closure(_event_typeinfo_ptr)
      Fiddle::Closure::BlockCaller.new(
        W::LONG, [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::WORD, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL
      ) { |*| 0 } # NOERROR; replaced with real dispatch in Task 12
    end

    # Builds a fresh 7-slot IDispatch-shaped vtable (QueryInterface, AddRef,
    # Release, GetTypeInfoCount, GetTypeInfo, GetIDsOfNames, Invoke -- same
    # order/shape as ext/win32ole/win32ole_event.c's IEventSinkVtbl) backed
    # by Fiddle::Closure::BlockCaller trampolines. Returns raw addresses (not
    # Fiddle::Pointer wrappers) plus the closures themselves -- the CALLER
    # must keep `closures` referenced for as long as the sink is advised
    # (GC'ing a Closure frees its native trampoline), matching variant.rb's
    # own finalizer-state discipline (see #advise, Task 13).
    def build_sink(source_iid_bytes, event_typeinfo_ptr)
      sink_ptr = Fiddle::Pointer.malloc(W::PTR_SIZE)
      sink_addr = sink_ptr.to_i
      refcount = [0]

      closures = [
        query_interface_closure(sink_addr, source_iid_bytes, refcount),
        add_ref_closure(refcount),
        release_closure(refcount),
        get_type_info_count_closure,
        get_type_info_closure,
        get_ids_of_names_closure(event_typeinfo_ptr),
        invoke_closure(event_typeinfo_ptr)
      ]

      vtable_ptr = Fiddle::Pointer.malloc(W::PTR_SIZE * SINK_VTBL_SLOTS)
      closures.each_with_index { |c, i| vtable_ptr[i * W::PTR_SIZE, W::PTR_SIZE] = [c.to_i].pack(W::PACK_PTR) }
      sink_ptr[0, W::PTR_SIZE] = [vtable_ptr.to_i].pack(W::PACK_PTR)

      [sink_addr, vtable_ptr.to_i, closures]
    end
  end
end
