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
      @event_typeinfo_ptr = nil
      @ole = nil

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

    def unadvise
      return nil if @finalizer_state.nil? || @finalizer_state[:cp_ptr].nil?

      # ext/win32ole/win32ole_event.c:1164 (fev_unadvise): drain the message
      # queue BEFORE Unadvise, so an already-queued event can't fire into a
      # sink we are about to tear down. Deliberately NOT done in
      # self.finalizer -- a finalizer runs at GC time with no well-defined
      # thread or message-loop context.
      W.pump_windows_messages

      cp_ptr = @finalizer_state[:cp_ptr]
      W.vtable_function(cp_ptr, 6, [W::VOIDP, W::DWORD], W::LONG).call(cp_ptr, @finalizer_state[:cookie])
      W.vtable_function(cp_ptr, 2, [W::VOIDP], W::DWORD).call(cp_ptr)
      ti_ptr = @finalizer_state[:ti_ptr]
      W.vtable_function(ti_ptr, 2, [W::VOIDP], W::DWORD).call(ti_ptr) if ti_ptr && !ti_ptr.zero?
      Fiddle.free(Fiddle::Pointer.new(@finalizer_state[:sink_addr])) if @finalizer_state[:sink_addr]
      Fiddle.free(Fiddle::Pointer.new(@finalizer_state[:vtable_addr])) if @finalizer_state[:vtable_addr]

      @finalizer_state[:cp_ptr] = nil
      @finalizer_state[:ti_ptr] = nil
      @finalizer_state[:sink_addr] = nil
      @finalizer_state[:vtable_addr] = nil
      @sink_closures = nil
      @event_typeinfo_ptr = nil
      nil
    end

    private

    TKIND_COCLASS = 5
    private_constant :TKIND_COCLASS

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
    private_constant :IMPLTYPEFLAG_FDEFAULT, :IMPLTYPEFLAG_FSOURCE, :GUIDKIND_DEFAULT_SOURCE_DISP_IID

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
      if target_guid.nil?
        release_ptr(tlib_ptr)
        return [nil, nil]
      end

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
          # A candidate whose own TYPEATTR fetch failed has no GUID to
          # compare -- treat it as no match, never as equal to a
          # likewise-missing target GUID (nil == nil would be a false hit).
          !ref_guid.nil? && ref_guid == target_guid
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
      hr = W.vtable_function(pci2_ptr, 4, [W::VOIDP, W::DWORD, W::VOIDP], W::LONG).call(
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
      hr = W.vtable_function(pci_ptr, 3, [W::VOIDP, W::VOIDP], W::LONG).call(pci_ptr, ti_out)
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

    # ext/win32ole/win32ole_event.c:900-973 (ev_advise)
    def advise(ole, itf)
      iid_bytes, event_typeinfo_ptr = resolve_event_source(ole, itf)

      idispatch_ptr = ole.dispatch_ptr
      container_ptr = W.query_interface(idispatch_ptr, W::IID_ICONNECTIONPOINTCONTAINER)
      if container_ptr.nil?
        release_ptr(event_typeinfo_ptr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('query IConnectionPointContainer', 'E_NOINTERFACE')
      end

      cp_out = ("\x00" * W::PTR_SIZE).b
      hr = W.vtable_function(container_ptr, 4, [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG).call(
        container_ptr, iid_bytes, cp_out
      )
      release_ptr(container_ptr)
      if W.failed?(hr)
        release_ptr(event_typeinfo_ptr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('query IConnectionPoint', W.hr_hex(hr))
      end
      connection_point_ptr = cp_out.unpack1(W::PACK_PTR)

      sink_addr, vtable_addr, closures = build_sink(iid_bytes, event_typeinfo_ptr)
      @sink_closures = closures # keep the trampolines alive; see build_sink's own comment

      cookie_out = ("\x00" * 4).b
      # cookie_out is a DWORD* OUT-parameter (a Ruby byte buffer), so its
      # declared type is VOIDP -- the pointer -- not DWORD, the value it
      # points at.
      hr = W.vtable_function(connection_point_ptr, 5, [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG).call(
        connection_point_ptr, sink_addr, cookie_out
      )
      if W.failed?(hr)
        release_ptr(connection_point_ptr)
        release_ptr(event_typeinfo_ptr)
        Fiddle.free(Fiddle::Pointer.new(sink_addr))
        Fiddle.free(Fiddle::Pointer.new(vtable_addr))
        @sink_closures = nil
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('Advise', W.hr_hex(hr))
      end

      @event_typeinfo_ptr = event_typeinfo_ptr
      # Keep the source WIN32OLE reachable for as long as this Event is: its
      # own finalizer would otherwise Release the IDispatch* we are advised
      # on. A plain Ruby reference, not a duplicate AddRef/Release pair --
      # the same keep-alive discipline Phase 1 established.
      @ole = ole
      @finalizer_state = {
        cp_ptr: connection_point_ptr, cookie: cookie_out.unpack1('L'),
        ti_ptr: event_typeinfo_ptr, sink_addr: sink_addr, vtable_addr: vtable_addr
      }
      ObjectSpace.define_finalizer(self, self.class.finalizer(@finalizer_state))
    end

    # Mirrors variant.rb:282-300's discipline exactly: the finalizer proc
    # captures ONLY this plain data hash, never self and never the live
    # Fiddle::Closure objects (@sink_closures) -- capturing either would
    # create a reference cycle that prevents GC from ever running the
    # finalizer at all.
    def self.finalizer(state)
      proc do
        cp_ptr = state[:cp_ptr]
        if cp_ptr && !cp_ptr.zero?
          W.vtable_function(cp_ptr, 6, [W::VOIDP, W::DWORD], W::LONG).call(cp_ptr, state[:cookie])
          W.vtable_function(cp_ptr, 2, [W::VOIDP], W::DWORD).call(cp_ptr)
        end
        ti_ptr = state[:ti_ptr]
        W.vtable_function(ti_ptr, 2, [W::VOIDP], W::DWORD).call(ti_ptr) if ti_ptr && !ti_ptr.zero?
        Fiddle.free(Fiddle::Pointer.new(state[:sink_addr])) if state[:sink_addr]
        Fiddle.free(Fiddle::Pointer.new(state[:vtable_addr])) if state[:vtable_addr]
      end
    end

    def register_event(event, block, with_outargs)
      if @finalizer_state.nil? || @finalizer_state[:cp_ptr].nil?
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
    private_constant :SINK_VTBL_SLOTS

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
      # `rescue Exception`, deliberately, at every closure boundary in this
      # file: the rule is that NO Ruby exception may cross back into the OLE
      # server's native stack frame, and SystemStackError, NoMemoryError,
      # Interrupt and SignalException all descend from Exception rather than
      # StandardError. Letting one of those unwind through a native frame is
      # undefined behaviour, so this is one of the rare places where the
      # broad rescue is the correct choice rather than an anti-pattern.
      rescue Exception => e
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
      rescue Exception => e
        warn_closure_exception('GetIDsOfNames', e)
        DISP_E_UNKNOWNNAME
      end
    end

    DISP_E_UNKNOWNNAME = -2147352570
    E_NOINTERFACE = -2147467262
    DISP_E_BADINDEX = -2147352565
    private_constant :DISP_E_UNKNOWNNAME, :E_NOINTERFACE, :DISP_E_BADINDEX

    def warn_closure_exception(where, error)
      backtrace = error.backtrace || []
      warn "#{backtrace.first}: #{error.message} (#{error.class}) in WIN32OLE::Event sink's #{where}"
      backtrace.drop(1).each { |line| warn "\tfrom #{line}" }
    end

    # The event ITypeInfo* this sink resolves names against is read from
    # @event_typeinfo_ptr (set once by #advise), not captured here -- the
    # parameter is kept only so build_sink calls every closure builder the
    # same way.
    def invoke_closure(_event_typeinfo_ptr)
      Fiddle::Closure::BlockCaller.new(
        W::LONG, [W::VOIDP, W::LONG, W::VOIDP, W::DWORD, W::WORD, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::STDCALL
      ) do |_this, dispid, _riid, _lcid, _wflags, pdispparams_ptr, pvarresult_ptr, _pexcepinfo_ptr, _puargerr_ptr|
        handle_invoke(dispid, pdispparams_ptr, pvarresult_ptr)
        0 # NOERROR, always -- see Global Constraints: no exception may cross this boundary.
      rescue Exception => e
        warn_closure_exception('Invoke', e)
        0
      end
    end

    def resolve_event_name(dispid)
      bstr_out = ("\x00" * W::PTR_SIZE).b
      count_out = ("\x00" * 4).b
      hr = TI.get_names_fn(@event_typeinfo_ptr).call(@event_typeinfo_ptr, dispid, bstr_out, 1, count_out)
      return nil if W.failed?(hr)

      bstr = bstr_out.unpack1(W::PACK_PTR)
      name = W.bstr_to_s(bstr)
      W.sys_free_string.call(bstr) unless bstr.zero?
      name
    end

    # ext/win32ole/win32ole_event.c:813-836 (ole_search_event): a NAMED
    # match wins immediately; otherwise fall back to the one catch-all
    # (nil-name) entry, if any. Returns [entry_or_nil, is_default] --
    # is_default is true only when a catch-all entry was actually found
    # (matching the C source, which leaves *is_default FALSE when the
    # array has no catch-all at all -- drives whether the event name gets
    # prepended to the callback's args).
    def find_event_entry(name)
      fallback = nil
      is_default = false
      @events.each do |e|
        return [e, false] if e[:name] == name

        if e[:name].nil?
          fallback = e
          is_default = true
        end
      end
      [fallback, is_default]
    end

    def read_dispparams(pdispparams_ptr)
      rgvarg_addr = pdispparams_ptr[0, W::PTR_SIZE].unpack1(W::PACK_PTR)
      cargs = pdispparams_ptr[2 * W::PTR_SIZE, 4].unpack1('L')
      [cargs, rgvarg_addr]
    end

    # ext/win32ole/win32ole_event.c:132-234 (EVENTSINK_Invoke), the
    # non-hash/non-outargs subset -- Task 14 adds the Hash/Array
    # out-argument write-back on top of this.
    def handle_invoke(dispid, pdispparams_ptr, pvarresult_ptr)
      name = resolve_event_name(dispid)
      return if name.nil?

      entry, is_default = find_event_entry(name)
      handler_obj = nil
      mid = nil
      with_outargs = false
      if entry
        handler_obj = entry[:proc]
        mid = :call
        with_outargs = entry[:with_outargs]
      elsif @handler
        on_name = "on#{name}"
        if @handler.respond_to?(on_name)
          handler_obj = @handler
          mid = on_name
          is_default = false
        elsif @handler.respond_to?(:method_missing)
          handler_obj = @handler
          mid = :method_missing
          is_default = true
        end
      end
      return if handler_obj.nil? || mid.nil?

      args = []
      args << name if is_default
      cargs, rgvarg_addr = read_dispparams(pdispparams_ptr)
      cargs.times do |i|
        var_ptr = Fiddle::Pointer.new(rgvarg_addr + (cargs - i - 1) * W::VARIANT_SIZE)
        args << WIN32OLE.variant_bytes_to_ruby_value(var_ptr[0, W::VARIANT_SIZE])
      end
      outargv = nil
      if with_outargs
        outargv = []
        args << outargv
      end

      result = begin
        handler_obj.send(mid, *args)
      rescue Exception => e # see query_interface_closure: still inside the closure's native frame
        warn_closure_exception('an event callback', e)
        nil
      end

      if result.is_a?(Hash)
        write_hash_result(result, dispid, cargs, rgvarg_addr)
        # key? rather than `||`, so a legitimate {'return' => false} yields
        # false instead of falling through to the symbol key.
        result = result.key?('return') ? result['return'] : result[:return]
      elsif with_outargs && outargv.is_a?(Array)
        write_array_outargs(outargv, cargs, rgvarg_addr)
      end

      return if pvarresult_ptr.nil? || pvarresult_ptr.to_i.zero?

      bytes = begin
        WIN32OLE.ruby_value_to_variant_bytes(result, [])
      rescue StandardError
        W.pack_variant(W::VT_EMPTY, W.pack_empty)
      end
      pvarresult_ptr[0, W::VARIANT_SIZE] = bytes
    end

    # ext/win32ole/win32ole_event.c:333-399 (ole_val2ptr_variant), ported
    # 1:1 including its "silently do nothing for an unhandled
    # type/VARTYPE combination" fallthrough -- e.g. writing a String into a
    # VT_I4|BYREF slot is a deliberate no-op in the original, not a bug we
    # should "fix" by raising.
    def write_byref_variant(var_ptr, value)
      vt = var_ptr[0, 2].unpack1('S')
      return if (vt & W::VT_BYREF).zero?

      ref_addr = var_ptr[8, W::PTR_SIZE].unpack1(W::PACK_PTR)
      return if ref_addr.zero?

      ref_ptr = Fiddle::Pointer.new(ref_addr)
      base_vt = vt & ~W::VT_BYREF
      case value
      when String
        ref_ptr[0, W::PTR_SIZE] = [W.sys_alloc_string.call(W.wstr(value))].pack(W::PACK_PTR) if base_vt == W::VT_BSTR
      when Integer
        case base_vt
        when W::VT_UI1 then ref_ptr[0, 1] = [value].pack('C')
        when W::VT_I2 then ref_ptr[0, 2] = [value].pack('s')
        when W::VT_I4 then ref_ptr[0, 4] = [value].pack('l')
        when W::VT_R4 then ref_ptr[0, 4] = [value.to_f].pack('f')
        when W::VT_R8 then ref_ptr[0, 8] = [value.to_f].pack('d')
        end
      when Float
        case base_vt
        when W::VT_I2 then ref_ptr[0, 2] = [value.to_i].pack('s')
        when W::VT_I4 then ref_ptr[0, 4] = [value.to_i].pack('l')
        when W::VT_R4 then ref_ptr[0, 4] = [value].pack('f')
        when W::VT_R8 then ref_ptr[0, 8] = [value].pack('d')
        end
      when true, false
        ref_ptr[0, 2] = [value ? -1 : 0].pack('s') if base_vt == W::VT_BOOL
      end
    end

    def write_array_outargs(ary, cargs, rgvarg_addr)
      ary.each_with_index do |value, i|
        break if i >= cargs

        var_ptr = Fiddle::Pointer.new(rgvarg_addr + (cargs - i - 1) * W::VARIANT_SIZE)
        write_byref_variant(var_ptr, value)
      end
    end

    # ext/win32ole/win32ole_event.c:401-428 (hash2ptr_dispparams)
    def write_hash_result(hash, dispid, cargs, rgvarg_addr)
      names_out = Fiddle::Pointer.malloc(W::PTR_SIZE * (cargs + 1))
      count_out = ("\x00" * 4).b
      hr = TI.get_names_fn(@event_typeinfo_ptr).call(@event_typeinfo_ptr, dispid, names_out, cargs + 1, count_out)
      return if W.failed?(hr)

      len = count_out.unpack1('L')
      (len - 1).times do |i|
        bstr = names_out[(i + 1) * W::PTR_SIZE, W::PTR_SIZE].unpack1(W::PACK_PTR)
        key_name = W.bstr_to_s(bstr)
        W.sys_free_string.call(bstr) unless bstr.zero?

        value = hash[i]
        value = hash[key_name] if value.nil?
        value = hash[key_name.to_sym] if value.nil? && key_name

        var_ptr = Fiddle::Pointer.new(rgvarg_addr + (cargs - i - 1) * W::VARIANT_SIZE)
        write_byref_variant(var_ptr, value)
      end
    ensure
      Fiddle.free(names_out) if names_out
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
