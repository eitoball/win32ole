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
  end
end
