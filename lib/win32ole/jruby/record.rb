# lib/win32ole/jruby/record.rb
require 'fiddle'
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'

class WIN32OLE
  class Record
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    # Full slot table transcribed from oaidl.h, kept for documentation
    # completeness (spec §3 lists everything but the six below as out of
    # scope for this phase -- unbound, no Fiddle::Function declared).
    IRECORDINFO_VTBL = {
      RecordInit: 3, RecordClear: 4, RecordCopy: 5, GetGuid: 6, GetName: 7,
      GetSize: 8, GetTypeInfo: 9, GetField: 10, GetFieldNoCopy: 11,
      PutField: 12, PutFieldNoCopy: 13, GetFieldNames: 14, IsMatchingType: 15,
      RecordCreate: 16, RecordCreateCopy: 17, RecordDestroy: 18
    }.freeze

    def self.record_init_fn(pri)
      @record_init_fns ||= {}
      @record_init_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:RecordInit], [W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.get_name_fn(pri)
      @get_name_fns ||= {}
      @get_name_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:GetName], [W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.get_size_fn(pri)
      @get_size_fns ||= {}
      @get_size_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:GetSize], [W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.get_field_no_copy_fn(pri)
      @get_field_no_copy_fns ||= {}
      @get_field_no_copy_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:GetFieldNoCopy], [W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.put_field_fn(pri)
      @put_field_fns ||= {}
      @put_field_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:PutField], [W::VOIDP, W::DWORD, W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.get_field_names_fn(pri)
      @get_field_names_fns ||= {}
      @get_field_names_fns[W.vtable_address(pri)] ||= W.vtable_function(
        pri, IRECORDINFO_VTBL[:GetFieldNames], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG
      )
    end

    def self.get_record_info_from_type_info_fn
      @get_record_info_from_type_info_fn ||= Fiddle::Function.new(
        W.oleaut32['GetRecordInfoFromTypeInfo'], [W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def initialize(typename, oleobj)
      typename = typename.to_s
      itypelib_ptr = resolve_itypelib_ptr(oleobj)
      pri = find_record_info(itypelib_ptr, typename)
      unless pri
        raise WIN32OLE::RuntimeError, "fail to query IRecordInfo interface for `#{typename}'"
      end

      set_record_info(pri, nil)
    end

    def to_h
      @fields
    end

    def typename
      @typename
    end

    def inspect
      "#<WIN32OLE::Record:#{@typename}>"
    end

    def method_missing(name, *args)
      sname = name.to_s
      case args.size
      when 0 then @fields.fetch(sname)
      when 1
        key = sname.end_with?('=') ? sname[0..-2] : sname
        @fields.fetch(key) # raises KeyError before writing, matching MRI
        @fields[key] = args.first
      else
        super
      end
    end

    def respond_to_missing?(name, include_private = false)
      @fields.key?(name.to_s.sub(/=\z/, '')) || super
    end

    def ole_instance_variable_get(name)
      unless name.is_a?(String) || name.is_a?(Symbol)
        raise TypeError, 'wrong argument type (expected String or Symbol)'
      end

      @fields.fetch(name.to_s)
    end

    def ole_instance_variable_set(name, val)
      unless name.is_a?(String) || name.is_a?(Symbol)
        raise TypeError, 'wrong argument type (expected String or Symbol)'
      end

      key = name.to_s
      @fields.fetch(key)
      @fields[key] = val
    end

    VT_RECORD_BODY_SIZE = W::PTR_SIZE * 2 # BRECORD: { PVOID pvRecord; IRecordInfo *pRecInfo; }

    def to_variant_bytes
      size_out = ("\x00" * 4).b
      hr = self.class.get_size_fn(@pri).call(@pri, size_out)
      raise WIN32OLE::RuntimeError, "failed to get size for allocation of VT_RECORD object: #{W.hr_hex(hr)}" if W.failed?(hr)

      # Port of ole_rec2variant (win32ole_record.c:101-103): free the
      # previously-allocated buffer (if any) before replacing it, so a
      # Record marshaled more than once doesn't leak one native buffer per
      # call. The freed/replaced pointer lives in @finalizer_state, the same
      # mutable Hash the GC finalizer (install_finalizer/self.finalizer)
      # reads from -- so the finalizer always frees whatever buffer is
      # current, without ever being re-registered.
      Fiddle.free(@finalizer_state[:buffer_ptr]) if @finalizer_state[:buffer_ptr]

      size = size_out.unpack1('L')
      buffer_ptr = Fiddle::Pointer.malloc(size)
      @finalizer_state[:buffer_ptr] = buffer_ptr
      hr = self.class.record_init_fn(@pri).call(@pri, buffer_ptr.to_i)
      raise WIN32OLE::RuntimeError, "failed to initialize VT_RECORD object: #{W.hr_hex(hr)}" if W.failed?(hr)

      @fields.each do |name, val|
        next if val.nil?

        # Port of hash2olerec (win32ole_record.c:73-81): free the temporary
        # VARIANT's own BSTR (there VariantClear, here bstrs_to_free) right
        # after PutField copies it, not batched at the end.
        bstrs_to_free = []
        var_bytes = WIN32OLE.ruby_value_to_variant_bytes(val, bstrs_to_free)
        hr = self.class.put_field_fn(@pri).call(
          @pri, W::DISPATCH_PROPERTYPUT, buffer_ptr.to_i, W.wstr(name), W.native_pointer_for(var_bytes)
        )
        bstrs_to_free.each { |bstr| W.sys_free_string.call(bstr) }
        raise WIN32OLE::RuntimeError, "failed to putfield of `#{name}': #{W.hr_hex(hr)}" if W.failed?(hr)
      end

      body = [buffer_ptr.to_i, @pri].pack("#{W::PACK_PTR}2")
      W.pack_variant(W::VT_RECORD, body)
    end

    def self.from_irecordinfo_and_buffer(pri, prec)
      allocate.tap { |rec| rec.send(:set_record_info, pri, prec) }
    end

    private

    def resolve_itypelib_ptr(oleobj)
      case oleobj
      when WIN32OLE::TypeLib
        oleobj.instance_variable_get(:@ptr)
      when WIN32OLE
        oleobj.ole_typelib.instance_variable_get(:@ptr)
      else
        raise TypeError, "2nd argument should be WIN32OLE object or WIN32OLE::TypeLib object, got #{oleobj.class}"
      end
    end

    # Port of recordinfo_from_itypelib (win32ole_record.c:29-59): linear-scan
    # the typelib's own members for a name match (GetDocumentation), same
    # already-open-typelib walk Phase 2's TypeLib#ole_types performs -- NOT
    # the registry-tree walk Phase 2 §3 excluded. Every non-matching
    # ITypeInfo* GetTypeInfo AddRef'd along the way must be Released, or
    # the scan leaks a reference per member it skips past.
    def find_record_info(itypelib_ptr, typename)
      count = TI.type_info_count_fn(itypelib_ptr).call(itypelib_ptr)
      count.times do |i|
        ti_out = ("\x00" * W::PTR_SIZE).b
        next if W.failed?(TI.type_info_fn(itypelib_ptr).call(itypelib_ptr, i, ti_out))

        itypeinfo_ptr = ti_out.unpack1(W::PACK_PTR)
        name_out = ("\x00" * W::PTR_SIZE).b
        TI.documentation_fn_for_typeinfo(itypeinfo_ptr).call(itypeinfo_ptr, -1, name_out, nil, nil, nil)
        name_bstr = name_out.unpack1(W::PACK_PTR)
        name = W.bstr_to_s(name_bstr)
        W.sys_free_string.call(name_bstr) unless name_bstr.zero?

        if name == typename
          pri_out = ("\x00" * W::PTR_SIZE).b
          hr = self.class.get_record_info_from_type_info_fn.call(itypeinfo_ptr, pri_out)
          release_itypeinfo(itypeinfo_ptr)
          return W.failed?(hr) ? nil : pri_out.unpack1(W::PACK_PTR)
        end
        release_itypeinfo(itypeinfo_ptr)
      end
      nil
    end

    def release_itypeinfo(itypeinfo_ptr)
      W.vtable_function(itypeinfo_ptr, 2, [W::VOIDP], W::DWORD).call(itypeinfo_ptr)
    end

    # Shared by .new (prec = nil, every field starts nil -- see this
    # task's own correction above) and Task 10's .from_irecordinfo_and_buffer
    # (prec = a real native buffer, fields read via GetFieldNoCopy). Ports
    # olerecord_set_ivar (win32ole_record.c:122-169).
    def set_record_info(pri, prec)
      @pri = pri
      # Mutable state the GC finalizer closure (below) shares a reference
      # to -- Task 10's to_variant_bytes mutates buffer_ptr in place on
      # every call. install_finalizer must run exactly once per instance
      # (set_record_info itself only ever runs once, from .new or
      # .from_irecordinfo_and_buffer): ObjectSpace.define_finalizer is
      # additive, so re-registering here would run two finalizers at GC
      # time and double-Release @pri.
      @finalizer_state = { pri: pri, buffer_ptr: nil }
      install_finalizer

      name_out = ("\x00" * W::PTR_SIZE).b
      if self.class.get_name_fn(pri).call(pri, name_out).zero?
        bstr = name_out.unpack1(W::PACK_PTR)
        @typename = W.bstr_to_s(bstr)
        W.sys_free_string.call(bstr) unless bstr.zero?
      end

      count_out = ("\x00" * 4).b
      hr = self.class.get_field_names_fn(pri).call(pri, count_out, nil)
      count = count_out.unpack1('L')
      @fields = {}
      return if W.failed?(hr) || count.zero?

      names_out = ("\x00" * (count * W::PTR_SIZE)).b
      count_out = [count].pack('L')
      self.class.get_field_names_fn(pri).call(pri, count_out, names_out)
      bstrs = names_out.unpack(W::PACK_PTR * count)

      bstrs.each do |bstr|
        name = W.bstr_to_s(bstr)
        val = nil
        if prec
          var_out = ("\x00" * W::VARIANT_SIZE).b
          pdata_out = ("\x00" * W::PTR_SIZE).b
          hr = self.class.get_field_no_copy_fn(pri).call(pri, prec, bstr, var_out, pdata_out)
          val = WIN32OLE.variant_bytes_to_ruby_value(var_out) if hr.zero?
        end
        @fields[name] = val
        W.sys_free_string.call(bstr) unless bstr.zero?
      end
    end

    def install_finalizer
      release_fn = W.vtable_function(@pri, 2, [W::VOIDP], W::DWORD)
      ObjectSpace.define_finalizer(self, self.class.finalizer(@finalizer_state, release_fn))
    end

    # state is the same Hash to_variant_bytes mutates in place -- read
    # state[:pri]/state[:buffer_ptr] fresh here (at GC time), never close
    # over a value frozen when the proc was created, or a buffer allocated
    # after this finalizer was installed would never be freed.
    def self.finalizer(state, release_fn)
      proc do
        pri = state[:pri]
        release_fn.call(pri) unless pri.nil? || pri.zero?
        buffer_ptr = state[:buffer_ptr]
        Fiddle.free(buffer_ptr) if buffer_ptr
      end
    end
  end
end
