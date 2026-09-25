# lib/win32ole/jruby/record.rb
require 'fiddle'
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'
require 'win32ole/jruby/win32ole'
require 'win32ole/jruby/typelib'

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
      return if W.failed?(hr) || count.zero?

      names_out = ("\x00" * (count * W::PTR_SIZE)).b
      count_out = [count].pack('L')
      self.class.get_field_names_fn(pri).call(pri, count_out, names_out)
      bstrs = names_out.unpack(W::PACK_PTR * count)

      @fields = {}
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
      pri = @pri
      release_fn = W.vtable_function(pri, 2, [W::VOIDP], W::DWORD)
      ObjectSpace.define_finalizer(self, self.class.finalizer(pri, release_fn))
    end

    def self.finalizer(pri, release_fn)
      proc { release_fn.call(pri) unless pri.zero? }
    end
  end
end
