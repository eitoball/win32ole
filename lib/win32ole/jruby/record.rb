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
  end
end
