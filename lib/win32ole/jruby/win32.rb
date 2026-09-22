require 'fiddle'

class WIN32OLE
  module Win32
    STDCALL = if Fiddle::Function.const_defined?(:STDCALL)
                Fiddle::Function::STDCALL
              else
                Fiddle::Function::DEFAULT
              end

    VOIDP = Fiddle::TYPE_VOIDP
    LONG  = Fiddle::TYPE_LONG
    DWORD = -Fiddle::TYPE_INT
    WORD  = -Fiddle::TYPE_SHORT
    VOID  = Fiddle::TYPE_VOID

    PTR_SIZE     = Fiddle::SIZEOF_VOIDP
    PACK_PTR     = PTR_SIZE == 8 ? 'Q' : 'L' # native pointer width, for packing real structs (DISPPARAMS, pointer arrays) — NOT for a VARIANT's 8-byte value slot, which always uses 'Q' regardless of platform (see pack_pointer)
    VARIANT_SIZE = PTR_SIZE == 8 ? 24 : 16

    VT_EMPTY    = 0
    VT_I4       = 3
    VT_R8       = 5
    VT_BSTR     = 8
    VT_DISPATCH = 9
    VT_BOOL     = 11
    VT_UNKNOWN  = 13
    VT_I8       = 20

    DISPATCH_METHOD      = 1
    DISPATCH_PROPERTYGET = 2
    DISPATCH_PROPERTYPUT = 4
    DISPID_PROPERTYPUT   = -3

    CLSCTX_INPROC_SERVER = 0x1
    CLSCTX_LOCAL_SERVER  = 0x4

    IID_NULL      = ("\x00" * 16).b
    IID_IDISPATCH = [0x00020400, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46].pack('LSSC8')

    VT_FOR_TYPE = {
      i4: VT_I4, i8: VT_I8, r8: VT_R8, bool: VT_BOOL,
      empty: VT_EMPTY, bstr: VT_BSTR, dispatch: VT_DISPATCH
    }.freeze

    INT32_RANGE = (-(2**31))..(2**31 - 1)

    # EXCEPINFO (oaidl.h): WORD wCode; WORD wReserved; BSTR bstrSource;
    # BSTR bstrDescription; BSTR bstrHelpFile; DWORD dwHelpContext;
    # PVOID pvReserved; HRESULT(*pfnDeferredFillIn)(...); SCODE scode;
    if PTR_SIZE == 8
      EXCEPINFO_SIZE = 64
      EXCEPINFO_OFFSETS = {
        wCode: 0, bstrSource: 8, bstrDescription: 16, bstrHelpFile: 24,
        dwHelpContext: 32, pvReserved: 40, pfnDeferredFillIn: 48, scode: 56
      }.freeze
    else
      EXCEPINFO_SIZE = 32
      EXCEPINFO_OFFSETS = {
        wCode: 0, bstrSource: 4, bstrDescription: 8, bstrHelpFile: 12,
        dwHelpContext: 16, pvReserved: 20, pfnDeferredFillIn: 24, scode: 28
      }.freeze
    end

    module_function

    def wstr(str)
      "#{str}\x00".encode('UTF-16LE').b
    end

    def pack_variant(vt, payload)
      payload = payload.b
      unless payload.bytesize == 8
        raise ArgumentError, "payload must be 8 bytes, got #{payload.bytesize}"
      end

      [vt, 0, 0, 0].pack('S4') + payload + ("\x00".b * (VARIANT_SIZE - 16))
    end

    def unpack_variant(bytes)
      vt, = bytes.unpack1('S')
      [vt, bytes[8, 8]]
    end

    def pack_i4(value)    = [value].pack('l') + ("\x00".b * 4)
    def pack_i8(value)    = [value].pack('q')
    def pack_r8(value)    = [value].pack('d')
    def pack_bool(value)  = [value ? -1 : 0].pack('s') + ("\x00".b * 6)
    def pack_pointer(addr) = [addr].pack('Q')
    def pack_empty        = "\x00".b * 8

    def unpack_i4(payload)    = payload.unpack1('l')
    def unpack_i8(payload)    = payload.unpack1('q')
    def unpack_r8(payload)    = payload.unpack1('d')
    def unpack_bool(payload)  = payload.unpack1('s') != 0
    def unpack_pointer(payload) = payload.unpack1('Q')

    def ruby_to_variant_type(value)
      case value
      when String then :bstr
      when Integer then INT32_RANGE.cover?(value) ? :i4 : :i8
      when Float then :r8
      when true, false then :bool
      when nil then :empty
      when ::WIN32OLE then :dispatch
      else
        raise TypeError, "unsupported argument type for OLE call: #{value.class}"
      end
    end

    def variant_ruby_type(vt)
      case vt
      when VT_EMPTY then :empty
      when VT_I4 then :i4
      when VT_I8 then :i8
      when VT_R8 then :r8
      when VT_BOOL then :bool
      when VT_BSTR then :bstr
      when VT_DISPATCH, VT_UNKNOWN then :dispatch
      else
        raise NotImplementedError, "VARTYPE #{vt} is not supported yet"
      end
    end

    def dispatch_plan(name, args)
      if name.end_with?('=')
        unless args.size == 1
          raise ArgumentError, "property put takes exactly one argument, got #{args.size}"
        end

        { name: name[0..-2], wflags: DISPATCH_PROPERTYPUT, named_put: true }
      else
        { name: name, wflags: DISPATCH_METHOD | DISPATCH_PROPERTYGET, named_put: false }
      end
    end

    def failed?(hr)
      hr.negative?
    end

    def hr_hex(hr)
      format('0x%08x', hr & 0xFFFFFFFF)
    end

    def method_error_message(method_name, detail)
      "(in OLE method `#{method_name}': )#{detail}"
    end

    def property_put_error_message(property_name, detail)
      "(in setting property `#{property_name}': )#{detail}"
    end

    def unknown_server_error_message(server_name)
      "unknown OLE server: `#{server_name}'"
    end

    def parse_excepinfo(bytes)
      o = EXCEPINFO_OFFSETS
      ptr_fmt = PTR_SIZE == 8 ? 'Q' : 'L'
      {
        w_code: bytes[o[:wCode], 2].unpack1('S'),
        bstr_source_ptr: bytes[o[:bstrSource], PTR_SIZE].unpack1(ptr_fmt),
        bstr_description_ptr: bytes[o[:bstrDescription], PTR_SIZE].unpack1(ptr_fmt),
        scode: bytes[o[:scode], 4].unpack1('l')
      }
    end

    def ole32
      @ole32 ||= Fiddle.dlopen('ole32')
    end

    def oleaut32
      @oleaut32 ||= Fiddle.dlopen('oleaut32')
    end

    def kernel32
      @kernel32 ||= Fiddle.dlopen('kernel32')
    end

    def co_initialize
      @co_initialize ||= Fiddle::Function.new(ole32['CoInitialize'], [VOIDP], LONG, STDCALL)
    end

    def co_uninitialize
      @co_uninitialize ||= Fiddle::Function.new(ole32['CoUninitialize'], [], VOID, STDCALL)
    end

    def clsid_from_progid
      @clsid_from_progid ||= Fiddle::Function.new(ole32['CLSIDFromProgID'], [VOIDP, VOIDP], LONG, STDCALL)
    end

    def clsid_from_string
      @clsid_from_string ||= Fiddle::Function.new(ole32['CLSIDFromString'], [VOIDP, VOIDP], LONG, STDCALL)
    end

    def co_create_instance
      @co_create_instance ||= Fiddle::Function.new(
        ole32['CoCreateInstance'], [VOIDP, VOIDP, DWORD, VOIDP, VOIDP], LONG, STDCALL
      )
    end

    def sys_alloc_string
      @sys_alloc_string ||= Fiddle::Function.new(oleaut32['SysAllocString'], [VOIDP], VOIDP, STDCALL)
    end

    def sys_free_string
      @sys_free_string ||= Fiddle::Function.new(oleaut32['SysFreeString'], [VOIDP], VOID, STDCALL)
    end

    FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100
    FORMAT_MESSAGE_FROM_SYSTEM     = 0x00001000
    FORMAT_MESSAGE_IGNORE_INSERTS  = 0x00000200

    def format_message
      @format_message ||= Fiddle::Function.new(
        kernel32['FormatMessageW'], [DWORD, VOIDP, DWORD, DWORD, VOIDP, DWORD, VOIDP], DWORD, STDCALL
      )
    end

    def native_address_of(buffer)
      Fiddle::Pointer.to_ptr(buffer).to_i
    end

    def vtable_function(object_addr, index, arg_types, ret_type)
      vtable_addr = Fiddle::Pointer.new(object_addr)[0, PTR_SIZE].unpack1(PTR_SIZE == 8 ? 'Q' : 'L')
      func_addr = Fiddle::Pointer.new(vtable_addr)[index * PTR_SIZE, PTR_SIZE].unpack1(PTR_SIZE == 8 ? 'Q' : 'L')
      Fiddle::Function.new(func_addr, arg_types, ret_type, STDCALL)
    end

    def bstr_to_s(addr)
      return nil if addr.nil? || addr.zero?

      len_bytes = Fiddle::Pointer.new(addr - 4)[0, 4].unpack1('L')
      Fiddle::Pointer.new(addr)[0, len_bytes].dup.force_encoding('UTF-16LE').encode('UTF-8')
    end

    def hresult_system_message(hr)
      buf_ptr = ("\x00" * PTR_SIZE).b
      flags = FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS
      count = format_message.call(flags, nil, hr, 0, buf_ptr, 0, nil)
      return '' if count.zero?

      addr = buf_ptr.unpack1(PTR_SIZE == 8 ? 'Q' : 'L')
      msg = bstr_free_local_string(addr, count)
      msg.chomp
    end

    # count is the number of UTF-16LE *characters* FormatMessageW wrote,
    # not bytes — this exact kind of factor-of-2 slip is what the spike's
    # own retrospective (design §1.2) warns about: easy to get subtly
    # wrong, and it only shows up once you actually run it on Windows,
    # which is why this task's real verification is the CI run in Task 7,
    # not this write-up.
    def bstr_free_local_string(addr, count)
      ptr = Fiddle::Pointer.new(addr)
      msg = ptr[0, count * 2].dup.force_encoding('UTF-16LE').encode('UTF-8')
      local_free.call(addr)
      msg
    end

    def local_free
      @local_free ||= Fiddle::Function.new(kernel32['LocalFree'], [VOIDP], VOIDP, STDCALL)
    end
  end
end
