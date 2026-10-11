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
    VT_NULL     = 1
    VT_I2       = 2
    VT_I4       = 3
    VT_R4       = 4
    VT_R8       = 5
    VT_CY       = 6
    VT_DATE     = 7
    VT_BSTR     = 8
    VT_DISPATCH = 9
    VT_ERROR    = 10
    VT_BOOL     = 11
    VT_VARIANT  = 12
    VT_UNKNOWN  = 13
    VT_I1       = 16
    VT_UI1      = 17
    VT_UI2      = 18
    VT_UI4      = 19
    VT_I8       = 20
    VT_UI8      = 21
    VT_INT      = 22
    VT_UINT     = 23
    VT_RECORD   = 36
    VT_TYPEMASK = 0x0FFF
    VT_ARRAY    = 0x2000
    VT_BYREF    = 0x4000

    DISPATCH_METHOD      = 1
    DISPATCH_PROPERTYGET = 2
    DISPATCH_PROPERTYPUT = 4
    DISPATCH_PROPERTYPUTREF = 8
    DISPID_PROPERTYPUT   = -3
    DISPID_VALUE         = 0

    CLSCTX_INPROC_SERVER = 0x1
    CLSCTX_LOCAL_SERVER  = 0x4
    CLSCTX_REMOTE_SERVER = 0x10

    LOCALE_SYSTEM_DEFAULT = 0x0800
    LOCALE_USER_DEFAULT   = 0x0400
    LCID_INSTALLED        = 0x00000001

    CP_ACP        = 0
    CP_OEMCP      = 1
    CP_MACCP      = 2
    CP_THREAD_ACP = 3
    CP_SYMBOL     = 42
    CP_UTF7       = 65000
    CP_UTF8       = 65001
    CP_EUC_JP_51932 = 51932
    # set_ole_codepage accepts these even when not "installed".
    SPECIAL_CODEPAGES = [CP_ACP, CP_OEMCP, CP_MACCP, CP_THREAD_ACP, CP_SYMBOL, CP_UTF7, CP_UTF8, CP_EUC_JP_51932].freeze
    CODEPAGE_ERROR_MESSAGE = 'codepage should be WIN32OLE::CP_ACP, WIN32OLE::CP_OEMCP, WIN32OLE::CP_MACCP, ' \
                             'WIN32OLE::CP_THREAD_ACP, WIN32OLE::CP_SYMBOL, WIN32OLE::CP_UTF7, WIN32OLE::CP_UTF8, ' \
                             'or installed codepage.'

    # ext/win32ole/win32ole.c's ole_encoding2cp table (names compared
    # case-insensitively); anything else maps to CP_ACP.
    ENCODING_CODEPAGES = {
      'big5' => 950, 'cp51932' => 51932, 'cp850' => 850, 'cp852' => 852, 'cp855' => 855, 'cp949' => 949,
      'euc-jp' => 20932, 'euc-kr' => 51949, 'euc-tw' => 51950, 'gb18030' => 54936, 'gb2312' => 20936,
      'gbk' => 936, 'ibm437' => 437, 'ibm720' => 720, 'ibm737' => 737, 'ibm775' => 775, 'ibm852' => 852,
      'ibm855' => 855, 'ibm857' => 857, 'ibm860' => 860, 'ibm861' => 861, 'ibm862' => 862, 'ibm863' => 863,
      'ibm864' => 864, 'ibm865' => 865, 'ibm866' => 866, 'ibm869' => 869, 'iso-2022-jp' => 50220,
      'iso-8859-1' => 28591, 'iso-8859-15' => 28605, 'iso-8859-2' => 28592, 'iso-8859-3' => 28593,
      'iso-8859-4' => 28594, 'iso-8859-5' => 28595, 'iso-8859-6' => 28596, 'iso-8859-7' => 28597,
      'iso-8859-8' => 28598, 'iso-8859-9' => 28599, 'koi8-r' => 20866, 'koi8-u' => 21866, 'shift_jis' => 932,
      'utf-16be' => 1201, 'utf-16le' => 1200, 'utf-7' => 65000, 'utf-8' => 65001, 'windows-1250' => 1250,
      'windows-1251' => 1251, 'windows-1252' => 1252, 'windows-1253' => 1253, 'windows-1254' => 1254,
      'windows-1255' => 1255, 'windows-1256' => 1256, 'windows-1257' => 1257, 'windows-1258' => 1258,
      'windows-31j' => 932, 'windows-874' => 874, 'eucjp-ms' => 20932
    }.freeze

    CPINFOEX_SIZE = 4 + 2 + 12 + 2 + 4 + (260 * 2) # MaxCharSize..CodePage + WCHAR CodePageName[MAX_PATH]
    CPINFOEX_CODEPAGE_OFFSET = 20

    IID_NULL      = ("\x00" * 16).b
    IID_IDISPATCH = [0x00020400, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46].pack('LSSC8')
    IID_IENUMVARIANT = [0x00020404, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46].pack('LSSC8')
    IID_IPERSISTMEMORY = [0xBD1AE5E0, 0xA6AE, 0x11CE, 0xBD, 0x37, 0x50, 0x42, 0x00, 0xC1, 0x00, 0x00].pack('LSSC8')
    DISPID_NEWENUM = -4

    # Well-known, decades-stable OLE Automation interface IDs (ocidl.h /
    # objbase.h). Hand-transcribed, not pulled from a live header -- verify
    # against a real Windows build before anything else depends on it (same
    # caution Phase 3 gave IRecordInfo's vtable; see this plan's §8 risk #1).
    IID_IUNKNOWN = [0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46].pack('LSSC8')
    IID_ICONNECTIONPOINTCONTAINER = [0xB196B284, 0xBAB4, 0x101A, 0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07].pack('LSSC8')
    IID_IPROVIDECLASSINFO = [0xB196B283, 0xBAB4, 0x101A, 0xB6, 0x9C, 0x00, 0xAA, 0x00, 0x34, 0x1D, 0x07].pack('LSSC8')
    IID_IPROVIDECLASSINFO2 = [0xA6BC3AC0, 0xDBAA, 0x11CE, 0x9D, 0xE3, 0x00, 0xAA, 0x00, 0x4B, 0xB8, 0x51].pack('LSSC8')

    VT_FOR_TYPE = {
      i4: VT_I4, i8: VT_I8, r8: VT_R8, bool: VT_BOOL,
      empty: VT_EMPTY, bstr: VT_BSTR, dispatch: VT_DISPATCH, date: VT_DATE
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

    # Ruby -> NUL-terminated UTF-16LE. Like ole_vstr2wc, the string's own
    # encoding decides the conversion; an encoding Ruby can't transcode
    # (a dummy one such as "UTF-16", or binary with non-ASCII bytes) falls
    # back to the ANSI code page, which is what ole_encoding2cp's CP_ACP
    # default does for encodings outside its table.
    def wstr(str)
      s = str.is_a?(::String) ? str : str.to_s
      ansi = s.encoding.dummy? || (s.encoding == Encoding::BINARY && !s.ascii_only?)
      wide = ansi ? multi_byte_to_wide_char(CP_ACP, s) : s.encode('UTF-16LE').b
      wide << "\x00\x00".b
    end

    def encoding_to_codepage(enc)
      ENCODING_CODEPAGES.fetch(enc.name.downcase, CP_ACP)
    end

    def codepage
      init_codepage unless @codepage
      @codepage
    end

    # set_ole_codepage. On top of C's checks, the codepage must map to a
    # Ruby Encoding ("CP<n>", dummy ones included): C defines a dummy
    # encoding for any that doesn't, which isn't possible from Ruby, and
    # strings returned from COM have to be tagged with something.
    def codepage=(cp)
      raise TypeError, "no implicit conversion of #{cp.class} into Integer" unless cp.is_a?(::Integer)
      unless SPECIAL_CODEPAGES.include?(cp) || is_valid_code_page.call(cp) != 0
        raise WIN32OLE::RuntimeError, CODEPAGE_ERROR_MESSAGE
      end

      real = real_codepage(cp)
      enc = begin
        Encoding.find("CP#{real}")
      rescue ArgumentError
        raise WIN32OLE::RuntimeError, "#{CODEPAGE_ERROR_MESSAGE} (no Ruby encoding for CP#{real})"
      end
      @codepage = cp
      @codepage_encoding = enc
      @real_codepage = real
    end

    def init_codepage
      self.codepage = encoding_to_codepage(Encoding.default_internal || Encoding.default_external)
    end

    # ole_cp2encoding's mapping of the pseudo codepages to a real one.
    def real_codepage(cp)
      case cp
      when CP_ACP then get_acp.call
      when CP_OEMCP then get_oemcp.call
      when CP_MACCP, CP_THREAD_ACP
        info = ("\x00" * CPINFOEX_SIZE).b
        raise WIN32OLE::RuntimeError, 'cannot map codepage to encoding.' if get_cp_info_ex.call(cp, 0, info).zero?

        info[CPINFOEX_CODEPAGE_OFFSET, 4].unpack1('L')
      else cp
      end
    end

    # UTF-16LE bytes -> a String in WIN32OLE.codepage, the way ole_wc2vstr
    # tags COM strings with cWIN32OLE_enc. Unconvertible characters become
    # '?', as WideCharToMultiByte's default character does.
    def utf16_to_s(utf16)
      init_codepage unless @codepage
      if @codepage_encoding.dummy?
        wide_char_to_multi_byte(@real_codepage, utf16).force_encoding(@codepage_encoding)
      else
        utf16.dup.force_encoding('UTF-16LE').encode(@codepage_encoding, invalid: :replace, undef: :replace)
      end
    end

    def pack_variant(vt, body)
      body = body.b
      max = VARIANT_SIZE - 8
      if body.bytesize > max
        raise ArgumentError, "body must be <= #{max} bytes, got #{body.bytesize}"
      end

      [vt, 0, 0, 0].pack('S4') + body + ("\x00".b * (max - body.bytesize))
    end

    def unpack_variant(bytes, body_size: 8)
      vt, = bytes.unpack1('S')
      [vt, bytes[8, body_size]]
    end

    def pack_i1(value)    = [value].pack('c') + ("\x00".b * 7)
    def pack_ui1(value)   = [value].pack('C') + ("\x00".b * 7)
    def pack_i2(value)    = [value].pack('s') + ("\x00".b * 6)
    def pack_ui2(value)   = [value].pack('S') + ("\x00".b * 6)
    def pack_i4(value)    = [value].pack('l') + ("\x00".b * 4)
    def pack_ui4(value)   = [value].pack('L') + ("\x00".b * 4)
    def pack_i8(value)    = [value].pack('q')
    def pack_ui8(value)   = [value].pack('Q')
    def pack_int(value)   = pack_i4(value)
    def pack_uint(value)  = pack_ui4(value)
    def pack_r4(value)    = [value].pack('f') + ("\x00".b * 4)
    def pack_r8(value)    = [value].pack('d')
    def pack_bool(value)  = [value ? -1 : 0].pack('s') + ("\x00".b * 6)
    def pack_pointer(addr) = [addr].pack('Q')
    def pack_empty        = "\x00".b * 8
    def pack_error(value) = pack_i4(value)

    def unpack_i1(payload)    = payload.unpack1('c')
    def unpack_ui1(payload)   = payload.unpack1('C')
    def unpack_i2(payload)    = payload.unpack1('s')
    def unpack_ui2(payload)   = payload.unpack1('S')
    def unpack_i4(payload)    = payload.unpack1('l')
    def unpack_ui4(payload)   = payload.unpack1('L')
    def unpack_i8(payload)    = payload.unpack1('q')
    def unpack_ui8(payload)   = payload.unpack1('Q')
    def unpack_int(payload)   = unpack_i4(payload)
    def unpack_uint(payload)  = unpack_ui4(payload)
    def unpack_r4(payload)    = payload.unpack1('f')
    def unpack_r8(payload)    = payload.unpack1('d')
    def unpack_bool(payload)  = payload.unpack1('s') != 0
    def unpack_pointer(payload) = payload.unpack1('Q')
    def unpack_error(payload) = unpack_i4(payload)

    def ruby_to_variant_type(value)
      case value
      when String then :bstr
      when Integer then INT32_RANGE.cover?(value) ? :i4 : :i8
      when Float then :r8
      when true, false then :bool
      when nil then :empty
      when ::WIN32OLE then :dispatch
      when ::Time then :date
      else
        raise TypeError, "unsupported argument type for OLE call: #{value.class}"
      end
    end

    def variant_ruby_type(vt)
      case vt
      # A bare VT_VARIANT holds no value: ole_variant2val's VT_BSTR
      # conversion of it always fails, giving nil.
      when VT_EMPTY, VT_NULL, VT_VARIANT then :empty
      when VT_DATE then :date
      when VT_I4 then :i4
      when VT_I8 then :i8
      when VT_R8 then :r8
      when VT_BOOL then :bool
      when VT_BSTR then :bstr
      when VT_DISPATCH, VT_UNKNOWN then :dispatch
      when VT_I1 then :i1
      when VT_UI1 then :ui1
      when VT_I2 then :i2
      when VT_UI2 then :ui2
      when VT_UI4 then :ui4
      when VT_UI8 then :ui8
      when VT_INT then :int
      when VT_UINT then :uint
      when VT_R4 then :r4
      when VT_ERROR then :error
      else :other # VT_CY, VT_DECIMAL, ...: ole_variant2val's VT_BSTR fallback
      end
    end

    # Bytes a VT_BYREF pointer refers to, per referenced type.
    VALUE_SIZES = {
      VT_I1 => 1, VT_UI1 => 1, VT_I2 => 2, VT_UI2 => 2, VT_BOOL => 2,
      VT_I4 => 4, VT_UI4 => 4, VT_INT => 4, VT_UINT => 4, VT_R4 => 4, VT_ERROR => 4
    }.freeze

    def value_size(vt)
      return PTR_SIZE if (vt & VT_ARRAY) != 0 || [VT_BSTR, VT_DISPATCH, VT_UNKNOWN].include?(vt)

      VALUE_SIZES.fetch(vt, 8)
    end

    # The non-BYREF VARIANT a VT_BYREF one refers to (a copy of its bytes),
    # following VT_VARIANT|VT_BYREF chains as ole_variant2val does; a
    # non-BYREF VARIANT is returned as is.
    def deref_variant(bytes)
      vt, payload = unpack_variant(bytes)
      while (vt & VT_BYREF) != 0
        base_vt = vt & ~VT_BYREF
        addr = unpack_pointer(payload)
        return pack_variant(base_vt, pack_empty) if addr.zero?

        if base_vt == VT_VARIANT
          bytes = Fiddle::Pointer.new(addr)[0, VARIANT_SIZE]
          vt, payload = unpack_variant(bytes)
        else
          data = Fiddle::Pointer.new(addr)[0, value_size(base_vt)]
          return pack_variant(base_vt, data.ljust(8, "\x00".b))
        end
      end
      bytes
    end

    # VariantChangeTypeEx under WIN32OLE.locale; returns the new VARIANT.
    def change_variant_type(variant_bytes, vt)
      dest = ("\x00" * VARIANT_SIZE).b
      hr = variant_change_type.call(dest, variant_bytes, WIN32OLE.locale, 0, vt)
      raise WIN32OLE::RuntimeError, "failed to change variant type: #{hr_hex(hr)}" if failed?(hr)

      dest
    end

    def add_ref(iunknown_ptr)
      vtable_function(iunknown_ptr, 1, [VOIDP], DWORD).call(iunknown_ptr)
    end

    OLE_DATE_EPOCH = Time.utc(1899, 12, 30)
    SECONDS_PER_DAY = 86_400

    # An OLE DATE is days since 1899-12-30, but for a negative DATE the
    # fraction is the time of day itself (-1.25 is 1899-12-29 06:00), not
    # an offset toward zero. Like vtdate2rbtime/rbtime2vtdate, these work
    # on wall-clock fields in the local zone; computing them in UTC avoids
    # VariantTimeToSystemTime's whole-second rounding that C has to undo.
    def vtdate_to_time(date)
      days = date.truncate
      seconds = (days * SECONDS_PER_DAY) + ((date - days).abs * SECONDS_PER_DAY)
      wall = OLE_DATE_EPOCH + seconds
      Time.new(wall.year, wall.month, wall.day, wall.hour, wall.min, wall.sec + wall.subsec)
    end

    def time_to_vtdate(time)
      wall = Time.utc(time.year, time.month, time.day, time.hour, time.min, time.sec + time.subsec)
      days = (wall - OLE_DATE_EPOCH) / SECONDS_PER_DAY
      whole = days.floor
      whole.negative? && days != whole ? whole - (days - whole) : days.to_f
    end

    def dispatch_plan(name, args)
      if name.end_with?('=')
        unless args.size == 1
          raise ArgumentError, "property put takes exactly one argument, got #{args.size}"
        end

        # ole_propertyput sets both: PROPERTYPUTREF lets a by-reference
        # (object-valued) property be assigned as well.
        { name: name[0..-2], wflags: DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF, named_put: true }
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

    def query_interface_error_message(operation, detail)
      "failed to #{operation}: #{detail}"
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

    def is_valid_code_page
      @is_valid_code_page ||= Fiddle::Function.new(kernel32['IsValidCodePage'], [DWORD], DWORD, STDCALL)
    end

    def get_acp
      @get_acp ||= Fiddle::Function.new(kernel32['GetACP'], [], DWORD, STDCALL)
    end

    def get_oemcp
      @get_oemcp ||= Fiddle::Function.new(kernel32['GetOEMCP'], [], DWORD, STDCALL)
    end

    def get_cp_info_ex
      @get_cp_info_ex ||= Fiddle::Function.new(kernel32['GetCPInfoExW'], [DWORD, DWORD, VOIDP], DWORD, STDCALL)
    end

    def multi_byte_to_wide_char_fn
      @multi_byte_to_wide_char_fn ||= Fiddle::Function.new(
        kernel32['MultiByteToWideChar'], [DWORD, DWORD, VOIDP, LONG, VOIDP, LONG], LONG, STDCALL
      )
    end

    def wide_char_to_multi_byte_fn
      @wide_char_to_multi_byte_fn ||= Fiddle::Function.new(
        kernel32['WideCharToMultiByte'], [DWORD, DWORD, VOIDP, LONG, VOIDP, LONG, VOIDP, VOIDP], LONG, STDCALL
      )
    end

    # Returns UTF-16LE bytes (no terminator).
    def multi_byte_to_wide_char(cp, str)
      bytes = str.b
      return ''.b if bytes.empty?

      count = multi_byte_to_wide_char_fn.call(cp, 0, bytes, bytes.bytesize, nil, 0)
      raise WIN32OLE::RuntimeError, "fail to convert CP#{cp} to Unicode" if count.zero?

      out = ("\x00" * (count * 2)).b
      multi_byte_to_wide_char_fn.call(cp, 0, bytes, bytes.bytesize, out, count)
      out
    end

    # UTF-16LE bytes in, codepage bytes (binary String) out.
    def wide_char_to_multi_byte(cp, utf16)
      chars = utf16.bytesize / 2
      return ''.b if chars.zero?

      size = wide_char_to_multi_byte_fn.call(cp, 0, utf16, chars, nil, 0, nil, nil)
      raise WIN32OLE::RuntimeError, "fail to convert Unicode to CP#{cp}" if size.zero?

      out = ("\x00" * size).b
      wide_char_to_multi_byte_fn.call(cp, 0, utf16, chars, out, size, nil, nil)
      out
    end

    def expand_environment_strings_fn
      @expand_environment_strings_fn ||= Fiddle::Function.new(
        kernel32['ExpandEnvironmentStringsW'], [VOIDP, VOIDP, DWORD], DWORD, STDCALL
      )
    end

    # Registry values like a COM class's InprocServer32 path may contain
    # %ENVVAR%-style references (e.g. "%SystemRoot%\\system32\\shell32.dll")
    # that ExpandEnvironmentStrings resolves -- mirrors
    # ext/win32ole/win32ole_typelib.c's typelib_file_from_clsid.
    def expand_environment_strings(str)
      return str unless str.include?('%')

      out = ("\x00" * 520).b # 260 WCHARs, generous for a MAX_PATH-style value
      len = expand_environment_strings_fn.call(wstr(str), out, 260)
      return str if len.zero? || len > 260

      out[0, (len - 1) * 2].force_encoding('UTF-16LE').encode('UTF-8')
    end

    def ole_initialize_fn
      @ole_initialize_fn ||= Fiddle::Function.new(ole32['OleInitialize'], [VOIDP], LONG, STDCALL)
    end

    def ole_uninitialize_fn
      @ole_uninitialize_fn ||= Fiddle::Function.new(ole32['OleUninitialize'], [], VOID, STDCALL)
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

    def co_create_instance_ex
      @co_create_instance_ex ||= Fiddle::Function.new(
        ole32['CoCreateInstanceEx'], [VOIDP, VOIDP, DWORD, VOIDP, DWORD, VOIDP], LONG, STDCALL
      )
    end

    HKEY_LOCAL_MACHINE = 0x80000002

    def reg_connect_registry
      @reg_connect_registry ||= Fiddle::Function.new(
        TypeInfo.advapi32['RegConnectRegistryW'], [VOIDP, VOIDP, VOIDP], LONG, STDCALL
      )
    end

    # clsid_from_remote: +progid+'s CLSID as registered on +host+
    # (HKLM\SOFTWARE\Classes\<progid>\CLSID there), or nil.
    def clsid_from_remote(host, progid)
      out = ("\x00" * PTR_SIZE).b
      return nil unless reg_connect_registry.call(wstr(host), HKEY_LOCAL_MACHINE, out).zero?

      hlm = out.unpack1(PACK_PTR)
      begin
        value = TypeLib.with_reg_subkey(hlm, "SOFTWARE\\Classes\\#{progid}\\CLSID") do |hkey|
          TypeLib.reg_default_string_value(hkey)
        end
        return nil unless value

        clsid = ("\x00" * 16).b
        failed?(clsid_from_string.call(wstr(value), clsid)) ? nil : clsid
      ensure
        TypeInfo.reg_close_key.call(hlm)
      end
    end

    def create_bind_ctx
      @create_bind_ctx ||= Fiddle::Function.new(ole32['CreateBindCtx'], [DWORD, VOIDP], LONG, STDCALL)
    end

    def mk_parse_display_name
      @mk_parse_display_name ||= Fiddle::Function.new(
        ole32['MkParseDisplayName'], [VOIDP, VOIDP, VOIDP, VOIDP], LONG, STDCALL
      )
    end

    def get_active_object
      @get_active_object ||= Fiddle::Function.new(oleaut32['GetActiveObject'], [VOIDP, VOIDP, VOIDP], LONG, STDCALL)
    end

    def co_task_mem_free
      @co_task_mem_free ||= Fiddle::Function.new(ole32['CoTaskMemFree'], [VOIDP], VOID, STDCALL)
    end

    def co_create_guid
      @co_create_guid ||= Fiddle::Function.new(ole32['CoCreateGuid'], [VOIDP], LONG, STDCALL)
    end

    def prog_id_from_clsid_fn
      # ProgIDFromCLSID is an ole32 export, not oleaut32.
      @prog_id_from_clsid_fn ||= Fiddle::Function.new(ole32['ProgIDFromCLSID'], [VOIDP, VOIDP], LONG, STDCALL)
    end

    def prog_id_from_clsid(clsid_bytes)
      out = ("\x00" * PTR_SIZE).b
      hr = prog_id_from_clsid_fn.call(clsid_bytes, out)
      return nil if failed?(hr)

      # ProgIDFromCLSID's output is a CoTaskMemAlloc'd, NUL-terminated wide
      # string -- NOT a BSTR. Reading it with bstr_to_s (which trusts a
      # length prefix 4 bytes before the string) would read garbage-length
      # data out of bounds, and freeing it with SysFreeString (the BSTR
      # deallocator) would free CoTaskMem-allocated memory with the wrong
      # allocator, corrupting the heap.
      addr = out.unpack1(PACK_PTR)
      str = wstr_to_s(addr)
      co_task_mem_free.call(addr) unless addr.zero?
      str
    end

    def sys_alloc_string
      @sys_alloc_string ||= Fiddle::Function.new(oleaut32['SysAllocString'], [VOIDP], VOIDP, STDCALL)
    end

    def sys_free_string
      @sys_free_string ||= Fiddle::Function.new(oleaut32['SysFreeString'], [VOIDP], VOID, STDCALL)
    end

    def enum_system_locales_a
      @enum_system_locales_a ||= Fiddle::Function.new(
        kernel32['EnumSystemLocalesA'], [VOIDP, DWORD], DWORD, STDCALL
      )
    end

    # Mirrors ext/win32ole/win32ole.c's lcid_installed: EnumSystemLocalesA
    # hands each installed LCID to the callback as an 8-hex-digit ANSI
    # string (e.g. "00000409"), lowercase, zero-padded -- same format
    # produced by C's "%08lx". Returning 0 (FALSE) from the callback stops
    # the enumeration early once a match is found.
    def locale_installed?(lcid)
      target = format('%08x', lcid)
      installed = false
      callback = Fiddle::Closure::BlockCaller.new(DWORD, [VOIDP], STDCALL) do |str_ptr|
        if Fiddle::Pointer.new(str_ptr)[0, 8] == target
          installed = true
          0
        else
          1
        end
      end
      enum_system_locales_a.call(callback, LCID_INSTALLED)
      installed
    end

    FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100
    FORMAT_MESSAGE_FROM_SYSTEM     = 0x00001000
    FORMAT_MESSAGE_IGNORE_INSERTS  = 0x00000200

    def format_message
      @format_message ||= Fiddle::Function.new(
        kernel32['FormatMessageW'], [DWORD, VOIDP, DWORD, DWORD, VOIDP, DWORD, VOIDP], DWORD, STDCALL
      )
    end

    def native_pointer_for(buffer)
      Fiddle::Pointer.to_ptr(buffer)
    end

    def persistent_pointer_for(bytes)
      ptr = Fiddle::Pointer.malloc(bytes.bytesize)
      ptr[0, bytes.bytesize] = bytes
      ptr
    end

    # Real COM pointers are always heap addresses far above this floor (the
    # low 64KB of a Windows process's address space is permanently reserved
    # and never backed by a valid allocation). Rejecting anything below it
    # here, before the first native dereference, turns what would otherwise
    # be an unrescuable native access violation (a JVM/process crash) on a
    # bogus caller-supplied "pointer" into a clean Ruby TypeError.
    MIN_PLAUSIBLE_POINTER = 0x10000

    # The vtable address IS the COM object's implementation identity — every
    # instance of the same concrete class shares one vtable. Callers that
    # memoize a resolved Fiddle::Function should key their cache on this,
    # not on the object's own address: after an object is released, the OS
    # can and does reuse its heap address for an unrelated object with a
    # different vtable, so keying by object address risks returning a
    # function pointer resolved from the WRONG vtable for a live object
    # that merely reused a dead one's address.
    def vtable_address(object_addr)
      unless object_addr.is_a?(Integer) && object_addr >= MIN_PLAUSIBLE_POINTER
        raise TypeError, "expected a native pointer address (Integer >= #{MIN_PLAUSIBLE_POINTER}), got #{object_addr.inspect}"
      end

      Fiddle::Pointer.new(object_addr)[0, PTR_SIZE].unpack1(PTR_SIZE == 8 ? 'Q' : 'L')
    end

    def vtable_function(object_addr, index, arg_types, ret_type)
      vtable_addr = vtable_address(object_addr)
      func_addr = Fiddle::Pointer.new(vtable_addr)[index * PTR_SIZE, PTR_SIZE].unpack1(PTR_SIZE == 8 ? 'Q' : 'L')
      Fiddle::Function.new(func_addr, arg_types, ret_type, STDCALL)
    end

    # Generic COM QueryInterface, usable against any interface pointer --
    # every other vtable helper in this codebase is for a FIXED interface
    # (IDispatch/ITypeInfo/ITypeLib); WIN32OLE::Event is the first caller
    # that needs to ask an arbitrary object for an arbitrary interface.
    def query_interface(obj_addr, iid_bytes)
      ppv = ("\x00" * PTR_SIZE).b
      hr = vtable_function(obj_addr, 0, [VOIDP, VOIDP, VOIDP], LONG).call(obj_addr, iid_bytes, ppv)
      return nil if failed?(hr)

      ppv.unpack1(PACK_PTR)
    end

    def bstr_to_s(addr)
      return nil if addr.nil? || addr.zero?

      len_bytes = Fiddle::Pointer.new(addr - 4)[0, 4].unpack1('L')
      utf16 = Fiddle::Pointer.new(addr)[0, len_bytes]
      # Like ole_wc2vstr (strlen after conversion): stop at the first NUL,
      # which a BSTR's length can include.
      nul = utf16.unpack('v*').index(0)
      utf16_to_s(nul ? utf16[0, nul * 2] : utf16)
    end

    # For plain NUL-terminated wide strings (LPOLESTR/LPWSTR), unlike BSTR,
    # which has no length prefix -- e.g. ProgIDFromCLSID's CoTaskMemAlloc'd
    # output. Scans for the terminating UTF-16 NUL code unit.
    def wstr_to_s(addr)
      return nil if addr.nil? || addr.zero?

      ptr = Fiddle::Pointer.new(addr)
      len = 0
      len += 1 while ptr[len * 2, 2].unpack1('S') != 0
      utf16_to_s(ptr[0, len * 2])
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

    # var's body holds a pointer INTO realvar's own already-packed bytes --
    # realvar must outlive var (Phase 1 §4.5's "native address embedded as
    # data" keep-alive discipline: the caller, WIN32OLE::Variant, is what
    # keeps realvar's String reachable for as long as var is in use).
    # VT_VARIANT|VT_BYREF is the one exception (mirrors MRI's
    # ole_set_byref): the pointer targets realvar's own start (offset 0,
    # the whole VARIANT), not offset 8 (one scalar slot within it).
    #
    # realvar_ptr must be a persistent Fiddle::Pointer (from
    # persistent_pointer_for), NOT a plain Ruby String -- on JRuby,
    # Fiddle::Pointer.to_ptr(a_string) always returns a fresh, disconnected
    # copy with no write-back, so pointing into a String's own bytes here
    # would silently break VT_BYREF's whole purpose (a native out-parameter
    # write becoming visible when Ruby later reads realvar).
    def pack_byref(vt, realvar_ptr)
      offset = vt == VT_VARIANT ? 0 : 8
      pack_variant(vt | VT_BYREF, pack_pointer((realvar_ptr + offset).to_i))
    end

    def variant_change_type
      @variant_change_type ||= Fiddle::Function.new(
        oleaut32['VariantChangeTypeEx'], [VOIDP, VOIDP, DWORD, WORD, WORD], LONG, STDCALL
      )
    end

    PM_REMOVE = 0x0001
    MSG_SIZE = 64

    def user32
      @user32 ||= Fiddle.dlopen('user32')
    end

    def get_desktop_window
      @get_desktop_window ||= Fiddle::Function.new(user32['GetDesktopWindow'], [], VOIDP, STDCALL)
    end

    HH_DISPLAY_TOPIC = 0x00
    HH_HELP_CONTEXT = 0x0F

    # HtmlHelpA from HHCTRL.OCX, or nil when that isn't available.
    def html_help_fn
      return @html_help_fn if defined?(@html_help_fn)

      @html_help_fn =
        begin
          Fiddle::Function.new(Fiddle.dlopen('hhctrl.ocx')['HtmlHelpA'], [VOIDP, VOIDP, DWORD, VOIDP], VOIDP, STDCALL)
        rescue Fiddle::DLError
          nil
        end
    end

    # ole_show_help: open +helpfile+ at +context+ (falling back to just the
    # file), returning the help window handle, 0 when it couldn't.
    def show_html_help(helpfile, context)
      return 0 unless html_help_fn

      file = "#{helpfile}\x00".b
      desktop = get_desktop_window.call
      hwnd = html_help_fn.call(desktop, file, HH_HELP_CONTEXT, context).to_i
      hwnd = html_help_fn.call(desktop, file, HH_DISPLAY_TOPIC, context).to_i if hwnd.zero?
      hwnd
    end

    def peek_message
      @peek_message ||= Fiddle::Function.new(
        user32['PeekMessageW'], [VOIDP, VOIDP, DWORD, DWORD, DWORD], DWORD, STDCALL
      )
    end

    def translate_message
      @translate_message ||= Fiddle::Function.new(user32['TranslateMessage'], [VOIDP], DWORD, STDCALL)
    end

    # DispatchMessageW returns LRESULT, which is pointer-width (64-bit on
    # x64), not the 32-bit LONG this file uses for HRESULTs. Declaring it as
    # VOIDP keeps the return-value width right on both x86 and x64;
    # pump_windows_messages discards the value either way.
    def dispatch_message
      @dispatch_message ||= Fiddle::Function.new(user32['DispatchMessageW'], [VOIDP], VOIDP, STDCALL)
    end

    def pump_windows_messages
      msg = Fiddle::Pointer.malloc(MSG_SIZE)
      while peek_message.call(msg, nil, 0, 0, PM_REMOVE) != 0
        translate_message.call(msg)
        dispatch_message.call(msg)
      end
      nil
    ensure
      Fiddle.free(msg) if msg
    end
  end
end
