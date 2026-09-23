# lib/win32ole/jruby/typelib.rb
require 'win32ole/jruby/win32'
require 'win32ole/jruby/typeinfo'
require 'win32ole/jruby/type'

class WIN32OLE
  class TypeLib
    W = Win32
    TI = TypeInfo
    private_constant :W, :TI

    LIBFLAG_FHIDDEN = 0x1

    def initialize(itypelib_ptr)
      @ptr = itypelib_ptr

      lib_attr_out = ("\x00" * W::PTR_SIZE).b
      hr = TI.lib_attr_fn(@ptr).call(@ptr, lib_attr_out)
      if W.failed?(hr)
        raise WIN32OLE::QueryInterfaceError, W.query_interface_error_message('GetLibAttr', W.hr_hex(hr))
      end
      attr_ptr = lib_attr_out.unpack1(W::PACK_PTR)
      p = Fiddle::Pointer.new(attr_ptr)
      @guid_bytes = p[0, 16]
      @lcid = p[16, 4].unpack1('L')
      @major = p[16 + 4 + 4, 2].unpack1('S') # skip lcid(4) + syskind(4, enum-sized)
      @minor = p[16 + 4 + 4 + 2, 2].unpack1('S')
      @lib_flags = p[16 + 4 + 4 + 2 + 2, 2].unpack1('S')
      TI.release_tlib_attr_fn(@ptr).call(@ptr, attr_ptr)

      @name, @helpstring, @help_context, @helpfile = read_documentation(@ptr, -1)

      install_finalizer
    end

    def guid
      d1, d2, d3 = @guid_bytes.unpack('LSS')
      d4 = @guid_bytes[8, 8].unpack('C8')
      format('{%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x}', d1, d2, d3, *d4)
    end

    def name
      @name
    end

    def version
      "#{@major}.#{@minor}"
    end

    def major_version
      @major
    end

    def minor_version
      @minor
    end

    def visible?
      (@lib_flags & LIBFLAG_FHIDDEN) == 0
    end

    def library_name
      @name
    end

    def path
      reg_path_lookup(guid, version, @lcid)
    end

    def ole_types
      count = TI.type_info_count_fn(@ptr).call(@ptr)
      Array.new(count) do |i|
        ti_out = ("\x00" * W::PTR_SIZE).b
        hr = TI.type_info_fn(@ptr).call(@ptr, i, ti_out)
        next nil if W.failed?(hr)

        WIN32OLE::Type.new(ti_out.unpack1(W::PACK_PTR))
      end.compact
    end

    def inspect
      "#<WIN32OLE::TypeLib:#{name}>"
    end

    private

    def read_documentation(itypelib_ptr, index)
      name_out = ("\x00" * W::PTR_SIZE).b
      docstring_out = ("\x00" * W::PTR_SIZE).b
      helpcontext_out = ("\x00" * 4).b
      helpfile_out = ("\x00" * W::PTR_SIZE).b
      TI.documentation_fn_for_typelib(itypelib_ptr).call(
        itypelib_ptr, index, name_out, docstring_out, helpcontext_out, helpfile_out
      )
      name_bstr = name_out.unpack1(W::PACK_PTR)
      docstring_bstr = docstring_out.unpack1(W::PACK_PTR)
      helpfile_bstr = helpfile_out.unpack1(W::PACK_PTR)
      name = W.bstr_to_s(name_bstr)
      helpstring = W.bstr_to_s(docstring_bstr)
      helpfile = W.bstr_to_s(helpfile_bstr)
      [name_bstr, docstring_bstr, helpfile_bstr].each { |b| W.sys_free_string.call(b) unless b.zero? }
      [name, helpstring, helpcontext_out.unpack1('L'), helpfile]
    end

    def reg_path_lookup(guid_str, version_str, lcid)
      # HKEY_CLASSES_ROOT is itself the merged Classes root, so the subkey
      # path must NOT be prefixed with "SOFTWARE\Classes\" (that prefix is
      # only needed when opening under HKEY_LOCAL_MACHINE/HKEY_CURRENT_USER
      # directly, as MRI's own clsid_from_remote does for the DCOM path).
      key_path = W.wstr("TypeLib\\#{guid_str}\\#{version_str}\\#{lcid}\\win32")
      hkey_out = ("\x00" * W::PTR_SIZE).b
      err = TI.reg_open_key_ex.call(TI::HKEY_CLASSES_ROOT, key_path, 0, TI::KEY_READ, hkey_out)
      return nil unless err.zero?

      hkey = hkey_out.unpack1(W::PACK_PTR)
      begin
        empty_name = W.wstr('')
        type_out = ("\x00" * 4).b
        size_out = [520].pack('L') # 260 WCHARs, generous for a MAX_PATH-style value
        data_out = ("\x00" * 520).b
        err = TI.reg_query_value_ex.call(hkey, empty_name, nil, type_out, data_out, size_out)
        return nil unless err.zero? && type_out.unpack1('L') == TI::REG_SZ

        data_out[0, size_out.unpack1('L')].force_encoding('UTF-16LE').encode('UTF-8').delete("\x00")
      ensure
        TI.reg_close_key.call(hkey)
      end
    end

    def install_finalizer
      ptr = @ptr
      release_fn = W.vtable_function(ptr, 2, [W::VOIDP], W::DWORD)
      ObjectSpace.define_finalizer(self, self.class.finalizer(ptr, release_fn))
    end

    def self.finalizer(ptr, release_fn)
      proc { release_fn.call(ptr) unless ptr.zero? }
    end
  end
end
