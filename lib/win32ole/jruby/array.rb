# lib/win32ole/jruby/array.rb
require 'fiddle'
require 'win32ole/jruby/win32'

class WIN32OLE
  module SafeArray
    W = Win32
    private_constant :W

    module_function

    def oleaut32
      W.oleaut32
    end

    def safe_array_create
      @safe_array_create ||= Fiddle::Function.new(
        oleaut32['SafeArrayCreate'], [W::WORD, W::DWORD, W::VOIDP], W::VOIDP, W::STDCALL
      )
    end

    def safe_array_create_vector
      @safe_array_create_vector ||= Fiddle::Function.new(
        oleaut32['SafeArrayCreateVector'], [W::WORD, W::LONG, W::DWORD], W::VOIDP, W::STDCALL
      )
    end

    def safe_array_destroy
      @safe_array_destroy ||= Fiddle::Function.new(oleaut32['SafeArrayDestroy'], [W::VOIDP], W::LONG, W::STDCALL)
    end

    def safe_array_lock
      @safe_array_lock ||= Fiddle::Function.new(oleaut32['SafeArrayLock'], [W::VOIDP], W::LONG, W::STDCALL)
    end

    def safe_array_unlock
      @safe_array_unlock ||= Fiddle::Function.new(oleaut32['SafeArrayUnlock'], [W::VOIDP], W::LONG, W::STDCALL)
    end

    def safe_array_get_dim
      @safe_array_get_dim ||= Fiddle::Function.new(oleaut32['SafeArrayGetDim'], [W::VOIDP], W::DWORD, W::STDCALL)
    end

    def safe_array_get_lbound
      @safe_array_get_lbound ||= Fiddle::Function.new(
        oleaut32['SafeArrayGetLBound'], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_get_ubound
      @safe_array_get_ubound ||= Fiddle::Function.new(
        oleaut32['SafeArrayGetUBound'], [W::VOIDP, W::DWORD, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_ptr_of_index
      @safe_array_ptr_of_index ||= Fiddle::Function.new(
        oleaut32['SafeArrayPtrOfIndex'], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_put_element
      @safe_array_put_element ||= Fiddle::Function.new(
        oleaut32['SafeArrayPutElement'], [W::VOIDP, W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_access_data
      @safe_array_access_data ||= Fiddle::Function.new(
        oleaut32['SafeArrayAccessData'], [W::VOIDP, W::VOIDP], W::LONG, W::STDCALL
      )
    end

    def safe_array_unaccess_data
      @safe_array_unaccess_data ||= Fiddle::Function.new(
        oleaut32['SafeArrayUnaccessData'], [W::VOIDP], W::LONG, W::STDCALL
      )
    end
  end
end
