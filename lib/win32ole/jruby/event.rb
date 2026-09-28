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

    private

    # Built up across Tasks 9-13; a successful construction isn't
    # exercised by any test until Task 13 wires the real implementation in.
    def advise(ole, itf)
      raise NotImplementedError, 'advise is implemented in Task 13'
    end
  end
end
