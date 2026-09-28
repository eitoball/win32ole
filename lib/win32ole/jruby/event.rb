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
