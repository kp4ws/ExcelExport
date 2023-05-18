class EventSystem:
    def __init__(self):
        self.listeners = {}

    def subscribe(self, event, listener):
        if event not in self.listeners:
            self.listeners[event] = []
        self.listeners[event].append(listener)

    def emit(self, event, data=None):
        if event in self.listeners:
            for listener in self.listeners[event]:
                listener(data)