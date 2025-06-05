# core/events.py - Sistema de eventos
from typing import Callable, Dict, List
import threading


class EventManager:
    def __init__(self):
        self._listeners: Dict[str, List[Callable]] = {}
        self._lock = threading.Lock()

    def subscribe(self, event_type: str, callback: Callable):
        """Suscribe un callback a un tipo de evento"""
        with self._lock:
            if event_type not in self._listeners:
                self._listeners[event_type] = []
            self._listeners[event_type].append(callback)

    def emit(self, event_type: str, *args, **kwargs):
        """Emite un evento a todos los listeners suscritos"""
        with self._lock:
            listeners = self._listeners.get(event_type, [])

        for callback in listeners:
            try:
                callback(*args, **kwargs)
            except Exception as e:
                print(f"Error en callback de evento {event_type}: {e}")
