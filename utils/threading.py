"""
COM-safe threading utilities for Office automation.

CRITICAL: COM objects MUST be created and used within the same thread.
This module provides a worker thread pattern that ensures COM safety.
"""
import threading
import queue
from typing import Callable, Any, Optional
from utils.logging import get_logger

logger = get_logger(__name__)


class ConversionWorker:
    """
    A worker thread that processes conversion jobs in a COM-safe manner.
    
    The worker thread creates COM objects internally and processes jobs
    from a queue, ensuring thread safety for Office automation.
    """
    
    def __init__(self):
        self._queue: queue.Queue = queue.Queue()
        self._thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._running = False
        
    def start(self):
        """Start the worker thread."""
        if self._running:
            logger.warning("Worker already running")
            return
            
        self._stop_event.clear()
        self._thread = threading.Thread(target=self._worker_loop, daemon=True)
        self._thread.start()
        self._running = True
        logger.info("Conversion worker started")
        
    def stop(self, timeout: float = 5.0):
        """
        Stop the worker thread gracefully.
        
        Args:
            timeout: Maximum time to wait for thread to finish
        """
        if not self._running:
            return
            
        self._stop_event.set()
        if self._thread:
            self._thread.join(timeout=timeout)
        self._running = False
        logger.info("Conversion worker stopped")
        
    def submit(self, task: Callable, callback: Callable[[Any], None] = None):
        """
        Submit a task to the worker queue.
        
        Args:
            task: Callable that performs the conversion (must be COM-safe)
            callback: Optional callback to invoke with the result (runs on worker thread)
        """
        self._queue.put((task, callback))
        logger.debug(f"Task submitted to worker queue (queue size: {self._queue.qsize()})")
        
    def _worker_loop(self):
        """Main worker loop (runs in separate thread)."""
        logger.debug("Worker loop started")
        
        while not self._stop_event.is_set():
            try:
                # Wait for task with timeout to allow checking stop event
                task, callback = self._queue.get(timeout=0.5)
                
                try:
                    logger.debug("Executing task")
                    result = task()
                    
                    if callback:
                        callback(result)
                        
                except Exception as e:
                    logger.error(f"Task execution failed: {e}", exc_info=True)
                    if callback:
                        callback(e)
                finally:
                    self._queue.task_done()
                    
            except queue.Empty:
                continue
                
        logger.debug("Worker loop exited")
