"""Background generator subscriptions and the latest RTD update for each topic."""
import asyncio
from collections import OrderedDict
from contextlib import aclosing, closing
import inspect
import json
import threading
import traceback


class Streams:
    def __init__(self, before_next=lambda: None):
        self.before_next = before_next
        self.condition = threading.Condition()
        self.active = {}
        self.pending = OrderedDict()
        self.loop = None
        self.async_count = 0

    def open(self, key, generator, encode):
        with self.condition:
            self.active[key] = None
            try:
                if inspect.isasyncgen(generator):
                    if self.loop is None:
                        loop = asyncio.new_event_loop()
                        threading.Thread(target=self._run_loop, args=(loop,), daemon=True).start()
                        self.loop = loop
                    self.loop.call_soon_threadsafe(self._start_async, key, generator, encode)
                    self.async_count += 1
                else:
                    threading.Thread(target=self._run_sync, args=(key, generator, encode), daemon=True).start()
            except BaseException:
                del self.active[key]
                raise

    @staticmethod
    def _run_loop(loop):
        asyncio.set_event_loop(loop)
        with asyncio.Runner(loop_factory=lambda: loop) as runner:
            runner.get_loop().run_forever()

    def _start_async(self, key, generator, encode):
        with self.condition:
            task = self.loop.create_task(self._run_async(key, generator, encode))
            if key in self.active:
                self.active[key] = task
            task.add_done_callback(lambda finished: self._async_done(key))

    def _async_done(self, key):
        with self.condition:
            if key in self.active:
                self.active[key] = None
            self.async_count -= 1
            if self.async_count == 0:
                self.loop.stop()
                self.loop = None

    def close(self, key):
        with self.condition:
            task = self.active.pop(key, None)
            if task is not None and not task.done():
                task.get_loop().call_soon_threadsafe(task.cancel)
            self.pending.pop(key, None)
            self.condition.notify()

    def _publish(self, key, **update):
        with self.condition:
            if key in self.active:
                self.pending[key] = {"id": key, **update}
                self.condition.notify()

    def _yield(self, key, value, encode):
        # Snapshot mutable yields before advancing. A bad value affects only its own topic.
        self._publish(key, json=json.dumps(encode(value), allow_nan=False))

    def _run_sync(self, key, generator, encode):
        try:
            empty = True
            with closing(generator):
                while key in self.active:
                    try:
                        self.before_next()
                        value = next(generator)
                    except StopIteration:
                        break
                    self._yield(key, value, encode)
                    empty = False
            if empty:
                self._publish(key, json="null")
        except BaseException:
            self._publish(key, failure=traceback.format_exc())

    async def _run_async(self, key, generator, encode):
        try:
            empty = True
            async with aclosing(generator):
                if key not in self.active:
                    return
                # Once running, close() schedules cancellation on this loop. Let that
                # unwind the generator before aclose starts its async cleanup.
                while True:
                    try:
                        self.before_next()
                        value = await anext(generator)
                    except StopAsyncIteration:
                        break
                    self._yield(key, value, encode)
                    empty = False
                    # A generator can yield indefinitely without an await of its own.
                    await asyncio.sleep(0)
            if empty:
                self._publish(key, json="null")
        except BaseException:
            self._publish(key, failure=traceback.format_exc())

    def read(self):
        with self.condition:
            self.condition.wait_for(lambda: self.pending or not self.active)
            if self.pending:
                return self.pending.popitem(last=False)[1]
            return None
