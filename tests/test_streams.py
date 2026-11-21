import asyncio
import json
from pathlib import Path
import sys
import threading
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "python"))
from _xll_host import _encode, _register
from _xll_streams import Streams
from xll_python import function


class StreamsTest(unittest.TestCase):
    def setUp(self):
        self.streams = Streams()

    def start(self, key, gen):
        self.streams.open(key, gen, _encode)
        self.addCleanup(self.streams.close, key)

    def update(self):
        with self.streams.condition:
            self.assertTrue(self.streams.condition.wait_for(lambda: self.streams.pending, timeout=3))
            return self.streams.read()

    def test_close_during_blocking_next_discards_late_yield(self):
        waiting, resume, closed = (threading.Event() for _ in range(3))

        def values():
            try:
                yield 1
                waiting.set()
                resume.wait(3)
                yield 2
            finally:
                closed.set()

        self.start("sync", values())
        self.assertTrue(waiting.wait(3))
        self.assertEqual(self.update(), {"id": "sync", "json": "1"})
        self.streams.close("sync")
        resume.set()
        self.assertTrue(closed.wait(3))
        self.assertIsNone(self.streams.read())

    def test_async_disconnect_interrupts_await_and_runs_async_cleanup(self):
        closed = threading.Event()

        async def values():
            try:
                yield [[1, 2]]
                await asyncio.Event().wait()
            finally:
                await asyncio.sleep(0)
                closed.set()

        self.start("async", values())
        self.assertEqual(json.loads(self.update()["json"]), [[1, 2]])
        self.streams.close("async")
        self.assertTrue(closed.wait(3))

    def test_empty_and_failed_streams_are_independent(self):
        self.start("bad", (x for x in [object()]))
        self.start("empty", (x for x in []))
        updates = {msg["id"]: msg for msg in (self.update(), self.update())}
        self.assertIn("not JSON serializable", updates["bad"]["failure"])
        self.assertIsNone(json.loads(updates["empty"]["json"]))

    def test_invalid_generator_options(self):
        def values():
            yield 1

        for fn, opts in [(values, {"generator_instances": "cell"}),
                         (values, {"asynchronous": False}),
                         (lambda: 1, {"name": "ordinary", "generator_instances": "many"})]:
            with self.subTest(opts=opts), self.assertRaises(ValueError):
                _register(function(**opts)(fn), {}, [])


if __name__ == "__main__":
    unittest.main()
