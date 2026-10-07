"""Bind a model snapshot to each request without changing global defaults."""
from contextlib import contextmanager
from contextvars import ContextVar
from app.config import settings

_active_model = ContextVar('request_model', default=None)

def get_active_model():
    return _active_model.get() or settings.LLM_MODEL

@contextmanager
def use_model(model_id):
    token = _active_model.set(model_id)
    try:
        yield
    finally:
        _active_model.reset(token)

async def stream_with_model(model_id, factory):
    with use_model(model_id):
        generator = factory()
    try:
        while True:
            try:
                with use_model(model_id):
                    event = await generator.__anext__()
            except StopAsyncIteration:
                return
            yield event
    finally:
        with use_model(model_id):
            await generator.aclose()
