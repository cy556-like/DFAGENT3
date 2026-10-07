"""Offline regression tests; never make a paid provider request."""
import ast
import asyncio
import importlib
import os
from pathlib import Path
import sys
import tempfile
import threading
import time
from types import SimpleNamespace
import unittest
from unittest.mock import patch

ROOT = Path(os.environ.get('RELATED_SOURCE', Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(ROOT))
TEMP = tempfile.TemporaryDirectory(prefix='related-model-tests-')
os.environ.update(DATA_DIR=TEMP.name, DEEPSEEK_API_KEY='test-ark', QWEN_API_KEY='test-qwen',
                  MOONSHOT_API_KEY='test-kimi', MIMO_API_KEY='test-mimo', DEEPSEEK_V41_API_KEY='test-v41',
                  LLM_API_KEY='test-aliyun', LLM_BASE_URL='https://dashscope.aliyuncs.com/compatible-mode/v1',
                  LLM_API_KEY_BACKUP='test-old-zhipu', LLM_BASE_URL_BACKUP='https://open.bigmodel.cn/api/anthropic')
for key in ('LLM_MODEL', 'ARK_CODING_API_KEY', 'ARK_CODING_BASE_URL'):
    os.environ.pop(key, None)
from app import config
from app.agent.model_routing import RoutedLLM

QZ = hasattr(config, 'JSXF_DEMO_LOGIN_ENABLED') or hasattr(config.settings, 'JSXF_DEMO_LOGIN_ENABLED')
SB = hasattr(config, 'get_user_model')
resolve = getattr(config, 'resolve_model_id', None) or config.resolve_effective_model


class Failure(Exception):
    def __init__(self, status, message):
        self.status_code = status
        super().__init__(message)


def factory_scope():
    tree = ast.parse((ROOT / 'app/agent/core.py').read_text(encoding='utf-8-sig'))
    factory = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == '_create_llm_client')
    scope = {**vars(config), 'time': time, 'threading': threading, '_llm_cache': {},
             '_LLM_CACHE_TTL': 900, '_primary_key_failed': True, '_primary_key_lock': threading.Lock(),
             'MANUAL_IDLE_TIMEOUT_SECONDS': 1800, 'get_active_model': lambda: config.settings.LLM_MODEL,
             'ChatOpenAI': lambda **options: SimpleNamespace(**options)}
    exec(compile(ast.Module(body=[factory], type_ignores=[]), '<actual-factory>', 'exec'), scope)
    return scope


class RoutingTests(unittest.TestCase):
    def setUp(self):
        self.calls = []
        self.failure = Failure(401, 'invalid api key')

    def factory(self, model_override=None, **options):
        self.calls.append((model_override, options))
        outer = self
        class Client:
            def invoke(self, messages, **kwargs):
                if model_override == 'glm-5.3':
                    raise outer.failure
                return 'good'
            async def ainvoke(self, messages, **kwargs):
                return self.invoke(messages, **kwargs)
            async def astream(self, messages, **kwargs):
                yield SimpleNamespace(content=self.invoke(messages, **kwargs), tool_call_chunks=[])
        return Client()

    def test_catalog_migrated(self):
        ids = {m['id'] for m in config.AVAILABLE_MODELS}
        self.assertIn('doubao-seed-2.1-pro', ids)
        self.assertIn('glm-5.3', ids)
        self.assertNotIn('Doubao-Seed-2.0-pro', ids)
        self.assertNotIn('glm-5.2', ids)

    def test_old_aliases(self):
        self.assertEqual(config.normalize_model_id('Doubao-Seed-2.0-pro'), 'doubao-seed-2.1-pro')
        self.assertEqual(config.normalize_model_id('glm-5.2'), 'glm-5.3')
        self.assertEqual(config.normalize_model_id('glm-4v-flash'), 'glm-5.3-flash')

    def test_backup_migrated_without_using_old_credential(self):
        self.assertEqual(config.get_model_connection(config.ARK_BACKUP_MODEL, backup=True), ('test-ark', config.ARK_CHAT_BASE_URL))

    def test_vision_pair(self):
        self.assertEqual(config.DEFAULT_VISION_MODEL, 'glm-5.3-flash')
        self.assertEqual(config.get_model_connection(config.DEFAULT_VISION_MODEL), ('test-ark', config.ARK_CHAT_BASE_URL))
        self.assertEqual(RoutedLLM(self.factory, config.DEFAULT_VISION_MODEL, {}).candidates(), [config.DEFAULT_VISION_MODEL])

    def test_model_provider_pairs(self):
        self.assertEqual(config.get_model_connection('qwen3.7-plus')[0], 'test-qwen')
        self.assertEqual(config.get_model_connection('doubao-seed-2.1-pro')[0], 'test-ark')
        self.assertEqual(config.get_model_connection('DeepSeek-V4.1-Flash')[0], 'test-v41')

    def test_peak_boundaries(self):
        from datetime import datetime
        for hour, minute, expected in [(8, 59, 'DeepSeek-V4.1-Flash'), (9, 0, 'DeepSeek-V4-Flash'),
                                       (11, 59, 'DeepSeek-V4-Flash'), (12, 0, 'DeepSeek-V4.1-Flash'),
                                       (14, 0, 'DeepSeek-V4-Flash'), (18, 0, 'DeepSeek-V4.1-Flash')]:
            self.assertEqual(resolve('DeepSeek-V4.1-Flash', datetime(2026, 10, 8, hour, minute)), expected)

    def test_factory_never_reads_global_failure(self):
        client = factory_scope()['_create_llm_client'](model_override='qwen3.7-plus')
        self.assertEqual(client.api_key, 'test-qwen')
        self.assertEqual(client.model, 'qwen3.7-plus')

    def test_factory_explicit_backup_pair(self):
        client = factory_scope()['_create_llm_client'](model_override='qwen3.7-plus', force_backup=True)
        self.assertEqual(client.api_key, 'test-ark')
        self.assertEqual(client.model, 'glm-5.3')
        self.assertEqual(client.base_url, config.ARK_CHAT_BASE_URL)

    def test_factory_keeps_resolved_snapshot(self):
        client = factory_scope()['_create_llm_client'](model_override='DeepSeek-V4.1-Flash', resolved_model=True)
        self.assertEqual(client.api_key, 'test-v41')
        self.assertEqual(client.model, config.settings.DEEPSEEK_V41_MODEL)

    def test_auth_same_request_fallback_no_duplicate_credentials(self):
        llm = RoutedLLM(self.factory, 'glm-5.3', {})
        self.assertEqual(llm.invoke([]), 'good')
        self.assertEqual(len(self.calls), 2)
        self.assertNotEqual(self.calls[1][0], 'DeepSeek-V4-Flash')

    def test_async_auth_fallback(self):
        self.assertEqual(asyncio.run(RoutedLLM(self.factory, 'glm-5.3', {}).ainvoke([])), 'good')
        self.assertEqual(len(self.calls), 2)

    def test_stream_auth_fallback(self):
        async def run():
            return [chunk.content async for chunk in RoutedLLM(self.factory, 'glm-5.3', {}).astream([])]
        self.assertEqual(asyncio.run(run()), ['good'])
        self.assertEqual(len(self.calls), 2)

    def test_quota_can_try_other_model_same_key(self):
        self.failure = Failure(429, 'insufficient_quota')
        self.assertEqual(RoutedLLM(self.factory, 'glm-5.3', {}).invoke([]), 'good')
        self.assertEqual(self.calls[1][0], 'DeepSeek-V4-Flash')

    def test_timeout_never_switches(self):
        self.failure = TimeoutError('read timeout')
        with self.assertRaises(TimeoutError):
            RoutedLLM(self.factory, 'glm-5.3', {}).invoke([])
        self.assertEqual(len(self.calls), 1)

    def test_bad_request_never_switches(self):
        self.failure = Failure(400, 'bad tool schema')
        with self.assertRaises(Failure):
            RoutedLLM(self.factory, 'glm-5.3', {}).invoke([])
        self.assertEqual(len(self.calls), 1)

    def test_targets_deduplicate_backup(self):
        llm = RoutedLLM(self.factory, 'glm-5.3', {})
        ids = [(model, *llm.credentials(model, backup)) for model, backup in llm.targets()]
        self.assertEqual(len(ids), len(set(ids)))

    def test_distinct_backup_is_bound_to_ark_model(self):
        with patch.object(config.settings, 'LLM_API_KEY_BACKUP', 'test-different-ark'):
            llm = RoutedLLM(self.factory, 'qwen3.7-plus', {})
            self.assertEqual(llm.targets()[-1], ('glm-5.3', True))

    def test_partial_stream_never_retries(self):
        calls = []
        def factory(**options):
            calls.append(options)
            class Client:
                async def astream(self, *args, **kwargs):
                    yield SimpleNamespace(content='partial', tool_call_chunks=[])
                    raise Failure(429, 'insufficient_quota')
            return Client()
        async def run():
            return [chunk async for chunk in RoutedLLM(factory, 'glm-5.3', {}).astream([])]
        with self.assertRaises(Failure):
            asyncio.run(run())
        self.assertEqual(len(calls), 1)

    def test_partial_tool_stream_never_retries(self):
        calls = []
        def factory(**options):
            calls.append(options)
            class Client:
                async def astream(self, *args, **kwargs):
                    yield SimpleNamespace(content='', tool_call_chunks=[{'name': 'search'}])
                    raise Failure(429, 'insufficient_quota')
            return Client()
        async def run():
            return [chunk async for chunk in RoutedLLM(factory, 'glm-5.3', {}).astream([])]
        with self.assertRaises(Failure):
            asyncio.run(run())
        self.assertEqual(len(calls), 1)

    def test_preferences_two_accounts(self):
        if SB:
            save, load = config.set_user_model, config.get_user_model
        else:
            from app.model_preferences import save_user_model as save, load_user_model as load
        save('admin', 'glm-5.3')
        save('adminquanzhi', 'qwen3.7-plus')
        self.assertEqual(load('admin'), 'glm-5.3')
        self.assertEqual(load('adminquanzhi'), 'qwen3.7-plus')
        save('admin', 'Doubao-Seed-2.0-pro')
        self.assertEqual(load('admin'), 'doubao-seed-2.1-pro')

    def test_legacy_user_defaults_do_not_change_other_account(self):
        default = config.settings.LLM_MODEL
        if SB:
            config.set_user_model('one', 'glm-5.3')
        else:
            from app.model_preferences import save_user_model
            save_user_model('one', 'glm-5.3')
        self.assertEqual(config.settings.LLM_MODEL, default)

    @unittest.skipUnless(QZ, 'QZ long-generation policy')
    def test_long_manual_auth_does_not_switch(self):
        with self.assertRaises(Failure):
            RoutedLLM(self.factory, 'glm-5.3', {'long_generation': True}).invoke([])
        self.assertEqual(len(self.calls), 1)

    @unittest.skipUnless(QZ, 'QZ long-generation policy')
    def test_long_manual_plain_429_does_not_switch(self):
        self.failure = Failure(429, 'too many requests')
        with self.assertRaises(Failure):
            RoutedLLM(self.factory, 'glm-5.3', {'long_generation': True}).invoke([])
        self.assertEqual(len(self.calls), 1)

    @unittest.skipUnless(QZ, 'QZ long-generation policy')
    def test_long_manual_quota_switches(self):
        self.failure = Failure(429, 'insufficient_quota')
        self.assertEqual(RoutedLLM(self.factory, 'glm-5.3', {'long_generation': True}).invoke([]), 'good')

    @unittest.skipUnless(QZ, 'QZ long-generation policy')
    def test_manual_timeout_not_shortened(self):
        client = factory_scope()['_create_llm_client'](model_override='glm-5.3', long_generation=True)
        self.assertEqual(client.request_timeout, 1800)
        self.assertEqual(client.max_retries, 0)


if __name__ == '__main__':
    unittest.main(verbosity=2)
