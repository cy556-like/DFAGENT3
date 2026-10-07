"""Exercise production handlers, JWT and storage; only temporary local data."""
import ast
import asyncio
from concurrent.futures import ThreadPoolExecutor
import importlib.util
import logging
import os
from pathlib import Path
import shutil
import sys
import tempfile
import time
from types import SimpleNamespace
from typing import Optional
from unittest.mock import patch
from fastapi import APIRouter, Depends, FastAPI, File, Form, Header, HTTPException, Query, Request, UploadFile
from fastapi.responses import StreamingResponse
from fastapi.testclient import TestClient
from langchain_core.messages import HumanMessage
from pydantic import BaseModel

ROOT = Path(os.environ.get('RELATED_SOURCE', Path(__file__).resolve().parents[1]))
sys.path.insert(0, str(ROOT))
temporary = tempfile.TemporaryDirectory(prefix='related-api-isolation-')
os.environ.update(DATA_DIR=temporary.name, JWT_SECRET='offline-test-only', LLM_API_KEY='')
from app import config
from app.config import settings, AVAILABLE_MODELS
from app.memory import manager as memory
from app.auth import jwt_handler as jwt

source = ROOT / 'app/api/routes.py'
scope = dict(vars(memory), asyncio=asyncio, Depends=Depends, Header=Header, File=File, Form=Form,
             Query=Query, Request=Request, UploadFile=UploadFile, HTTPException=HTTPException,
             StreamingResponse=StreamingResponse, BaseModel=BaseModel, Optional=Optional,
             router=APIRouter(), get_username_from_token=jwt.get_username_from_token,
             shutil=shutil, settings=settings, time=time, os=os, record_session=lambda: None,
             cleanup_export_files=lambda **kwargs: 0, logger=logging.getLogger('isolation'),
             _request_model=lambda username, model_id=None: model_id or 'glm-5.3',
             _request_model_id=lambda *args: 'glm-5.3',
             _record_request=lambda *args, **kwargs: None)
names = {'get_current_user', 'require_auth', 'ensure_requested_username', 'ensure_chat_ownership',
         'ChatRequest', 'ChatResponse', 'RenameRequest', 'get_chats', 'create_chat_api', 'get_history',
         'delete_history', 'delete_chat_api', 'rename_chat_api', 'chat_api', 'chat_stream_api',
         'chat_with_file_stream', 'export_chat'}
if hasattr(config, 'get_user_model'):
    names.add('_request_auth_token')
    scope['_identity_username_active'] = lambda username: True
nodes = [node for node in ast.parse(source.read_text(encoding='utf-8-sig')).body
         if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)) and node.name in names]
assert {node.name for node in nodes} == names
exec(compile(ast.Module(body=nodes, type_ignores=[]), str(source), 'exec'), scope)
application = FastAPI()
application.include_router(scope['router'], prefix='/api/v1')
with TestClient(application) as client:
    accounts = ('admin', 'adminquanzhi')
    headers = {name: {'Authorization': 'Bearer ' + jwt.create_token(name, role='admin')} for name in accounts}
    chats = {}
    for name in accounts:
        response = client.post('/api/v1/chats', params={'username': name, 'agent_id': 'same-agent'}, headers=headers[name])
        assert response.status_code == 200, response.text
        chats[name] = response.json()['chat']['chat_id']
        history = memory.get_session_history(chats[name])
        history.add_message(HumanMessage(content='private-' + name))
        memory.flush_session(chats[name])
    assert chats['admin'] != chats['adminquanzhi']
    checks = 0
    for own, other in (accounts, accounts[::-1]):
        own_headers = headers[own]
        foreign = chats[other]
        cases = [
            ('GET', '/chats', {'params': {'username': other}}),
            ('POST', '/chats', {'params': {'username': other}}),
            ('GET', '/history/' + foreign, {}),
            ('DELETE', '/history/' + foreign, {}),
            ('DELETE', '/chats/' + foreign, {'params': {'username': own}}),
            ('PUT', '/chats/' + foreign + '/rename', {'json': {'username': own, 'chat_id': foreign, 'new_title': 'bad'}}),
            ('GET', '/export/' + foreign, {}),
            ('POST', '/chat', {'json': {'session_id': foreign, 'message': 'no model request'}}),
            ('POST', '/chat/stream', {'json': {'session_id': foreign, 'message': 'no model request'}}),
            ('POST', '/chat-with-file/stream', {'data': {'session_id': foreign}, 'files': {'file': ('test.py', b'test')}}),
            ('GET', '/history/default', {}),
        ]
        for method, path, arguments in cases:
            response = client.request(method, '/api/v1' + path, headers=own_headers, **arguments)
            assert response.status_code in (403, 404), (method, path, response.status_code, response.text)
            checks += 1
        listing = client.get('/api/v1/chats', params={'username': own}, headers=own_headers)
        assert [chat['chat_id'] for chat in listing.json()['chats']] == [chats[own]]
    def read_own(name):
        response = client.get('/api/v1/history/' + chats[name], headers=headers[name])
        assert response.status_code == 200, response.text
        assert response.json()['messages'][0]['content'] == 'private-' + name
    with ThreadPoolExecutor(max_workers=4) as pool:
        list(pool.map(read_own, accounts * 20))
    assert client.get('/api/v1/history/' + chats['admin']).status_code == 401
    assert client.get('/api/v1/history/' + chats['admin'], headers={'Authorization': 'Bearer invalid'}).status_code == 401
    for own in accounts:
        response = client.put('/api/v1/chats/' + chats[own] + '/rename', headers=headers[own],
                              json={'username': own, 'chat_id': chats['adminquanzhi' if own == 'admin' else 'admin'], 'new_title': 'bad'})
        assert response.status_code == 400, response.text
    retained = [memory.create_chat('admin', agent_id='same-agent')['chat_id'] for _ in range(4)]
    memory.get_session_history(retained[0]).add_message(HumanMessage(content='generation-progress'))
    memory.flush_session(retained[0])
    final = memory.create_chat('admin', agent_id='same-agent')['chat_id']
    assert set(retained + [chats['admin'], final]) <= {chat['chat_id'] for chat in memory.list_chats('admin')}
    assert memory.get_session_history(retained[0]).messages[0].content == 'generation-progress'
    print(ROOT.name, 'PASS:', checks, 'foreign/default-session checks; 40 parallel own-history reads; missing/invalid JWT; path/body ID mismatch')
    print('NO_SERVER_ACCESS; NO_PAID_MODEL_CALLS')
