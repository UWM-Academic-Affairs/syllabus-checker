"""Tests for the shared Anthropic client helpers in llm/client.py."""

from unittest.mock import MagicMock

import pytest

from llm import client as llm_client


def _resp(blocks, stop_reason='end_turn'):
    r = MagicMock()
    r.content = blocks
    r.stop_reason = stop_reason
    return r


def _block(type_, text=''):
    b = MagicMock()
    b.type = type_
    b.text = text
    return b


def test_model_defaults(monkeypatch):
    monkeypatch.delenv('CLAUDE_MODEL', raising=False)
    monkeypatch.delenv('CLAUDE_FAST_MODEL', raising=False)
    assert llm_client.get_model() == 'claude-sonnet-5-5'
    assert llm_client.get_fast_model() == 'claude-haiku-5-5'


def test_model_env_overrides(monkeypatch):
    monkeypatch.setenv('CLAUDE_MODEL', 'a')
    monkeypatch.setenv('CLAUDE_FAST_MODEL', 'b')
    assert (llm_client.get_model(), llm_client.get_fast_model()) == ('a', 'b')


def test_response_text_joins_only_text_blocks():
    resp = _resp([_block('thinking', 'secret'), _block('text', 'Hello '), _block('text', 'world')])
    assert llm_client.response_text(resp) == 'Hello world'


def test_response_text_raises_on_empty_and_refusal():
    with pytest.raises(RuntimeError):
        llm_client.response_text(_resp([_block('thinking', 'x')]))
    with pytest.raises(RuntimeError):
        llm_client.response_text(_resp([_block('text', 'hi')], stop_reason='refusal'))


def test_require_complete_rejects_truncated_output():
    resp = _resp([_block('text', '{"a":')], stop_reason='max_tokens')
    assert llm_client.response_text(resp) == '{"a":'          # prose: partial text is kept
    with pytest.raises(RuntimeError):
        llm_client.response_text(resp, require_complete=True)  # JSON: unusable


def test_json_output_config_shape():
    schema = {'type': 'object'}
    assert llm_client.json_output_config(schema) == {
        'format': {'type': 'json_schema', 'schema': schema}}


def test_get_client_builds_sdk_client_with_given_key():
    c = llm_client.get_client('test-key')
    assert c.api_key == 'test-key'
