#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Shared, lazily-created Anthropic client used by every LLM-backed check.

Tests patch ``llm.client.get_client`` to intercept all API calls, so callers must
reach it through the module (``from llm import client as llm_client``), not by
importing the function by name.
"""

import os

DEFAULT_MODEL = 'claude-sonnet-5-5'       # deep analysis (was gpt-4o)
DEFAULT_FAST_MODEL = 'claude-haiku-5-5'   # simple extraction (was gpt-4o-mini)


def get_client(api_key: str):
    """Return an Anthropic client. The SDK is imported lazily so the app (and the
    no-key pattern-matching fallback) still works if it isn't installed."""
    import anthropic
    return anthropic.Anthropic(api_key=api_key)


def get_model() -> str:
    return os.environ.get('CLAUDE_MODEL', DEFAULT_MODEL)


def get_fast_model() -> str:
    return os.environ.get('CLAUDE_FAST_MODEL', DEFAULT_FAST_MODEL)


def json_output_config(schema: dict) -> dict:
    """output_config that makes Claude return JSON matching `schema` (structured outputs)."""
    return {'format': {'type': 'json_schema', 'schema': schema}}


def response_text(resp, *, require_complete: bool = False) -> str:
    """Join the text blocks of a Messages API response.

    Raises RuntimeError on a refusal or an empty response. With require_complete
    (use it for JSON output) it also raises if the output hit max_tokens, since
    truncated JSON is unusable.
    """
    stop_reason = getattr(resp, 'stop_reason', None)
    if stop_reason == 'refusal':
        raise RuntimeError('Claude declined to answer this request')
    if require_complete and stop_reason == 'max_tokens':
        raise RuntimeError('Claude response was cut off at max_tokens')
    text = ''.join(b.text for b in resp.content if getattr(b, 'type', None) == 'text')
    if not text.strip():
        raise RuntimeError('Claude returned no text content')
    return text
