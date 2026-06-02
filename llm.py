"""
llm.py — Provider-abstracted LLM access for SheetMind.

Primary provider is OpenRouter (https://openrouter.ai), which exposes an
OpenAI-compatible chat-completions API to 400+ models. SheetMind defaults to
Moonshot AI's **Kimi K2.6** — a trillion-parameter open MoE model purpose-built
for agentic, tool-calling workflows — and transparently falls back across a
configurable list of models when one is unavailable / rate-limited.

The module exposes a single high-level helper, ``chat()``, that speaks the
OpenAI tool-calling protocol. The command engine (engine.py) builds on top of
it; nothing here knows anything about spreadsheets.

Configuration (read from environment / .env):
    OPENROUTER_API_KEY        required to use the LLM at all
    SHEETMIND_MODEL           primary model slug  (default: moonshotai/kimi-k2.6)
    SHEETMIND_MODEL_FALLBACKS comma-separated fallback slugs
    SHEETMIND_LLM_BASE_URL    override base URL (default OpenRouter)
    SHEETMIND_LLM_TIMEOUT     per-request timeout in seconds (default 60)
"""
from __future__ import annotations

import os
import time
import json
from typing import Optional

import requests

DEFAULT_BASE_URL = "https://openrouter.ai/api/v1"
DEFAULT_MODEL = "moonshotai/kimi-k2.6"
# Sensible, capable open-weight fallbacks (all support tool calling on OpenRouter).
DEFAULT_FALLBACKS = [
    "moonshotai/kimi-k2-0905",
    "deepseek/deepseek-chat-v3.1:free",
    "google/gemma-4-31b-it:free",
    "google/gemma-4-26b-a4b-it:free",
]


class LLMError(RuntimeError):
    """Raised when every configured model fails."""


class LLMNotConfigured(LLMError):
    """Raised when no API key is present — lets callers fall back to local parsing."""


def _api_key() -> str:
    key = (
        os.environ.get("OPENROUTER_API_KEY")
        or os.environ.get("SHEETMIND_LLM_API_KEY")
        or ""
    ).strip()
    if not key:
        raise LLMNotConfigured(
            "OPENROUTER_API_KEY missing from .env — add a free key from "
            "https://openrouter.ai/keys to enable AI commands."
        )
    return key


def _base_url() -> str:
    return os.environ.get("SHEETMIND_LLM_BASE_URL", DEFAULT_BASE_URL).rstrip("/")


def model_chain() -> list[str]:
    """Ordered list of model slugs to try: primary first, then fallbacks."""
    primary = os.environ.get("SHEETMIND_MODEL", DEFAULT_MODEL).strip()
    raw = os.environ.get("SHEETMIND_MODEL_FALLBACKS", "").strip()
    fallbacks = [m.strip() for m in raw.split(",") if m.strip()] or DEFAULT_FALLBACKS
    chain: list[str] = []
    for m in [primary, *fallbacks]:
        if m and m not in chain:
            chain.append(m)
    return chain


def is_configured() -> bool:
    try:
        _api_key()
        return True
    except LLMNotConfigured:
        return False


def chat(
    messages: list[dict],
    tools: Optional[list[dict]] = None,
    tool_choice: Optional[object] = None,
    temperature: float = 0.0,
    max_tokens: int = 2048,
    models: Optional[list[str]] = None,
) -> dict:
    """
    Send a chat-completion request and return the assistant message dict
    (``{"role": "assistant", "content": ..., "tool_calls": [...]}``).

    Tries each model in the chain, retrying transient errors (429/503) with
    exponential backoff. Raises LLMError if all models fail.
    """
    key = _api_key()
    url = f"{_base_url()}/chat/completions"
    timeout = float(os.environ.get("SHEETMIND_LLM_TIMEOUT", "60"))
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        # Optional attribution headers OpenRouter recommends — harmless elsewhere.
        "HTTP-Referer": "https://github.com/sheetmind",
        "X-Title": "SheetMind",
    }

    chain = models or model_chain()
    last_err = "no models configured"

    for model in chain:
        payload: dict = {
            "model": model,
            "messages": messages,
            "temperature": temperature,
            "max_tokens": max_tokens,
        }
        if tools:
            payload["tools"] = tools
            payload["tool_choice"] = tool_choice or "auto"

        for attempt in range(3):
            try:
                resp = requests.post(url, headers=headers, json=payload, timeout=timeout)
            except requests.RequestException as e:
                last_err = f"{model}: network error: {e}"
                time.sleep(1.5 * (attempt + 1))
                continue

            if resp.status_code == 200:
                data = resp.json()
                choices = data.get("choices") or []
                if not choices:
                    last_err = f"{model}: empty response ({json.dumps(data)[:200]})"
                    break
                return choices[0].get("message", {}) or {}

            # Transient — retry this model with backoff.
            if resp.status_code in (429, 500, 502, 503):
                last_err = f"{model}: HTTP {resp.status_code}"
                time.sleep(2 ** attempt)
                continue

            # Model-specific hard failure (bad slug, unsupported) — try next model.
            last_err = f"{model}: HTTP {resp.status_code}: {resp.text[:200]}"
            break

    raise LLMError(f"All models failed. Last error: {last_err}")
