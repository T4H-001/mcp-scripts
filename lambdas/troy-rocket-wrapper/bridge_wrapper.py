"""
/rocket Bridge Wrapper — Lambda v1.0
Deployed as: troy-rocket-wrapper
Role: intercept all LLM task dispatches, apply canonical contract, log execution.

Environment variables required:
  SUPABASE_URL           - https://lzfgigiyqpuuxslsygjt.supabase.co
  SUPABASE_SERVICE_KEY   - service_role key (from cap_secrets)
  DEFAULT_CONTRACT_MODE  - rocket (default)
  CONTRACT_CACHE_TTL_S   - 300 (5 min cache, avoid Supabase round-trips)
"""

import json
import os
import time
import uuid
import logging
from functools import lru_cache
from typing import Optional

import boto3
import httpx

logger = logging.getLogger()
logger.setLevel(logging.INFO)

SUPABASE_URL = os.environ["SUPABASE_URL"].rstrip("/")
SUPABASE_KEY = os.environ["SUPABASE_SERVICE_KEY"]
DEFAULT_MODE  = os.environ.get("DEFAULT_CONTRACT_MODE", "rocket")
CACHE_TTL     = int(os.environ.get("CONTRACT_CACHE_TTL_S", "300"))

# ─── Contract cache ────────────────────────────────────────────────────────────

_contract_cache: dict = {}
_cache_ts: float = 0.0

def get_active_contract(mode: str = "rocket") -> dict:
    """Fetch contract from Supabase with TTL cache."""
    global _contract_cache, _cache_ts

    now = time.time()
    cache_key = f"contract:{mode}"

    if cache_key in _contract_cache and (now - _cache_ts) < CACHE_TTL:
        return _contract_cache[cache_key]

    resp = httpx.post(
        f"{SUPABASE_URL}/rest/v1/rpc/rpc_get_active_contract",
        headers={
            "apikey":        SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type":  "application/json",
        },
        json={"p_mode": mode},
        timeout=5.0,
    )
    resp.raise_for_status()
    rows = resp.json()
    if not rows:
        raise ValueError(f"No active contract found for mode={mode}")

    contract = rows[0]
    _contract_cache[cache_key] = contract
    _cache_ts = now
    logger.info(f"[CONTRACT] loaded mode={mode} version={contract['version']}")
    return contract


def get_adapter(provider: str) -> dict:
    """Fetch LLM adapter config for provider."""
    resp = httpx.get(
        f"{SUPABASE_URL}/rest/v1/llm_contract_adapter",
        headers={
            "apikey":        SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
        },
        params={
            "provider":  f"eq.{provider}",
            "model":     "eq.__all__",
            "is_active": "eq.true",
            "select":    "*",
            "limit":     "1",
        },
        timeout=5.0,
    )
    resp.raise_for_status()
    rows = resp.json()
    return rows[0] if rows else {"prepend_style": "inline", "supports_hard_mode": True}


# ─── Core logic ────────────────────────────────────────────────────────────────

def detect_header_mode(task_payload: str) -> str:
    stripped = task_payload.strip()
    if stripped.startswith("/rocket-hard"):
        return "explicit_rocket_hard"
    if stripped.startswith("/rocket"):
        return "explicit_rocket"
    return "default_rocket"


def build_effective_prompt(task_payload: str, contract: dict, adapter: dict) -> tuple[str, str]:
    """
    Returns (effective_prompt, header_mode).
    Passthrough if task already carries /rocket header.
    """
    mode = detect_header_mode(task_payload)
    task = task_payload.strip()

    if mode in ("explicit_rocket", "explicit_rocket_hard"):
        return task, mode   # passthrough — respect what was sent

    # Auto-apply default contract
    header = contract["body"]

    if adapter.get("prepend_style") == "split":
        # Cursor / split-mode: system prompt separate
        effective = f"[SYSTEM]\n{header}\n\n[USER]\n{task}"
    else:
        effective = f"{header}\n\n{task}"

    return effective, "default_rocket"


def log_execution(payload: dict) -> Optional[str]:
    """Fire-and-forget execution log to Supabase."""
    try:
        resp = httpx.post(
            f"{SUPABASE_URL}/rest/v1/rpc/rpc_log_contract_execution",
            headers={
                "apikey":        SUPABASE_KEY,
                "Authorization": f"Bearer {SUPABASE_KEY}",
                "Content-Type":  "application/json",
            },
            json=payload,
            timeout=3.0,
        )
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        logger.warning(f"[LOG] execution log failed: {e}")
        return None


# ─── Lambda handler ────────────────────────────────────────────────────────────

def lambda_handler(event: dict, context) -> dict:
    """
    Expected event shape:
    {
      "fn": "troy-rocket-wrapper",
      "task": "<raw task payload>",
      "provider": "claude",          // claude | gpt | gemini | perplexity | cursor
      "model": "claude-sonnet-4-6",  // optional
      "mode": "rocket",              // optional: rocket | rocket-hard
      "task_id": "<uuid>"            // optional: link to ai_task_queue
    }

    Returns:
    {
      "effective_prompt": "...",
      "header_mode": "...",
      "contract_name": "...",
      "contract_version": "...",
      "log_id": "..."
    }
    """
    t_start = time.time()

    raw_task  = event.get("task", "").strip()
    provider  = event.get("provider", "bridge").lower()
    model     = event.get("model")
    req_mode  = event.get("mode", DEFAULT_MODE)
    task_id   = event.get("task_id")

    if not raw_task:
        return {"error": "task field is required", "statusCode": 400}

    try:
        contract = get_active_contract(req_mode)
        adapter  = get_adapter(provider)

        effective_prompt, header_mode = build_effective_prompt(raw_task, contract, adapter)

        execution_ms = int((time.time() - t_start) * 1000)

        log_id = log_execution({
            "p_contract_name":    contract["contract_name"],
            "p_contract_version": contract["version"],
            "p_provider":         provider,
            "p_model":            model,
            "p_raw_prompt":       raw_task[:4000],        # truncate for storage
            "p_effective_prompt": effective_prompt[:4000],
            "p_header_mode":      header_mode,
            "p_dry_run_status":   "skipped",
            "p_writeback_status": "not_applicable",
            "p_execution_ms":     execution_ms,
            "p_outcome":          "success",
            "p_task_id":          task_id,
        })

        logger.info(f"[ROCKET] provider={provider} mode={header_mode} contract={contract['contract_name']} v={contract['version']}")

        return {
            "statusCode":       200,
            "effective_prompt": effective_prompt,
            "header_mode":      header_mode,
            "contract_name":    contract["contract_name"],
            "contract_version": contract["version"],
            "log_id":           log_id,
        }

    except Exception as e:
        logger.error(f"[ROCKET ERROR] {e}", exc_info=True)
        execution_ms = int((time.time() - t_start) * 1000)
        log_execution({
            "p_contract_name":    req_mode,
            "p_contract_version": "unknown",
            "p_provider":         provider,
            "p_model":            model,
            "p_raw_prompt":       raw_task[:4000],
            "p_header_mode":      "none",
            "p_execution_ms":     execution_ms,
            "p_outcome":          "failure",
            "p_error_detail":     str(e),
            "p_task_id":          task_id,
        })
        return {"statusCode": 500, "error": str(e)}


# ─── Local test ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import os
    os.environ.setdefault("SUPABASE_URL",         "https://lzfgigiyqpuuxslsygjt.supabase.co")
    os.environ.setdefault("SUPABASE_SERVICE_KEY", "YOUR_SERVICE_KEY")

    result = lambda_handler({
        "fn":       "troy-rocket-wrapper",
        "task":     "Audit the 28 businesses for operational blockers.",
        "provider": "claude",
        "model":    "claude-sonnet-4-6",
    }, None)

    print(json.dumps(result, indent=2))
