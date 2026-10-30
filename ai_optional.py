"""OpenAI-only adapter, OFF by default. No AI requests during document generation."""
from __future__ import annotations
import os
from typing import Any

def request_summary(text: str, *, enabled: bool=False, model: str | None=None) -> dict[str,Any]:
    """Explicitly opt-in. Returns response + actual token counts. Never logs raw content."""
    if not enabled:
        raise RuntimeError("AI is disabled by default. Pass enabled=True explicitly.")
    if not os.getenv("OPENAI_API_KEY"):
        raise RuntimeError("OPENAI_API_KEY is missing.")
    from openai import OpenAI  # optional install: pip install -e '.[ai]'
    model = model or os.getenv("OPENAI_MODEL")
    if not model:
        raise RuntimeError("Set an explicit supported model or OPENAI_MODEL.")
    client=OpenAI()
    answer=client.responses.create(
        model=model,
        store=False,
        instructions=("סכם את הטקסט בעברית קצרה, אל תמציא חומר חסר, "
                      "סמן אי-ודאות וטעויות אפשריות לבדיקה ידנית."),
        input=text[:20000],
    )
    usage=answer.usage
    return {
        "text": answer.output_text,
        "input_tokens": int(usage.input_tokens) if usage else None,
        "output_tokens": int(usage.output_tokens) if usage else None,
        "total_tokens": int(usage.total_tokens) if usage else None,
        "model": model
    }
