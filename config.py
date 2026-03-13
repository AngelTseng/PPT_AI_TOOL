"""Shared runtime configuration for the PPT AI tool."""

import os

OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-4.1-mini")
