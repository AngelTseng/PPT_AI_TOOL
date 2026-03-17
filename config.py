"""Shared runtime configuration for the PPT AI tool."""

import os

# Model name passed to the OpenAI-compatible client.
# In current setup, this project uses Gemini's OpenAI-compatible endpoint.
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gemini-3-flash-preview")