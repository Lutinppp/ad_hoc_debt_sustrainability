"""Configuration for IMF API access and output behavior."""

from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Dict


@dataclass(frozen=True)
class IMFConfig:
    """Configuration object used by the IMF API client.

    Paste your IMF API key/token below if your environment requires authenticated access.
    Environment variables take precedence when set.
    """

    base_url: str = "https://www.imf.org/external/datamapper/api/v1"
    timeout_seconds: int = 60

    # Optional credentials: paste here if needed.
    api_key: str = "6d2a81abe2884cc38950d3391e337860"
    bearer_token: str = ""

    @classmethod
    def from_env(cls) -> "IMFConfig":
        return cls(
            base_url=os.getenv("IMF_BASE_URL", cls.base_url),
            timeout_seconds=int(os.getenv("IMF_TIMEOUT_SECONDS", str(cls.timeout_seconds))),
            api_key=os.getenv("IMF_API_KEY", cls.api_key),
            bearer_token=os.getenv("IMF_BEARER_TOKEN", cls.bearer_token),
        )

    def headers(self) -> Dict[str, str]:
        headers: Dict[str, str] = {"Accept": "application/json"}
        if self.api_key:
            headers["x-api-key"] = self.api_key
        if self.bearer_token:
            headers["Authorization"] = f"Bearer {self.bearer_token}"
        return headers
