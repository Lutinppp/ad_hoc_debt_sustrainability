"""IMF DataMapper API client."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable

import requests

from .config import IMFConfig


@dataclass(frozen=True)
class SeriesRequest:
    indicator: str
    countries: Iterable[str]
    start_year: int
    end_year: int


class IMFClient:
    def __init__(self, config: IMFConfig) -> None:
        self.config = config

    def fetch_indicator(self, indicator: str) -> Dict[str, Dict[int, float]]:
        """Fetch one indicator and return {country_code: {year: value}}."""
        url = f"{self.config.base_url}/{indicator}"
        response = requests.get(
            url,
            headers=self.config.headers(),
            timeout=self.config.timeout_seconds,
        )
        response.raise_for_status()
        payload = response.json()

        values = payload.get("values", {}).get(indicator)
        if not isinstance(values, dict):
            raise ValueError(f"Indicator '{indicator}' not found in IMF response.")

        cleaned: Dict[str, Dict[int, float]] = {}
        for country_code, year_map in values.items():
            if not isinstance(year_map, dict):
                continue
            parsed_years: Dict[int, float] = {}
            for year_str, value in year_map.items():
                try:
                    year = int(year_str)
                except (TypeError, ValueError):
                    continue
                try:
                    parsed_years[year] = float(value)
                except (TypeError, ValueError):
                    continue
            if parsed_years:
                cleaned[country_code] = parsed_years

        return cleaned

    def fetch_series(self, req: SeriesRequest) -> Dict[str, Dict[int, float]]:
        all_values = self.fetch_indicator(req.indicator)
        year_filtered: Dict[str, Dict[int, float]] = {}
        wanted = set(req.countries)
        for country_code, year_map in all_values.items():
            if country_code not in wanted:
                continue
            subset = {
                y: v
                for y, v in year_map.items()
                if req.start_year <= y <= req.end_year
            }
            year_filtered[country_code] = subset
        return year_filtered
