from __future__ import annotations

import json
import os
from abc import ABC, abstractmethod
from datetime import datetime
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

from telkom_batch_check import CrdbClient, LookupResult, parse_result, provider_is_telkom, query_number


class LookupProvider(ABC):
    name = "base"
    supports_captcha = False

    @abstractmethod
    def lookup(self, clean_number: str) -> LookupResult:
        raise NotImplementedError

    def new_captcha(self, clean_number: str) -> dict[str, str]:
        raise NotImplementedError("This lookup provider does not expose captcha challenges.")

    def submit_captcha(self, clean_number: str, code: str, pending: dict[str, str]) -> LookupResult:
        raise NotImplementedError("This lookup provider does not support captcha submission.")


class PublicCrdbProvider(LookupProvider):
    name = "public_crdb"
    supports_captcha = True

    def __init__(self, timeout: int = 45) -> None:
        self.client = CrdbClient(timeout=timeout)

    def lookup(self, clean_number: str) -> LookupResult:
        return query_number(self.client, clean_number)

    def new_captcha(self, clean_number: str) -> dict[str, str]:
        response = self.client.request("GET", "captcha/captcha-gen")
        return {
            "captchaEncrypt": str(response.get("stringData") or ""),
            "imageData": str(response.get("imageData") or ""),
        }

    def submit_captcha(self, clean_number: str, code: str, pending: dict[str, str]) -> LookupResult:
        payload = {
            "number": clean_number,
            "captcha": code.strip(),
            "captchaEncrypt": pending["captchaEncrypt"],
            "puid": self.client.puid,
        }
        response = self.client.request("POST", "publicInquiry/submitRequest", payload)
        return parse_result(clean_number, response)


class MockLookupProvider(LookupProvider):
    name = "mock"

    def lookup(self, clean_number: str) -> LookupResult:
        provider = "TELKOM" if clean_number.endswith(("0", "2", "4", "6", "8")) else "BACKSPACE"
        return LookupResult(
            clean_number,
            "Found",
            provider,
            "Yes" if provider_is_telkom(provider) else "No",
            f"Mock lookup result for development: serviced by {provider}",
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )


class LicensedHttpProvider(LookupProvider):
    name = "licensed_http"

    def __init__(self, endpoint: str, api_key: str = "", timeout: int = 30) -> None:
        self.endpoint = endpoint.rstrip("/")
        self.api_key = api_key
        self.timeout = timeout

    def lookup(self, clean_number: str) -> LookupResult:
        if not self.endpoint:
            return LookupResult(
                clean_number,
                "Lookup failed",
                "",
                "Unknown",
                "Licensed provider endpoint is not configured. Set LICENSED_LOOKUP_URL.",
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            )
        payload = json.dumps({"number": clean_number, "msisdn": clean_number}).encode("utf-8")
        headers = {"Content-Type": "application/json", "Accept": "application/json"}
        if self.api_key:
            headers["Authorization"] = f"Bearer {self.api_key}"
            headers["X-API-Key"] = self.api_key
        req = Request(self.endpoint, data=payload, method="POST", headers=headers)
        try:
            with urlopen(req, timeout=self.timeout) as response:
                data = json.loads(response.read().decode("utf-8", errors="replace"))
        except HTTPError as exc:
            body = exc.read().decode("utf-8", errors="replace")
            raw = f"Licensed provider HTTP {exc.code}: {exc.reason}. {body[:500]}"
            return self.failed(clean_number, raw)
        except URLError as exc:
            return self.failed(clean_number, f"Licensed provider network error: {exc.reason}")
        except TimeoutError:
            return self.failed(clean_number, "Licensed provider network timeout")
        except Exception as exc:
            return self.failed(clean_number, f"Licensed provider error: {type(exc).__name__}: {exc}")

        provider = str(
            data.get("current_provider")
            or data.get("provider")
            or data.get("operator")
            or data.get("network")
            or ""
        ).strip().upper()
        status = str(data.get("lookup_status") or data.get("status") or "").strip()
        raw = data.get("raw_result") or data.get("message") or json.dumps(data, ensure_ascii=False)
        if provider:
            return LookupResult(
                clean_number,
                "Found",
                provider,
                "Yes" if provider_is_telkom(provider) else "No",
                str(raw),
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            )
        return LookupResult(
            clean_number,
            status or "Needs review",
            "",
            "Unknown",
            str(raw),
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )

    @staticmethod
    def failed(clean_number: str, raw: str) -> LookupResult:
        return LookupResult(
            clean_number,
            "Lookup failed",
            "",
            "Unknown",
            raw,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        )


def build_lookup_provider() -> LookupProvider:
    provider = os.environ.get("LOOKUP_PROVIDER", "public_crdb").strip().lower()
    if provider == "mock":
        return MockLookupProvider()
    if provider in {"licensed", "licensed_http", "http"}:
        return LicensedHttpProvider(
            endpoint=os.environ.get("LICENSED_LOOKUP_URL", ""),
            api_key=os.environ.get("LICENSED_LOOKUP_API_KEY", ""),
            timeout=int(os.environ.get("LICENSED_LOOKUP_TIMEOUT", "30")),
        )
    return PublicCrdbProvider(timeout=int(os.environ.get("CRDB_TIMEOUT", "45")))
