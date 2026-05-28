from __future__ import annotations

import base64
import json
import os
import secrets
import sys
from pathlib import Path
from typing import Any
from urllib.error import HTTPError
from urllib.parse import urlencode
from urllib.request import Request, urlopen


ROOT = Path(__file__).resolve().parent
FILES_TO_UPLOAD = [
    "render_app.py",
    "telkom_batch_check.py",
    "requirements.txt",
    "render.yaml",
    "RENDER_DEPLOY.md",
]


def load_env(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            values[key.strip()] = value.strip().strip('"').strip("'")
    for key, value in os.environ.items():
        values.setdefault(key, value)
    return values


def fail(message: str) -> None:
    print(f"ERROR: {message}")
    sys.exit(1)


def request_json(method: str, url: str, headers: dict[str, str], body: Any | None = None) -> Any:
    data = None
    request_headers = dict(headers)
    if body is not None:
        data = json.dumps(body).encode("utf-8")
        request_headers["Content-Type"] = "application/json"
    req = Request(url, data=data, method=method, headers=request_headers)
    try:
        with urlopen(req, timeout=90) as response:
            raw = response.read().decode("utf-8", errors="replace")
            return json.loads(raw) if raw else {}
    except HTTPError as exc:
        raw = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"{method} {url} failed with HTTP {exc.code}: {raw[:2000]}") from exc


def github_headers(token: str) -> dict[str, str]:
    return {
        "Accept": "application/vnd.github+json",
        "Authorization": f"Bearer {token}",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "telkom-render-deployer",
    }


def render_headers(token: str) -> dict[str, str]:
    return {
        "Accept": "application/json",
        "Authorization": f"Bearer {token}",
        "User-Agent": "telkom-render-deployer",
    }


def get_or_create_repo(env: dict[str, str]) -> dict[str, Any]:
    token = env["GITHUB_TOKEN"]
    repo_name = env.get("GITHUB_REPO", "telkom-manual-checker")
    private = env.get("GITHUB_PRIVATE", "true").lower() == "true"
    headers = github_headers(token)
    repo_full_name = env.get("GITHUB_REPO_FULL_NAME", "").strip()
    if repo_full_name:
        repo = request_json("GET", f"https://api.github.com/repos/{repo_full_name}", headers)
        print(f"Using existing GitHub repo: {repo['html_url']}")
        return repo

    user = request_json("GET", "https://api.github.com/user", headers)
    owner = user["login"]
    repo_url = f"https://api.github.com/repos/{owner}/{repo_name}"
    try:
        repo = request_json("GET", repo_url, headers)
        print(f"Using existing GitHub repo: {repo['html_url']}")
        return repo
    except RuntimeError as exc:
        if "HTTP 404" not in str(exc):
            raise
    repo = request_json(
        "POST",
        "https://api.github.com/user/repos",
        headers,
        {
            "name": repo_name,
            "private": private,
            "auto_init": True,
            "description": "Manual captcha-assisted Telkom CRDB checking app.",
        },
    )
    print(f"Created GitHub repo: {repo['html_url']}")
    return repo


def put_file(repo: dict[str, Any], token: str, relative_path: str) -> None:
    owner = repo["owner"]["login"]
    repo_name = repo["name"]
    path = ROOT / relative_path
    content = base64.b64encode(path.read_bytes()).decode("ascii")
    headers = github_headers(token)
    encoded_path = relative_path.replace("\\", "/")
    url = f"https://api.github.com/repos/{owner}/{repo_name}/contents/{encoded_path}"
    sha = None
    try:
        existing = request_json("GET", url, headers)
        sha = existing.get("sha")
    except RuntimeError as exc:
        if "HTTP 404" not in str(exc):
            raise
    body = {
        "message": f"Deploy Telkom checker: {relative_path}",
        "content": content,
        "branch": repo.get("default_branch", "main"),
    }
    if sha:
        body["sha"] = sha
    request_json("PUT", url, headers, body)
    print(f"Uploaded {relative_path}")


def upload_repo_files(repo: dict[str, Any], env: dict[str, str]) -> None:
    for relative_path in FILES_TO_UPLOAD:
        put_file(repo, env["GITHUB_TOKEN"], relative_path)


def get_owner_id(env: dict[str, str]) -> str:
    if env.get("RENDER_OWNER_ID"):
        return env["RENDER_OWNER_ID"]
    owners = request_json(
        "GET",
        "https://api.render.com/v1/owners?limit=100",
        render_headers(env["RENDER_API_KEY"]),
    )
    if len(owners) == 1:
        owner_id = owners[0]["owner"]["id"] if "owner" in owners[0] else owners[0]["id"]
        print(f"Using Render owner/workspace: {owner_id}")
        return owner_id
    print("Render workspaces available:")
    for item in owners:
        owner = item.get("owner", item)
        print(f"- {owner.get('id')}  {owner.get('name') or owner.get('email')}")
    fail("Set RENDER_OWNER_ID in .env to the workspace ID you want to use.")


def existing_render_service(env: dict[str, str], owner_id: str) -> dict[str, Any] | None:
    query = urlencode({"name": env["RENDER_SERVICE_NAME"], "ownerId": owner_id, "limit": 20})
    services = request_json(
        "GET",
        f"https://api.render.com/v1/services?{query}",
        render_headers(env["RENDER_API_KEY"]),
    )
    for item in services:
        service = item.get("service", item)
        if service.get("name") == env["RENDER_SERVICE_NAME"]:
            return service
    return None


def create_render_service(repo: dict[str, Any], env: dict[str, str]) -> dict[str, Any]:
    owner_id = get_owner_id(env)
    existing = existing_render_service(env, owner_id)
    if existing:
        print(f"Render service already exists: {existing.get('dashboardUrl')}")
        return existing

    app_password = env.get("APP_PASSWORD") or secrets.token_urlsafe(18)
    env["APP_PASSWORD"] = app_password
    payload = {
        "type": "web_service",
        "name": env["RENDER_SERVICE_NAME"],
        "ownerId": owner_id,
        "repo": repo["html_url"],
        "branch": repo.get("default_branch", "main"),
        "autoDeploy": "yes",
        "envVars": [
            {"key": "DATA_DIR", "value": "/tmp/telkom-render-checker"},
            {"key": "SECRET_KEY", "generateValue": True},
            {"key": "APP_PASSWORD", "value": app_password},
        ],
        "serviceDetails": {
            "runtime": "python",
            "plan": env.get("RENDER_PLAN", "free"),
            "region": env.get("RENDER_REGION", "oregon"),
            "healthCheckPath": "/health",
            "envSpecificDetails": {
                "buildCommand": "pip install -r requirements.txt",
                "startCommand": "gunicorn render_app:app",
            },
        },
    }
    created = request_json("POST", "https://api.render.com/v1/services", render_headers(env["RENDER_API_KEY"]), payload)
    service = created.get("service", created)
    print(f"Created Render service: {service.get('dashboardUrl')}")
    print(f"App password: {app_password}")
    return service


def main() -> None:
    env = load_env(ROOT / ".env")
    missing = [key for key in ["GITHUB_TOKEN", "RENDER_API_KEY"] if not env.get(key)]
    if missing:
        fail(f"Missing required values in .env: {', '.join(missing)}")
    env.setdefault("GITHUB_REPO", "telkom-manual-checker")
    env.setdefault("GITHUB_PRIVATE", "true")
    env.setdefault("RENDER_SERVICE_NAME", "telkom-manual-checker")
    env.setdefault("RENDER_PLAN", "free")
    env.setdefault("RENDER_REGION", "oregon")

    for relative_path in FILES_TO_UPLOAD:
        if not (ROOT / relative_path).exists():
            fail(f"Required file missing: {relative_path}")

    repo = get_or_create_repo(env)
    upload_repo_files(repo, env)
    service = create_render_service(repo, env)

    service_url = f"https://{service.get('slug', env['RENDER_SERVICE_NAME'])}.onrender.com/"
    print("")
    print("Deployment has been created or reused.")
    print(f"GitHub repo: {repo['html_url']}")
    print(f"Render dashboard: {service.get('dashboardUrl')}")
    print(f"App URL: {service_url}?key={env.get('APP_PASSWORD', '<existing-password>')}")
    print("")
    print("Render will still need a few minutes to build and start the app.")


if __name__ == "__main__":
    main()
