from __future__ import annotations

import asyncio
import base64
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import MagicMock

from starlette.requests import Request
from starlette.responses import JSONResponse, Response

from horoshop_sets import Credentials, load_settings
import sets_server


class AuthAndConfigTests(unittest.TestCase):
    def test_load_settings_reads_horoshop_and_auth_credentials(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            config_path = Path(temp_dir) / "config.json"
            config_path.write_text(
                json.dumps(
                    {
                        "server": {"host": "0.0.0.0", "port": 8093},
                        "horoshop": {
                            "domain": "https://attoys.com.ua",
                            "login": "z.halatsyn@atuatoys.kiev.ua",
                            "password": "Bublik4466Mv99oqrp0p",
                        },
                        "auth": {
                            "enabled": True,
                            "login": "admin",
                            "password": "Bublik44",
                        },
                        "admin_password": "Bublik44",
                    }
                ),
                encoding="utf-8",
            )
            settings = load_settings(config_path)
            self.assertEqual(settings.domain, "https://attoys.com.ua")
            self.assertEqual(settings.login, "z.halatsyn@atuatoys.kiev.ua")
            self.assertEqual(settings.password, "Bublik4466Mv99oqrp0p")
            self.assertTrue(settings.auth_enabled)
            self.assertEqual(settings.auth_login, "admin")
            self.assertEqual(settings.auth_password, "Bublik44")

    def test_credentials_fall_back_to_settings(self) -> None:
        runtime_settings, _ = sets_server.get_runtime()
        # With empty login and password, credentials_from_form and credentials_from_json
        # must take runtime_settings.login and password from config.json
        form_creds = sets_server.credentials_from_form({})
        json_creds = sets_server.credentials_from_json({})
        self.assertEqual(form_creds.login, runtime_settings.login)
        self.assertEqual(form_creds.password, runtime_settings.password)
        self.assertEqual(json_creds.login, runtime_settings.login)
        self.assertEqual(json_creds.password, runtime_settings.password)

    def test_session_token_and_auth_verification(self) -> None:
        runtime_settings, _ = sets_server.get_runtime()
        token = sets_server.get_session_token(runtime_settings.auth_login, runtime_settings.auth_password)

        # 1. Unauthenticated request
        req_unauth = Request({"type": "http", "headers": [], "cookies": {}})
        self.assertFalse(sets_server.is_request_authenticated(req_unauth, runtime_settings))

        # 2. Authenticated via sets_session cookie
        req_cookie = Request({
            "type": "http",
            "headers": [(b"cookie", f"sets_session={token}".encode("ascii"))],
        })
        self.assertTrue(sets_server.is_request_authenticated(req_cookie, runtime_settings))

        # 3. Authenticated via admin_session cookie
        req_admin_cookie = Request({
            "type": "http",
            "headers": [(b"cookie", f"admin_session={token}".encode("ascii"))],
        })
        self.assertTrue(sets_server.is_request_authenticated(req_admin_cookie, runtime_settings))

        # 4. Authenticated via Basic Auth header
        basic_b64 = base64.b64encode(f"{runtime_settings.auth_login}:{runtime_settings.auth_password}".encode()).decode()
        req_basic = Request({
            "type": "http",
            "headers": [(b"authorization", f"Basic {basic_b64}".encode("ascii"))],
        })
        self.assertTrue(sets_server.is_request_authenticated(req_basic, runtime_settings))

        # 5. Invalid credentials fail
        bad_basic = base64.b64encode(b"admin:WrongPassword").decode()
        req_bad = Request({
            "type": "http",
            "headers": [(b"authorization", f"Basic {bad_basic}".encode("ascii"))],
        })
        self.assertFalse(sets_server.is_request_authenticated(req_bad, runtime_settings))

    def test_login_and_logout_endpoints(self) -> None:
        runtime_settings, _ = sets_server.get_runtime()

        # 1. Bad login
        scope_bad = {
            "type": "http",
            "headers": [(b"content-type", b"application/json")],
        }
        req_bad = Request(scope_bad)
        req_bad._json = {"login": "admin", "password": "WrongPassword"}
        async def fake_json_bad():
            return {"login": "admin", "password": "WrongPassword"}
        req_bad.json = fake_json_bad

        resp_bad = asyncio.run(sets_server.login_endpoint(req_bad))
        self.assertEqual(resp_bad.status_code, 401)

        # 2. Successful login
        scope_ok = {
            "type": "http",
            "headers": [(b"content-type", b"application/json")],
        }
        req_ok = Request(scope_ok)
        async def fake_json_ok():
            return {"login": runtime_settings.auth_login, "password": runtime_settings.auth_password}
        req_ok.json = fake_json_ok

        resp_ok = asyncio.run(sets_server.login_endpoint(req_ok))
        self.assertEqual(resp_ok.status_code, 200)
        # Verify Set-Cookie header contains session
        cookie_headers = [v.decode("latin1") if isinstance(v, bytes) else v for k, v in resp_ok.raw_headers if k.lower() == b"set-cookie"]
        self.assertTrue(any("sets_session=" in h for h in cookie_headers))

        # 3. Logout endpoint
        resp_logout = sets_server.logout_endpoint()
        self.assertEqual(resp_logout.status_code, 200)
        logout_cookies = [v.decode("latin1") if isinstance(v, bytes) else v for k, v in resp_logout.raw_headers if k.lower() == b"set-cookie"]
        self.assertTrue(any('sets_session=""' in h or "max-age=0" in h.lower() for h in logout_cookies))

    def test_index_redirects_when_unauthenticated(self) -> None:
        req_unauth = Request({"type": "http", "headers": [], "path": "/"})
        resp = sets_server.index(req_unauth)
        self.assertEqual(resp.status_code, 303)
        self.assertEqual(resp.headers["location"], "/login")


if __name__ == "__main__":
    unittest.main()
