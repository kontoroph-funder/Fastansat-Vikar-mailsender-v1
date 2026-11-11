"""Current-user Windows DPAPI storage; credentials have no plaintext fallback."""
from __future__ import annotations

import ctypes
from ctypes import wintypes
import json
import os
from pathlib import Path
import re
import tempfile
import threading


class VaultError(RuntimeError):
    pass


class _Blob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_ubyte))]


def _crypt(data: bytes, decrypt: bool = False) -> bytes:
    if os.name != "nt":
        raise VaultError("Credential-lageret kræver Windows DPAPI; klartekst er ikke tilladt.")
    crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    buffer = ctypes.create_string_buffer(data)
    source = _Blob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_ubyte)))
    destination = _Blob()
    flags = 0x01  # CRYPTPROTECT_UI_FORBIDDEN: never show a dialog in the service.
    if decrypt:
        fn = crypt32.CryptUnprotectData
        fn.argtypes = [ctypes.POINTER(_Blob), ctypes.c_void_p, ctypes.c_void_p,
                       ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_Blob)]
        fn.restype = wintypes.BOOL
        ok = fn(ctypes.byref(source), None, None, None, None, flags, ctypes.byref(destination))
    else:
        fn = crypt32.CryptProtectData
        fn.argtypes = [ctypes.POINTER(_Blob), wintypes.LPCWSTR, ctypes.c_void_p,
                       ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_Blob)]
        fn.restype = wintypes.BOOL
        ok = fn(ctypes.byref(source), "LønAutomatisering", None, None, None,
                flags, ctypes.byref(destination))  # Current Windows user only.
    if not ok:
        raise VaultError("Windows kunne ikke åbne eller beskytte credentials på denne computer.")
    try:
        return ctypes.string_at(destination.pbData, destination.cbData)
    finally:
        kernel32.LocalFree(destination.pbData)


class SecretStore:
    def __init__(self, root: Path):
        self.root = Path(root).resolve()
        self._lock = threading.RLock()

    def _path(self, name: str) -> Path:
        if not isinstance(name, str) or not re.fullmatch(r"[a-z][a-z0-9_]{0,63}", name):
            raise VaultError("Ugyldigt credential-navn.")
        path = self.root / (name + ".dpapi")
        if path.is_symlink():
            raise VaultError("Credential-filen må ikke være et symbolsk link.")
        return path

    def get(self, name: str) -> dict | None:
        with self._lock:
            path = self._path(name)
            try:
                encrypted = path.read_bytes()
            except FileNotFoundError:
                return None
            if len(encrypted) > 10 * 1024 * 1024:
                raise VaultError("Credential-filen er for stor.")
            try:
                value = json.loads(_crypt(encrypted, decrypt=True))
            except (ValueError, UnicodeError) as exc:
                raise VaultError("Credential-lageret er beskadiget.") from exc
            if not isinstance(value, dict):
                raise VaultError("Credential-lageret er beskadiget.")
            return value

    def set(self, name: str, value: dict) -> None:
        if not isinstance(value, dict):
            raise VaultError("Credentials skal være et objekt.")
        with self._lock:
            path = self._path(name)
            encrypted = _crypt(json.dumps(value, ensure_ascii=False).encode("utf-8"))
            self.root.mkdir(parents=True, exist_ok=True)
            fd, temporary = tempfile.mkstemp(prefix=".credential-", suffix=".tmp", dir=self.root)
            try:
                with os.fdopen(fd, "wb") as stream:
                    stream.write(encrypted)
                    stream.flush()
                    os.fsync(stream.fileno())
                os.replace(temporary, path)
            finally:
                if os.path.exists(temporary):
                    os.unlink(temporary)

    def delete(self, name: str) -> None:
        with self._lock:
            self._path(name).unlink(missing_ok=True)
