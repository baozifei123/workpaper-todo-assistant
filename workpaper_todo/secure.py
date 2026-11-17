# -*- coding: utf-8 -*-
"""Windows DPAPI helpers for optional encrypted backups.

The encrypted file can only be decrypted by the same Windows user on the same
computer (user-scope DPAPI).  Keep a normal backup as well.
"""
from __future__ import annotations

import ctypes
import ctypes.wintypes as wt
import os

crypt32 = ctypes.WinDLL("crypt32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

CRYPTPROTECT_UI_FORBIDDEN = 0x01
MAGIC = b"WPTENC1\n"


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wt.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_byte)),
    ]


crypt32.CryptProtectData.argtypes = [
    ctypes.POINTER(DATA_BLOB), wt.LPCWSTR, ctypes.POINTER(DATA_BLOB),
    ctypes.c_void_p, ctypes.c_void_p, wt.DWORD, ctypes.POINTER(DATA_BLOB),
]
crypt32.CryptProtectData.restype = wt.BOOL
crypt32.CryptUnprotectData.argtypes = [
    ctypes.POINTER(DATA_BLOB), ctypes.c_void_p, ctypes.POINTER(DATA_BLOB),
    ctypes.c_void_p, ctypes.c_void_p, wt.DWORD, ctypes.POINTER(DATA_BLOB),
]
crypt32.CryptUnprotectData.restype = wt.BOOL
kernel32.LocalFree.argtypes = [wt.HLOCAL]
kernel32.LocalFree.restype = wt.HLOCAL


def _make_blob(data: bytes):
    buffer = ctypes.create_string_buffer(data)
    blob = DATA_BLOB(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    return blob, buffer


def protect(data: bytes) -> bytes:
    in_blob, _buffer = _make_blob(data)
    out_blob = DATA_BLOB()
    if not crypt32.CryptProtectData(
            ctypes.byref(in_blob), "WorkpaperTodo backup", None,
            None, None, CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(out_blob)):
        raise OSError(ctypes.get_last_error(), "CryptProtectData failed")
    try:
        return ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(out_blob.pbData, wt.HLOCAL))


def unprotect(data: bytes) -> bytes:
    in_blob, _buffer = _make_blob(data)
    out_blob = DATA_BLOB()
    if not crypt32.CryptUnprotectData(
            ctypes.byref(in_blob), None, None, None, None,
            CRYPTPROTECT_UI_FORBIDDEN, ctypes.byref(out_blob)):
        raise OSError(ctypes.get_last_error(), "CryptUnprotectData failed")
    try:
        return ctypes.string_at(out_blob.pbData, out_blob.cbData)
    finally:
        kernel32.LocalFree(ctypes.cast(out_blob.pbData, wt.HLOCAL))


def encrypt_file(src_path: str, dest_path: str) -> None:
    with open(src_path, "rb") as f:
        protected = protect(f.read())
    with open(dest_path, "wb") as f:
        f.write(MAGIC)
        f.write(protected)


def decrypt_file(src_path: str, dest_path: str) -> None:
    with open(src_path, "rb") as f:
        data = f.read()
    if not data.startswith(MAGIC):
        raise ValueError("不是本工具生成的加密备份文件")
    plain = unprotect(data[len(MAGIC):])
    with open(dest_path, "wb") as f:
        f.write(plain)
