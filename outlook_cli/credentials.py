"""OS credential operations with an explicit, noninteractive macOS policy.

Security.framework calls below use the documented SecItem query interface. They
are isolated here rather than modifying keyring internals or process-wide UI
policy, which would be unsafe when several request threads are active.
"""
from __future__ import annotations

import ctypes
from ctypes.util import find_library
import sys

import keyring
import keyring.errors

from .exceptions import AuthRequiredError

_NOT_FOUND = -25300
_INTERACTION_REQUIRED = {-25308, -25293, -128}


def _native_call(operation: str, service: str, username: str, password: str | None = None):
    """Return (OSStatus, secret) using query-local authentication UI denial."""
    security = ctypes.CDLL(find_library("Security"))
    core = ctypes.CDLL(find_library("CoreFoundation"))
    ptr = ctypes.c_void_p
    cfindex = ctypes.c_long
    status_type = ctypes.c_int32
    core.CFStringCreateWithCString.argtypes = (ptr, ctypes.c_char_p, ctypes.c_uint32)
    core.CFStringCreateWithCString.restype = ptr
    core.CFDictionaryCreate.argtypes = (ptr, ctypes.POINTER(ptr), ctypes.POINTER(ptr), cfindex, ptr, ptr)
    core.CFDictionaryCreate.restype = ptr
    core.CFDataCreate.argtypes = (ptr, ptr, cfindex)
    core.CFDataCreate.restype = ptr
    core.CFDataGetBytePtr.argtypes = (ptr,)
    core.CFDataGetBytePtr.restype = ptr
    core.CFDataGetLength.argtypes = (ptr,)
    core.CFDataGetLength.restype = cfindex
    core.CFRelease.argtypes = (ptr,)
    core.CFRelease.restype = None
    for name, parameters in {
        "SecItemCopyMatching": (ptr, ctypes.POINTER(ptr)),
        "SecItemUpdate": (ptr, ptr),
        "SecItemAdd": (ptr, ctypes.POINTER(ptr)),
        "SecItemDelete": (ptr,),
    }.items():
        function = getattr(security, name)
        function.argtypes, function.restype = parameters, status_type

    owned = []
    def constant(name):
        return ptr.in_dll(security, name)
    def string(value):
        result = core.CFStringCreateWithCString(None, value.encode("utf-8"), 0x08000100)
        owned.append(result)
        return result
    def dictionary(values):
        result = core.CFDictionaryCreate(None,
            (ptr * len(values))(*(constant(key) for key in values)),
            (ptr * len(values))(*values.values()), len(values),
            ctypes.cast(core.kCFTypeDictionaryKeyCallBacks, ptr),
            ctypes.cast(core.kCFTypeDictionaryValueCallBacks, ptr))
        owned.append(result)
        return result
    try:
        query = {
            "kSecClass": constant("kSecClassGenericPassword"),
            "kSecAttrService": string(service),
            "kSecAttrAccount": string(username),
            # Although deprecated in favor of LAContext, this documented key
            # supports the deployed macOS SecItem API without a PyObjC dependency.
            "kSecUseAuthenticationUI": constant("kSecUseAuthenticationUIFail"),
        }
        if operation == "get":
            query.update(kSecMatchLimit=constant("kSecMatchLimitOne"),
                         kSecReturnData=ptr.in_dll(core, "kCFBooleanTrue"))
            result = ptr()
            status = security.SecItemCopyMatching(dictionary(query), ctypes.byref(result))
            if result.value:
                owned.append(result)
            secret = None
            if status == 0 and result.value:
                secret = ctypes.string_at(core.CFDataGetBytePtr(result), core.CFDataGetLength(result)).decode("utf-8")
            return status, secret
        if operation == "delete":
            return security.SecItemDelete(dictionary(query)), None
        if operation == "set":
            encoded = password.encode("utf-8")
            data = core.CFDataCreate(None, ctypes.c_char_p(encoded), len(encoded))
            owned.append(data)
            status = security.SecItemUpdate(dictionary(query), dictionary({"kSecValueData": data}))
            if status == _NOT_FOUND:
                query["kSecValueData"] = data
                status = security.SecItemAdd(dictionary(query), None)
            return status, None
        raise ValueError(f"Unknown credential operation: {operation}")
    finally:
        for value in reversed(owned):
            if value:
                core.CFRelease(value)


def _noninteractive(operation, service, username, password=None):
    if sys.platform != "darwin":
        raise AuthRequiredError("Noninteractive keychain access is unavailable for this platform. Provide the token through the environment or authenticate interactively.")
    try:
        status, secret = _native_call(operation, service, username, password)
    except (AttributeError, OSError, ValueError) as exc:
        raise AuthRequiredError("Noninteractive Keychain access is unavailable. Run the command interactively to authorize credential access.") from exc
    if status == 0:
        return secret
    if status == _NOT_FOUND:
        if operation == "delete":
            raise keyring.errors.PasswordDeleteError("Credential not found")
        if operation == "get":
            return None
        raise AuthRequiredError("Could not store credentials in Keychain without interaction.")
    if status in _INTERACTION_REQUIRED:
        raise AuthRequiredError("Keychain access requires user authorization and --no-input forbids a prompt. Run outlook login interactively using this Python installation, then retry.")
    raise AuthRequiredError(f"Noninteractive Keychain operation failed (OSStatus {status}). Run the command interactively to check credential access.")


def get_password(service: str, username: str, *, allow_interactive: bool = True) -> str | None:
    if allow_interactive:
        return keyring.get_password(service, username)
    return _noninteractive("get", service, username)


def set_password(service: str, username: str, password: str, *, allow_interactive: bool = True) -> None:
    if allow_interactive:
        keyring.set_password(service, username, password)
    else:
        _noninteractive("set", service, username, password)


def delete_password(service: str, username: str, *, allow_interactive: bool = True) -> None:
    if allow_interactive:
        keyring.delete_password(service, username)
    else:
        _noninteractive("delete", service, username)
