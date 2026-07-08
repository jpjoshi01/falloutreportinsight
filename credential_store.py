from __future__ import annotations


SERVICE_NAME = "MigrationFalloutDashboard"
SUPPORTED_PROVIDERS = ("Kimi", "OpenAI", "Claude", "Custom")
PREFERENCE_ACCOUNT = "selected_provider"


class CredentialStoreError(RuntimeError):
    pass


def _account_name(provider: str) -> str:
    if provider not in SUPPORTED_PROVIDERS:
        raise CredentialStoreError(f"Unsupported provider: {provider}")
    return f"{provider}:api_key"


def _keyring():
    try:
        import keyring
    except Exception as exc:
        raise CredentialStoreError("The keyring package is not available. Install requirements.txt and try again.") from exc
    return keyring


def save_api_key(provider: str, api_key: str) -> None:
    if not api_key:
        raise CredentialStoreError("Enter an API key before saving.")
    try:
        _keyring().set_password(SERVICE_NAME, _account_name(provider), api_key)
    except Exception as exc:
        raise CredentialStoreError("Could not save the API key to Windows Credential Manager.") from exc


def get_api_key(provider: str) -> str | None:
    try:
        return _keyring().get_password(SERVICE_NAME, _account_name(provider))
    except Exception as exc:
        raise CredentialStoreError("Could not read the API key from Windows Credential Manager.") from exc


def delete_api_key(provider: str) -> None:
    try:
        _keyring().delete_password(SERVICE_NAME, _account_name(provider))
    except Exception as exc:
        message = str(exc).lower()
        if "not found" in message or "not set" in message or "no such" in message:
            return
        raise CredentialStoreError("Could not delete the API key from Windows Credential Manager.") from exc


def has_saved_api_key(provider: str) -> bool:
    key = get_api_key(provider)
    return bool(key)


def save_provider_preference(provider: str) -> None:
    if provider not in SUPPORTED_PROVIDERS:
        raise CredentialStoreError(f"Unsupported provider: {provider}")
    try:
        _keyring().set_password(SERVICE_NAME, PREFERENCE_ACCOUNT, provider)
    except Exception as exc:
        raise CredentialStoreError("Could not save the selected AI provider preference.") from exc


def get_provider_preference() -> str | None:
    try:
        provider = _keyring().get_password(SERVICE_NAME, PREFERENCE_ACCOUNT)
    except Exception as exc:
        raise CredentialStoreError("Could not read the selected AI provider preference.") from exc
    return provider if provider in SUPPORTED_PROVIDERS else None


def mask_api_key(api_key: str | None) -> str:
    if not api_key:
        return "Not saved"
    if len(api_key) <= 8:
        return "*" * len(api_key)
    return f"{api_key[:4]}...{api_key[-4:]}"
