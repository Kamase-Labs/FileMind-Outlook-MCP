"""
Configuration for Outlook MCP Server.
Mirrors Procore MCP config.py exactly.
"""

from typing import Literal, Optional
from pydantic import Field, field_validator
from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    """Application settings loaded from environment variables."""

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
        case_sensitive=False,
        extra="ignore"
    )

    # Server
    server_host: str = Field(default="0.0.0.0")
    server_port: int = Field(default=8002, ge=1, le=65535)

    # Supabase
    supabase_url: str = Field(..., description="PostgreSQL connection URL")
    encryption_key: str = Field(..., description="Fernet key for token decryption")

    # Microsoft OAuth
    microsoft_client_id: str = Field(...)
    microsoft_client_secret: str = Field(...)
    microsoft_tenant_id: str = Field(default="common")

    # Auth Mode
    trust_x_user_id: bool = Field(default=True)

    # Logging
    log_level: Literal["DEBUG", "INFO", "WARNING", "ERROR"] = Field(default="INFO")

    # Email field selections (matches Node.js config.js exactly)
    email_list_fields: str = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,bodyPreview,hasAttachments,importance,isRead"
    email_detail_fields: str = "id,subject,from,toRecipients,ccRecipients,bccRecipients,receivedDateTime,bodyPreview,body,hasAttachments,importance,isRead"

    @property
    def microsoft_token_url(self) -> str:
        return f"https://login.microsoftonline.com/{self.microsoft_tenant_id}/oauth2/v2.0/token"

    @field_validator("encryption_key")
    @classmethod
    def validate_encryption_key(cls, v: str) -> str:
        if len(v) < 32:
            raise ValueError("Encryption key must be valid Fernet key")
        return v


# Singleton
_settings: Settings | None = None


def get_settings() -> Settings:
    global _settings
    if _settings is None:
        _settings = Settings()
    return _settings
