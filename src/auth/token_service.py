"""
Microsoft token service - fetches tokens from Supabase oauth_connections.

Uses the ACTUAL schema with provider_metadata JSONB column.
"""

import asyncio
import logging
from datetime import datetime, timedelta, timezone
from typing import Optional, Dict
from dataclasses import dataclass

import asyncpg
import httpx
from cryptography.fernet import Fernet

logger = logging.getLogger(__name__)

PROVIDER = "microsoft"  # Distinguishes from 'procore'


@dataclass
class MicrosoftToken:
    """A valid Microsoft access token ready to use."""
    access_token: str
    user_id: str
    expires_at: Optional[datetime] = None


class TokenServiceError(Exception):
    """Token service error with HTTP status code."""
    def __init__(self, message: str, status_code: int = 500):
        self.message = message
        self.status_code = status_code
        super().__init__(message)


class TokenService:
    """Fetches Microsoft tokens from Supabase oauth_connections table."""

    def __init__(
        self,
        database_url: str,
        encryption_key: str,
        client_id: str,
        client_secret: str,
        tenant_id: str = "common"
    ):
        self.database_url = database_url
        self.cipher = Fernet(encryption_key.encode())
        self.client_id = client_id
        self.client_secret = client_secret
        self.tenant_id = tenant_id
        self._pool: Optional[asyncpg.Pool] = None
        self._locks: Dict[str, asyncio.Lock] = {}

    async def _get_pool(self) -> asyncpg.Pool:
        if self._pool is None:
            self._pool = await asyncpg.create_pool(
                self.database_url,
                min_size=1,
                max_size=5,
                command_timeout=10,
                statement_cache_size=0  # Required for Supabase PgBouncer
            )
        return self._pool

    async def close(self):
        if self._pool:
            await self._pool.close()
            self._pool = None

    async def get_token(self, user_id: str) -> MicrosoftToken:
        """Get valid Microsoft token for user, refreshing if needed."""

        if user_id not in self._locks:
            self._locks[user_id] = asyncio.Lock()

        async with self._locks[user_id]:
            pool = await self._get_pool()

            # Query using actual schema with provider_metadata JSONB
            row = await pool.fetchrow("""
                SELECT id, user_id, access_token, refresh_token, expires_at,
                       provider_metadata
                FROM oauth_connections
                WHERE user_id = $1
                  AND provider = $2
                  AND is_active = TRUE
                ORDER BY created_at DESC
                LIMIT 1
            """, user_id, PROVIDER)

            if not row:
                raise TokenServiceError(
                    "No Microsoft connection found. Please connect your Outlook account.",
                    status_code=404
                )

            # Decrypt tokens (stored as TEXT, Fernet returns bytes)
            try:
                access_token = self.cipher.decrypt(row["access_token"].encode()).decode()
                refresh_token = self.cipher.decrypt(row["refresh_token"].encode()).decode()
            except Exception as e:
                logger.error(f"Token decryption failed for user {user_id}: {e}")
                raise TokenServiceError("Token decryption failed. Please reconnect.", status_code=500)

            expires_at = row["expires_at"]
            connection_id = row["id"]

            # Check if refresh needed (expires in < 5 minutes)
            now = datetime.now(timezone.utc)
            if expires_at.tzinfo is None:
                expires_at = expires_at.replace(tzinfo=timezone.utc)

            if (expires_at - now).total_seconds() < 300:
                logger.info(f"Refreshing token for user {user_id}")
                access_token = await self._refresh_token(refresh_token, connection_id, pool)

            # Update last_used_at
            await pool.execute(
                "UPDATE oauth_connections SET last_used_at = NOW() WHERE id = $1",
                connection_id
            )

            return MicrosoftToken(access_token=access_token, user_id=user_id, expires_at=expires_at)

    async def _refresh_token(self, refresh_token: str, connection_id: str, pool: asyncpg.Pool) -> str:
        """Refresh expired Microsoft token."""

        token_url = f"https://login.microsoftonline.com/{self.tenant_id}/oauth2/v2.0/token"

        async with httpx.AsyncClient() as client:
            response = await client.post(token_url, data={
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": "offline_access User.Read Mail.Read"
            })

            if response.status_code != 200:
                logger.error(f"Token refresh failed: {response.status_code}")
                raise TokenServiceError("Token refresh failed. Please reconnect.", status_code=401)

            data = response.json()
            new_access = data["access_token"]
            new_refresh = data.get("refresh_token", refresh_token)
            expires_in = data.get("expires_in", 3600)  # Microsoft: 1 hour

            # Encrypt and store (as TEXT)
            encrypted_access = self.cipher.encrypt(new_access.encode()).decode()
            encrypted_refresh = self.cipher.encrypt(new_refresh.encode()).decode()
            new_expires_at = datetime.now(timezone.utc) + timedelta(seconds=expires_in)

            await pool.execute("""
                UPDATE oauth_connections
                SET access_token = $1, refresh_token = $2, expires_at = $3, updated_at = NOW()
                WHERE id = $4
            """, encrypted_access, encrypted_refresh, new_expires_at, connection_id)

            return new_access
