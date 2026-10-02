"""Supabase Storage helper for product images.

Env vars (Koyeb / local):
  SUPABASE_URL              e.g. https://xxxx.supabase.co
  SUPABASE_SERVICE_ROLE_KEY service_role key (upload + overwrite)
  SUPABASE_BUCKET           default: order-images

Bucket should be Public (or use signed URLs later).
"""
from __future__ import annotations

import os
import re
import uuid
from typing import Optional, Tuple
from urllib import error, request


def storage_configured() -> bool:
    url = (os.environ.get("SUPABASE_URL") or "").strip().rstrip("/")
    key = (os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or os.environ.get("SUPABASE_KEY") or "").strip()
    return bool(url and key)


def _bucket() -> str:
    return (os.environ.get("SUPABASE_BUCKET") or "order-images").strip() or "order-images"


def _base() -> Tuple[str, str]:
    url = (os.environ.get("SUPABASE_URL") or "").strip().rstrip("/")
    key = (os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or os.environ.get("SUPABASE_KEY") or "").strip()
    return url, key


def _ext_from_mime(mime: str, fallback_name: str = "") -> str:
    m = (mime or "").lower()
    if "png" in m:
        return ".png"
    if "gif" in m:
        return ".gif"
    if "webp" in m:
        return ".webp"
    if "bmp" in m:
        return ".bmp"
    if fallback_name:
        low = fallback_name.lower()
        for e in (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp"):
            if low.endswith(e):
                return ".jpg" if e == ".jpeg" else e
    return ".jpg"


def upload_image(
    data: bytes,
    mime: str = "image/jpeg",
    filename_hint: str = "",
    folder: str = "products",
) -> Optional[str]:
    """
    Upload image bytes to Supabase Storage.
    Returns public URL on success, None on failure / not configured.
    """
    if not data:
        return None
    if not storage_configured():
        return None

    base, key = _base()
    bucket = _bucket()
    ext = _ext_from_mime(mime, filename_hint)
    # safe path
    safe_folder = re.sub(r"[^a-zA-Z0-9_/\-]", "", folder).strip("/") or "products"
    object_path = f"{safe_folder}/{uuid.uuid4().hex}{ext}"

    upload_url = f"{base}/storage/v1/object/{bucket}/{object_path}"
    headers = {
        "Authorization": f"Bearer {key}",
        "apikey": key,
        "Content-Type": mime or "application/octet-stream",
        "x-upsert": "true",
    }

    try:
        req = request.Request(upload_url, data=data, headers=headers, method="POST")
        with request.urlopen(req, timeout=60) as resp:
            _ = resp.read()
        public_url = f"{base}/storage/v1/object/public/{bucket}/{object_path}"
        print(f"[storage] uploaded {len(data)} bytes -> {public_url}")
        return public_url
    except error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")[:300]
        print(f"[storage] upload HTTP {e.code}: {body}")
        return None
    except Exception as e:
        print(f"[storage] upload failed: {e}")
        return None


def image_ref_from_bytes(
    data: bytes,
    mime: str = "image/jpeg",
    filename_hint: str = "",
) -> str:
    """
    Prefer Supabase Storage URL; fallback to data-URI base64.
    Always returns a string usable in <img src="...">.
    """
    import base64

    url = upload_image(data, mime=mime, filename_hint=filename_hint)
    if url:
        return url
    b64 = base64.b64encode(data).decode("ascii")
    return f"data:{mime or 'image/jpeg'};base64,{b64}"
