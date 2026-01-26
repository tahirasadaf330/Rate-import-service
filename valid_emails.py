from typing import List

VERIFIED_SENDERS: List[str] = []

def refresh_verified_senders(active_only: bool = True) -> List[str]:
    """
    Fetch emails from the database and replace the contents of VERIFIED_SENDERS
    *in-place* so existing imports keep seeing the updated values.

    Returns the refreshed VERIFIED_SENDERS.
    """
    try:
        # Local import to avoid circulars and allow this module to be imported
        # in environments where DB isn't available (e.g., tooling).
        # Verified senders are vendor contacts (not authorized_senders).
        from database import fetch_vendor_contact_emails
        emails = fetch_vendor_contact_emails(active_only=active_only)

        # mutate in place; do NOT rebind the name
        VERIFIED_SENDERS.clear()
        VERIFIED_SENDERS.extend(emails)
        return VERIFIED_SENDERS
    except Exception as e:
        # Don’t crash your app if DB is temporarily unavailable
        print(f"(warn) refresh_verified_senders: DB fetch failed: {e}")
        return VERIFIED_SENDERS  # whatever it had before (possibly empty)


def get_verified_senders(active_only: bool = True) -> List[str]:
    """
    Convenience accessor. If the list is empty, refresh it once.
    You can call this anywhere you currently use VERIFIED_SENDERS.
    """
    if not VERIFIED_SENDERS:
        refresh_verified_senders(active_only=active_only)
    return VERIFIED_SENDERS


