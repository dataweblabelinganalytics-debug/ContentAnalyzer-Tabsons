"""Run the idempotent genre/channel master migration."""

import json

from app import migrate_genre_hierarchy


if __name__ == "__main__":
    print(json.dumps(migrate_genre_hierarchy(force=True), indent=2))
