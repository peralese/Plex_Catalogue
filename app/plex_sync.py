import os
import threading
from datetime import datetime

from dotenv import load_dotenv

load_dotenv()

BACKUP_TAGS = ["dvd", "blu-ray", "iso", "ripped"]

_status = {"running": False, "last_run": None, "error": None, "counts": {}}
_lock = threading.Lock()


def get_status():
    return dict(_status)


def run_sync():
    with _lock:
        if _status["running"]:
            return False
        _status["running"] = True
        _status["error"] = None
    threading.Thread(target=_do_sync, daemon=True).start()
    return True


def _get_label_tags(item):
    try:
        return [lab.tag.lower() for lab in item.labels]
    except Exception:
        return []


def _detect_backup(tags, file_path):
    br_aliases = {"blu-ray", "blue-ray", "bluray"}
    found = set()
    for t in BACKUP_TAGS:
        if t == "blu-ray":
            if any(alias in tags for alias in br_aliases):
                found.add("blu-ray")
        elif t in tags:
            found.add(t)
    fp = (file_path or "").lower()
    if ".iso" in fp:
        found.add("iso")
    if "dvd" in fp or ".vob" in fp:
        found.add("dvd")
    return found, bool(found)


def _do_sync():
    from plexapi.server import PlexServer
    from app.db import get_db

    try:
        baseurl = os.getenv("PLEX_BASEURL")
        token = os.getenv("PLEX_TOKEN")
        ignore = [
            lib.strip()
            for lib in os.getenv("IGNORE_LIBRARIES", "").split(",")
            if lib.strip()
        ]

        plex = PlexServer(baseurl, token)
        now = datetime.now().isoformat()
        counts = {"movies": 0, "episodes": 0}

        with get_db() as conn:
            conn.execute("DELETE FROM movies")
            conn.execute("DELETE FROM episodes")

            for section in plex.library.sections():
                if section.title in ignore:
                    continue

                if section.type == "movie":
                    for m in section.all():
                        tags = _get_label_tags(m)
                        try:
                            fp = m.media[0].parts[0].file
                        except Exception:
                            fp = ""
                        found, backed = _detect_backup(tags, fp)
                        conn.execute(
                            """INSERT INTO movies
                               (title, library, backed_up, backup_types, file_path, synced_at)
                               VALUES (?, ?, ?, ?, ?, ?)""",
                            (
                                m.title,
                                section.title,
                                1 if backed else 0,
                                ", ".join(sorted(found)).upper(),
                                fp,
                                now,
                            ),
                        )
                        counts["movies"] += 1

                elif section.type == "show":
                    for show in section.all():
                        show_labels = _get_label_tags(show)
                        for ep in show.episodes():
                            tags = _get_label_tags(ep) or show_labels
                            try:
                                fp = ep.media[0].parts[0].file
                            except Exception:
                                fp = ""
                            found, backed = _detect_backup(tags, fp)
                            conn.execute(
                                """INSERT INTO episodes
                                   (show_title, library, season, episode_num,
                                    episode_title, backed_up, backup_types, file_path, synced_at)
                                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                (
                                    show.title,
                                    section.title,
                                    ep.seasonNumber,
                                    ep.index,
                                    ep.title,
                                    1 if backed else 0,
                                    ", ".join(sorted(found)).upper(),
                                    fp,
                                    now,
                                ),
                            )
                            counts["episodes"] += 1

        _status.update({"running": False, "last_run": now, "error": None, "counts": counts})

    except Exception as exc:
        _status.update({"running": False, "error": str(exc)})
