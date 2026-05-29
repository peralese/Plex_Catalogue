import sqlite3
import os
from contextlib import contextmanager

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "catalogue.db")


@contextmanager
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    with get_db() as conn:
        conn.executescript("""
            CREATE TABLE IF NOT EXISTS movies (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                title        TEXT NOT NULL,
                library      TEXT NOT NULL,
                backed_up    INTEGER DEFAULT 0,
                backup_types TEXT DEFAULT '',
                file_path    TEXT DEFAULT '',
                synced_at    TEXT
            );

            CREATE TABLE IF NOT EXISTS episodes (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                show_title    TEXT NOT NULL,
                library       TEXT NOT NULL,
                season        INTEGER,
                episode_num   INTEGER,
                episode_title TEXT,
                backed_up     INTEGER DEFAULT 0,
                backup_types  TEXT DEFAULT '',
                file_path     TEXT DEFAULT '',
                synced_at     TEXT
            );

            CREATE TABLE IF NOT EXISTS wishlist (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                title      TEXT NOT NULL,
                type       TEXT DEFAULT '',
                release    TEXT DEFAULT '',
                notes      TEXT DEFAULT '',
                created_at TEXT DEFAULT (datetime('now'))
            );
        """)


# ── Movies ────────────────────────────────────────────────────────────────────

def get_movies(library="", backup="", search=""):
    query = "SELECT * FROM movies WHERE 1=1"
    params = []
    if library:
        query += " AND library = ?"
        params.append(library)
    if backup == "yes":
        query += " AND backed_up = 1"
    elif backup == "no":
        query += " AND backed_up = 0"
    if search:
        query += " AND title LIKE ?"
        params.append(f"%{search}%")
    query += " ORDER BY title"
    with get_db() as conn:
        return [dict(r) for r in conn.execute(query, params).fetchall()]


def get_movie_libraries():
    with get_db() as conn:
        return [r[0] for r in conn.execute(
            "SELECT DISTINCT library FROM movies ORDER BY library"
        ).fetchall()]


# ── TV ────────────────────────────────────────────────────────────────────────

def get_tv_shows():
    with get_db() as conn:
        return [dict(r) for r in conn.execute("""
            SELECT show_title, library,
                   COUNT(*)         AS total_eps,
                   SUM(backed_up)   AS backed_eps,
                   ROUND(SUM(backed_up) * 100.0 / COUNT(*), 1) AS pct
            FROM episodes
            GROUP BY show_title, library
            ORDER BY show_title
        """).fetchall()]


def get_show_episodes(show_title):
    with get_db() as conn:
        return [dict(r) for r in conn.execute("""
            SELECT season, episode_num, episode_title,
                   backed_up, backup_types, file_path
            FROM episodes
            WHERE show_title = ?
            ORDER BY season, episode_num
        """, (show_title,)).fetchall()]


# ── Dashboard stats ───────────────────────────────────────────────────────────

def get_dashboard_stats():
    with get_db() as conn:
        total_movies  = conn.execute("SELECT COUNT(*) FROM movies").fetchone()[0]
        backed_movies = conn.execute("SELECT COUNT(*) FROM movies WHERE backed_up = 1").fetchone()[0]
        total_eps     = conn.execute("SELECT COUNT(*) FROM episodes").fetchone()[0]
        backed_eps    = conn.execute("SELECT COUNT(*) FROM episodes WHERE backed_up = 1").fetchone()[0]
        total_wishlist = conn.execute("SELECT COUNT(*) FROM wishlist").fetchone()[0]

        libs = [dict(r) for r in conn.execute("""
            SELECT library,
                   COUNT(*)       AS total,
                   SUM(backed_up) AS backed
            FROM movies
            GROUP BY library
            ORDER BY library
        """).fetchall()]

    return {
        "total_movies":   total_movies,
        "backed_movies":  backed_movies,
        "total_eps":      total_eps,
        "backed_eps":     backed_eps,
        "total_wishlist": total_wishlist,
        "libs":           libs,
    }


# ── Wishlist ──────────────────────────────────────────────────────────────────

def get_wishlist():
    with get_db() as conn:
        return [dict(r) for r in conn.execute(
            "SELECT * FROM wishlist ORDER BY created_at DESC"
        ).fetchall()]


def add_wishlist_item(title, type_, release, notes):
    with get_db() as conn:
        conn.execute(
            "INSERT INTO wishlist (title, type, release, notes) VALUES (?, ?, ?, ?)",
            (title, type_, release, notes),
        )


def update_wishlist_item(item_id, title, type_, release, notes):
    with get_db() as conn:
        conn.execute(
            "UPDATE wishlist SET title=?, type=?, release=?, notes=? WHERE id=?",
            (title, type_, release, notes, item_id),
        )


def delete_wishlist_item(item_id):
    with get_db() as conn:
        conn.execute("DELETE FROM wishlist WHERE id=?", (item_id,))
