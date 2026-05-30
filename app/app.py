from flask import Flask, render_template, request, jsonify
from dotenv import load_dotenv

from app.db import (
    init_db,
    get_dashboard_stats,
    get_movies,
    get_movie_libraries,
    get_tv_shows,
    get_show_episodes,
    get_wishlist,
    add_wishlist_item,
    update_wishlist_item,
    delete_wishlist_item,
)
from app.plex_sync import run_sync, get_status
from modules.google_sync import sync_wishlist_to_gsheet

load_dotenv()

app = Flask(__name__)
init_db()


# ── Pages ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    stats = get_dashboard_stats()
    sync = get_status()
    return render_template("index.html", stats=stats, sync=sync)


@app.route("/movies")
def movies():
    library = request.args.get("library", "")
    backup  = request.args.get("backup", "")
    search  = request.args.get("search", "")
    rows      = get_movies(library=library, backup=backup, search=search)
    libraries = get_movie_libraries()
    return render_template(
        "movies.html",
        movies=rows,
        libraries=libraries,
        selected_library=library,
        selected_backup=backup,
        search=search,
    )


@app.route("/tv")
def tv():
    shows = get_tv_shows()
    return render_template("tv.html", shows=shows)


@app.route("/tv/<path:show_title>")
def tv_show(show_title):
    episodes = get_show_episodes(show_title)
    return render_template("tv_show.html", show_title=show_title, episodes=episodes)


@app.route("/wishlist")
def wishlist():
    libraries = get_movie_libraries() + ["TV Shows"]
    return render_template("wishlist.html", libraries=libraries)


# ── Wishlist API ──────────────────────────────────────────────────────────────

@app.route("/api/wishlist", methods=["GET"])
def api_get_wishlist():
    return jsonify(get_wishlist())


@app.route("/api/wishlist", methods=["POST"])
def api_add_wishlist():
    d = request.get_json()
    add_wishlist_item(
        d.get("title", ""),
        d.get("type", ""),
        d.get("release", ""),
        d.get("notes", ""),
    )
    return jsonify({"ok": True}), 201


@app.route("/api/wishlist/<int:item_id>", methods=["PUT"])
def api_update_wishlist(item_id):
    d = request.get_json()
    update_wishlist_item(
        item_id,
        d.get("title", ""),
        d.get("type", ""),
        d.get("release", ""),
        d.get("notes", ""),
    )
    return jsonify({"ok": True})


@app.route("/api/wishlist/<int:item_id>", methods=["DELETE"])
def api_delete_wishlist(item_id):
    delete_wishlist_item(item_id)
    return jsonify({"ok": True})


@app.route("/api/wishlist/sync", methods=["POST"])
def api_wishlist_sync():
    try:
        items = get_wishlist()
        count = sync_wishlist_to_gsheet(items)
        return jsonify({"ok": True, "synced": count})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)}), 500


# ── Sync API ──────────────────────────────────────────────────────────────────

@app.route("/api/sync", methods=["POST"])
def api_sync():
    started = run_sync()
    return jsonify({"started": started, "status": get_status()})


@app.route("/api/sync/status")
def api_sync_status():
    return jsonify(get_status())
