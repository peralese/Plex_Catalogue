from app.app import app

if __name__ == "__main__":
    # Bind to all interfaces so devices on the local network can connect.
    # Access from other machines via http://<mac-mini-hostname>.local:5000
    app.run(host="0.0.0.0", port=5000, debug=False)
