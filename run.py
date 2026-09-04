"""Run the local commercial-proposal service on this computer only."""
from waitress import serve

from app import create_app

app = create_app()

if __name__ == "__main__":
    serve(app, host="127.0.0.1", port=8787, threads=4)
