"""Entry point: python run.py"""
import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

from web.app import create_app

if __name__ == "__main__":
    app = create_app()
    app.run(debug=True, host="0.0.0.0", port=5000, threaded=True)
