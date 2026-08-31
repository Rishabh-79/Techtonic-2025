import sys
import os

# Ensure the project root is on the path so Flask can find templates/ and static/
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import app

# Vercel expects a WSGI callable named `app`
# Flask's app object is already a WSGI callable — no extra wrapper needed.
