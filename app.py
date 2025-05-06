import os
import logging
from flask import Flask

# Configure app
app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key-change-in-production")

# We'll use Flask's built-in session functionality instead of flask_session
# and we'll handle CSRF protection manually

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Import routes after app is created to avoid circular imports
import routes
