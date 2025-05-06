import os
import logging
from flask import Flask
from flask_session import Session
from flask_wtf.csrf import CSRFProtect

# Configure app
app = Flask(__name__)
app.secret_key = os.environ.get("SESSION_SECRET", "dev-secret-key-change-in-production")
app.config['SESSION_TYPE'] = 'filesystem'
app.config['SESSION_PERMANENT'] = False

# Enable CSRF protection
csrf = CSRFProtect(app)

# Setup session
Session(app)

# Configure logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Import routes after app is created to avoid circular imports
import routes
