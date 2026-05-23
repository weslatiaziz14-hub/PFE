# -*- coding: utf-8 -*-
# =============================================================================
# Database Initialization (SQLAlchemy)
# =============================================================================

from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


def init_db(app):
    """Initialize database with Flask app."""
    db.init_app(app)
    with app.app_context():
        db.create_all()
