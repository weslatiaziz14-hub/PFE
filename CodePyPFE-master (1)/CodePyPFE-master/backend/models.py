# -*- coding: utf-8 -*-
# =============================================================================
# Database Models (SQLAlchemy)
# =============================================================================

from flask_sqlalchemy import SQLAlchemy
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import UserMixin
from datetime import datetime

# Import db from db module
from db import db


class User(db.Model, UserMixin):
    """User model for authentication."""
    __tablename__ = 'users'
    
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(100), unique=True, nullable=False, index=True)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), default='employee', nullable=False, index=True)  # 'employee' or 'manager'
    is_registered = db.Column(db.Boolean, default=False, nullable=False, index=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def set_password(self, password):
        """Hash and set password."""
        self.password_hash = generate_password_hash(password, method='pbkdf2:sha256')
    
    def check_password(self, password):
        """Verify password against hash."""
        return check_password_hash(self.password_hash, password)
    
    def is_manager(self):
        """Check if user is manager."""
        return self.role == 'manager'
    
    def __repr__(self):
        return f'<User {self.username} ({self.role})>'


class KPITarget(db.Model):
    """KPI Target model for configurable targets."""
    __tablename__ = 'kpi_targets'
    
    id = db.Column(db.Integer, primary_key=True)
    target_name = db.Column(db.String(100), unique=True, nullable=False, index=True)
    target_value = db.Column(db.Float, nullable=False)
    description = db.Column(db.String(255))
    category = db.Column(db.String(50))  # e.g., 'CNC', 'Assembly', 'MMC'
    unit = db.Column(db.String(50))  # e.g., '%', 'pieces', 'pieces/day'
    created_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def to_dict(self):
        """Convert to dictionary for JSON serialization."""
        return {
            'id': self.id,
            'target_name': self.target_name,
            'target_value': self.target_value,
            'description': self.description,
            'category': self.category,
            'unit': self.unit,
            'updated_at': self.updated_at.isoformat() if self.updated_at else None
        }
    
    def __repr__(self):
        return f'<KPITarget {self.target_name} = {self.target_value} {self.unit}>'
