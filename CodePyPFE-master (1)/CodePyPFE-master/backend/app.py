# -*- coding: utf-8 -*-
# =============================================================================
# Flask Application - KPI Dashboard Backend
# =============================================================================

import os
import sys
import re
import json
import threading
import subprocess
import time
from flask import Flask, render_template, request, redirect, url_for, session, jsonify
from flask_login import LoginManager, login_user, logout_user, login_required, current_user
from flask_cors import CORS
from werkzeug.security import generate_password_hash
import mysql.connector
from mysql.connector import Error as MySQLError

# Ensure project root is on sys.path so we can import KPI formula definitions
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from config import config
from db import db, init_db
from models import User, KPITarget

try:
    from kpi_automation_system import CHART_FORMULAS
except ImportError:
    CHART_FORMULAS = {}

ROOT_PATH = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))


def run_external_script(script_name):
    script_path = os.path.join(ROOT_PATH, script_name)
    if not os.path.exists(script_path):
        print(f"[TASK] Script not found: {script_path}")
        return
    try:
        print(f"[TASK] Running {script_name}...")
        completed = subprocess.run(
            [sys.executable, script_path],
            cwd=ROOT_PATH,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False
        )
        if completed.returncode != 0:
            print(f"[TASK] {script_name} failed with exit code {completed.returncode}")
            print(completed.stderr)
        else:
            print(f"[TASK] {script_name} finished successfully")
            if completed.stdout:
                print(completed.stdout)
    except Exception as err:
        print(f"[TASK] Failed to run {script_name}: {err}")


def run_periodic_automation():
    scripts = ['seed_users.py', 'kpi_automation_system.py']
    while True:
        for script in scripts:
            run_external_script(script)
        time.sleep(300)


# Initialize Flask app
def create_app(config_name='development'):
    app = Flask(__name__, template_folder='../frontend/templates', static_folder='../frontend/static')
    app.config.from_object(config[config_name])
    
    # Initialize extensions
    CORS(app)
    
    # Initialize database (this also calls db.init_app(app))
    with app.app_context():
        init_db(app)
    
    # Initialize Flask-Login
    login_manager = LoginManager()
    login_manager.init_app(app)
    login_manager.login_view = 'login'
    login_manager.login_message = 'Please log in to access this page.'
    
    @login_manager.user_loader
    def load_user(user_id):
        return User.query.get(int(user_id))
    
    # =====================================================================
    # UTILITY FUNCTIONS
    # =====================================================================
    
    def get_db_connection():
        """Get MySQL connection for raw queries."""
        try:
            conn = mysql.connector.connect(
                host=app.config['DB_HOST'],
                user=app.config['DB_USER'],
                password=app.config['DB_PASSWORD'],
                database=app.config['DB_NAME'],
                port=app.config['DB_PORT']
            )
            return conn
        except MySQLError as err:
            print(f"MySQL connection error: {err}")
            return None
    
    def get_unregistered_users():
        """Get list of unregistered users (for registration dropdown)."""
        conn = get_db_connection()
        if not conn:
            return []
        try:
            cursor = conn.cursor(dictionary=True)
            cursor.execute("SELECT username FROM users WHERE is_registered = 0 ORDER BY username")
            users = [row['username'] for row in cursor.fetchall()]
            cursor.close()
            return users
        except MySQLError as err:
            print(f"Error fetching users: {err}")
            return []
        finally:
            conn.close()
    def build_employee_match_clause(column, username):
        """Build SQL clause to match usernames within combined names like 'wahid/sami' or 'wahid+sami'."""
        safe_name = re.sub(r"([\\.^$*+?{}\[\]\\|()])", r"\\\\\1", username.strip().lower())
        return f"LOWER({column}) REGEXP '(^|[ /+]){safe_name}($|[ /+])'"
    
    def get_dashboard_data(user, year=None):
        """Get KPI data for user (filtered by username if employee, all if manager, and optionally by year)."""
        conn = get_db_connection()
        if not conn:
            return {}
        
        try:
            cursor = conn.cursor(dictionary=True)
            data = {}
            
            # Year filter clause for tables that contain annee
            year_filter = f" AND annee = {year}" if year else ""

            if user.is_manager():
                queries = {
                    'suivi_journalier': f"SELECT * FROM kpi_suivi_journalier WHERE 1=1{(' AND YEAR(date) = ' + str(year)) if year else ''} ORDER BY date DESC",
                    'cnc_mensuel': f"SELECT *, CONCAT(annee, '-', LPAD(mois_n, 2, '0'), ' (', mois, ')') as period FROM kpi_cnc_mensuel WHERE 1=1{year_filter} ORDER BY annee DESC, mois_n DESC",
                    'cnc_rft_hebdomadaire': f"SELECT *, CONCAT(annee, '-W', LPAD(kw, 2, '0')) as period FROM kpi_cnc_rft_hebdomadaire WHERE 1=1{year_filter} ORDER BY annee DESC, kw DESC",
                    'cnc_pareto_defauts': "SELECT * FROM kpi_cnc_pareto_defauts ORDER BY nombre_defauts DESC LIMIT 20",
                    'cnc_pareto_actions': "SELECT * FROM kpi_cnc_pareto_actions ORDER BY nombre_defauts DESC LIMIT 20",
                    'cnc_par_operateur': "SELECT * FROM kpi_cnc_par_operateur",
                    'cnc_par_statut': "SELECT * FROM kpi_cnc_par_statut",
                    'cnc_par_piece': "SELECT * FROM kpi_cnc_par_piece ORDER BY nombre_defauts DESC LIMIT 20",
                    'cnc_par_origine': "SELECT * FROM kpi_cnc_par_origine",
                    'cf_mensuel': f"SELECT *, CONCAT(annee, '-', LPAD(mois_n, 2, '0'), ' (', mois, ')') as period FROM kpi_cf_mensuel WHERE 1=1{year_filter} ORDER BY annee DESC, mois_n DESC",
                    'cf_scrap_rework': f"SELECT annee, mois_n, mois, type_action, total, CONCAT(annee, '-', LPAD(mois_n, 2, '0'), ' (', mois, ')') as period FROM kpi_cf_scrap_rework WHERE 1=1{year_filter} ORDER BY annee DESC, mois_n DESC",
                    'cf_pareto_actions': "SELECT * FROM kpi_cf_pareto_actions ORDER BY quantite DESC LIMIT 20",
                    'cf_par_origine': "SELECT * FROM kpi_cf_par_origine",
                    'cf_par_technicien': "SELECT * FROM kpi_cf_par_technicien",
                    'mmc_journalier': f"SELECT * FROM kpi_mmc_journalier WHERE 1=1{(' AND YEAR(date) = ' + str(year)) if year else ''} ORDER BY date DESC",
                    'mmc_finishgood_mensuel': f"SELECT *, CONCAT(annee, '-', LPAD(mois_n, 2, '0'), ' (', mois, ')') as period FROM kpi_mmc_finishgood_mensuel WHERE 1=1{year_filter} ORDER BY annee DESC, mois_n DESC",
                    'mmc_par_client': "SELECT * FROM kpi_mmc_par_client ORDER BY quantite_produite DESC",
                    'mmc_par_operateur': "SELECT * FROM kpi_mmc_par_operateur",
                    'mmc_tests': "SELECT * FROM kpi_mmc_tests",
                    'kpi_mensuel': f"SELECT *, CONCAT(annee, '-', LPAD(mois_n, 2, '0'), ' (', mois, ')') as period FROM kpi_kpi_mensuel WHERE 1=1{year_filter} ORDER BY annee DESC, mois_n DESC",
                    'rft_global': f"SELECT *, CONCAT(annee, '-', LPAD(mois_n, 2, '0'), ' (', mois, ')') as period FROM kpi_rft_global WHERE 1=1{year_filter} ORDER BY annee DESC, mois_n DESC",
                    'kpi_global': f"SELECT * FROM kpi_kpi_global WHERE 1=1{year_filter} ORDER BY annee DESC",
                    'cnc_journalier': f"SELECT * FROM kpi_cnc_journalier WHERE 1=1{(' AND YEAR(date) = ' + str(year)) if year else ''} ORDER BY date DESC LIMIT 500",
                    'cf_journalier':  f"SELECT * FROM kpi_cf_journalier  WHERE 1=1{(' AND YEAR(date) = ' + str(year)) if year else ''} ORDER BY date DESC LIMIT 500",
                }
            else:
                lower_username = user.username.strip().lower()
                operator_clause = build_employee_match_clause('operateur', lower_username)
                technician_clause = build_employee_match_clause('technicien', lower_username)

                queries = {
                    'cnc_par_operateur': f"SELECT * FROM kpi_cnc_par_operateur WHERE {operator_clause}",
                    'cf_par_technicien': f"SELECT * FROM kpi_cf_par_technicien WHERE {technician_clause}",
                    'mmc_par_operateur': f"SELECT * FROM kpi_mmc_par_operateur WHERE {operator_clause}",
                }
            
            # Execute standard queries
            for table_name, query in queries.items():
                try:
                    cursor.execute(query)
                    rows = cursor.fetchall()
                    data[table_name] = rows
                except MySQLError as err:
                    print(f"Error querying {table_name}: {err}")
                    data[table_name] = []
            
            # Dynamically fetch additional data files tables (kpi_additional_*)
            if user.is_manager():
                try:
                    cursor.execute(
                        "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES "
                        "WHERE TABLE_SCHEMA = %s AND TABLE_NAME LIKE 'kpi_additional_%'",
                        (app.config['DB_NAME'],)
                    )
                    additional_tables = cursor.fetchall()
                    
                    for table_row in additional_tables:
                        table_name = table_row['TABLE_NAME']
                        # Use snake_case as the key (remove 'kpi_additional_' prefix for clean keys)
                        key = table_name.replace('kpi_additional_', '')
                        try:
                            cursor.execute(f"SELECT * FROM `{table_name}` LIMIT 1000")
                            rows = cursor.fetchall()
                            data[key] = rows
                        except MySQLError as err:
                            print(f"Error querying {table_name}: {err}")
                except MySQLError as err:
                    print(f"Error discovering additional tables: {err}")
            
            cursor.close()
            return data
        except Exception as err:
            print(f"Error getting dashboard data: {err}")
            return {}
        finally:
            conn.close()
    
    # =====================================================================
    # ROUTES - AUTHENTICATION
    # =====================================================================
    
    @app.route('/login', methods=['GET', 'POST'])
    def login():
        """Login route."""
        if request.method == 'POST':
            username = request.form.get('username', '').strip()
            password = request.form.get('password', '').strip()
            
            if not username or not password:
                return render_template('login.html', error='Username and password are required'), 400
            
            # Query database
            user = User.query.filter_by(username=username, is_registered=True).first()
            
            if user and user.check_password(password):
                login_user(user, remember=True)
                session.permanent = True
                return redirect(url_for('dashboard'))
            else:
                return render_template('login.html', error='Invalid username or password'), 401
        
        return render_template('login.html')
    
    @app.route('/register', methods=['GET', 'POST'])
    def register():
        """Registration route - SIMPLIFIED: username + password only."""
        if request.method == 'POST':
            username = request.form.get('username', '').strip()
            password = request.form.get('password', '').strip()
            password_confirm = request.form.get('password_confirm', '').strip()

            if not username:
                return render_template('register.html',
                    error="Nom d'utilisateur requis"), 400

            if not password or not password_confirm:
                return render_template('register.html',
                    error='Mot de passe requis'), 400

            if password != password_confirm:
                return render_template('register.html',
                    error='Les mots de passe ne correspondent pas'), 400

            if len(password) < 6:
                return render_template('register.html',
                    error='Le mot de passe doit contenir au moins 6 caractères'), 400

            user = User.query.filter_by(username=username, is_registered=False).first()

            if not user:
                existing = User.query.filter_by(username=username).first()
                if existing and existing.is_registered:
                    return render_template('register.html',
                        error='Ce nom d\'utilisateur est déjà enregistré'), 400
                else:
                    return render_template('register.html',
                        error='Nom d\'utilisateur non trouvé. Contactez votre manager.'), 400

            user.set_password(password)
            user.is_registered = True
            db.session.commit()

            return render_template('register.html',
                success='Inscription réussie ! Vous pouvez maintenant vous connecter.'), 200

        return render_template('register.html')
    
    @app.route('/logout')
    @login_required
    def logout():
        """Logout route."""
        logout_user()
        return redirect(url_for('login'))
    
    # =====================================================================
    # ROUTES - DASHBOARD
    # =====================================================================
    
    @app.route('/')
    def index():
        """Index route - redirect to dashboard or login."""
        if current_user.is_authenticated:
            return redirect(url_for('dashboard'))
        return redirect(url_for('login'))
    
    @app.route('/dashboard')
    @login_required
    def dashboard():
        """Dashboard main page."""
        is_manager = current_user.is_manager()
        return render_template('dashboard.html', 
            username=current_user.username,
            is_manager=is_manager,
            chart_formulas=CHART_FORMULAS)
    
    # =====================================================================
    # API ROUTES - DATA
    # =====================================================================
    
    @app.route('/api/unregistered-users', methods=['GET'])
    def api_unregistered_users():
        """Get list of unregistered users for registration dropdown."""
        users = get_unregistered_users()
        return jsonify(users)
    
    @app.route('/api/dashboard-data', methods=['GET'])
    @login_required
    def api_dashboard_data():
        """Get KPI data for current user, optionally filtered by year."""
        year = request.args.get('year', None)
        if year:
            try:
                year = int(year)
            except ValueError:
                return jsonify({'error': 'Invalid year parameter'}), 400
        data = get_dashboard_data(current_user, year=year)
        return jsonify(data)
    
    @app.route('/api/available-years', methods=['GET'])
    @login_required
    def api_available_years():
        """Get list of available years in the database."""
        conn = get_db_connection()
        if not conn:
            return jsonify([]), 500
        
        try:
            cursor = conn.cursor(dictionary=True)
            # Get distinct years from all relevant year-based tables
            cursor.execute(
                "SELECT DISTINCT annee FROM kpi_cnc_mensuel "
                "UNION SELECT DISTINCT annee FROM kpi_cnc_rft_hebdomadaire "
                "UNION SELECT DISTINCT annee FROM kpi_cf_mensuel "
                "UNION SELECT DISTINCT annee FROM kpi_cf_scrap_rework "
                "UNION SELECT DISTINCT annee FROM kpi_mmc_finishgood_mensuel "
                "UNION SELECT DISTINCT annee FROM kpi_kpi_mensuel "
                "UNION SELECT DISTINCT annee FROM kpi_rft_global "
                "UNION SELECT DISTINCT annee FROM kpi_kpi_global "
                "ORDER BY annee DESC"
            )
            rows = cursor.fetchall()
            cursor.close()
            years = [row['annee'] for row in rows]
            return jsonify(years)
        except Exception as err:
            print(f"Error getting available years: {err}")
            return jsonify([]), 500
        finally:
            conn.close()
    
    @app.route('/api/additional-tables', methods=['GET'])
    @login_required
    def api_additional_tables():
        """Get list of available additional data tables."""
        if not current_user.is_manager():
            return jsonify([]), 403
        
        conn = get_db_connection()
        if not conn:
            return jsonify([]), 500
        
        try:
            cursor = conn.cursor(dictionary=True)
            cursor.execute(
                "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES "
                "WHERE TABLE_SCHEMA = %s AND TABLE_NAME LIKE 'kpi_additional_%' "
                "ORDER BY TABLE_NAME",
                (app.config['DB_NAME'],)
            )
            rows = cursor.fetchall()
            cursor.close()
            
            # Return clean names (remove 'kpi_additional_' prefix)
            tables = [row['TABLE_NAME'].replace('kpi_additional_', '') for row in rows]
            return jsonify(tables)
        except Exception as err:
            print(f"Error getting additional tables: {err}")
            return jsonify([]), 500
        finally:
            conn.close()
    
    @app.route('/api/user-info', methods=['GET'])
    @login_required
    def api_user_info():
        """Get current user info."""
        return jsonify({
            'id': current_user.id,
            'username': current_user.username,
            'role': current_user.role,
            'is_manager': current_user.is_manager()
        })
    
    # =====================================================================
    # API ROUTES - KPI TARGETS (NEW)
    # =====================================================================
    
    @app.route('/api/targets', methods=['GET'])
    @login_required
    def api_get_targets():
        """Get all KPI targets."""
        try:
            targets = KPITarget.query.all()
            return jsonify([target.to_dict() for target in targets])
        except Exception as err:
            print(f"Error getting targets: {err}")
            return jsonify({'error': 'Failed to retrieve targets'}), 500
    
    @app.route('/api/targets/<int:target_id>', methods=['GET'])
    @login_required
    def api_get_target(target_id):
        """Get a specific target by ID."""
        try:
            target = KPITarget.query.get(target_id)
            if not target:
                return jsonify({'error': 'Target not found'}), 404
            return jsonify(target.to_dict())
        except Exception as err:
            print(f"Error getting target {target_id}: {err}")
            return jsonify({'error': 'Failed to retrieve target'}), 500
    
    @app.route('/api/targets/<int:target_id>', methods=['PUT'])
    @login_required
    def api_update_target(target_id):
        """Update a target value (managers only)."""
        if not current_user.is_manager():
            return jsonify({'error': 'Only managers can update targets'}), 403
        
        try:
            target = KPITarget.query.get(target_id)
            if not target:
                return jsonify({'error': 'Target not found'}), 404
            
            data = request.get_json()
            if not data:
                return jsonify({'error': 'No data provided'}), 400
            
            # Update target value
            if 'target_value' in data:
                target.target_value = float(data['target_value'])
            
            # Update description if provided
            if 'description' in data:
                target.description = data['description']
            
            db.session.commit()
            return jsonify({
                'success': True,
                'message': f'Target {target.target_name} updated successfully',
                'target': target.to_dict()
            })
        except ValueError as err:
            return jsonify({'error': 'Invalid target value format'}), 400
        except Exception as err:
            db.session.rollback()
            print(f"Error updating target {target_id}: {err}")
            return jsonify({'error': 'Failed to update target'}), 500
    
    @app.route('/api/trigger-automation', methods=['POST'])
    @login_required
    def api_trigger_automation():
        """Trigger KPI automation immediately (managers only)."""
        if not current_user.is_manager():
            return jsonify({'error': 'Only managers can trigger automation'}), 403
        
        try:
            # Run the automation scripts immediately in a background thread
            automation_thread = threading.Thread(
                target=run_external_script,
                args=('kpi_automation_system.py',),
                daemon=True
            )
            automation_thread.start()
            
            return jsonify({
                'success': True,
                'message': 'KPI automation triggered successfully. Data will be updated shortly.'
            })
        except Exception as err:
            print(f"Error triggering automation: {err}")
            return jsonify({'error': 'Failed to trigger automation'}), 500
    
    # =====================================================================
    # ERROR HANDLERS
    # =====================================================================
    
    @app.errorhandler(404)
    def not_found_error(error):
        """404 error handler."""
        return render_template('404.html'), 404
    
    @app.errorhandler(500)
    def internal_error(error):
        """500 error handler."""
        return render_template('500.html'), 500
    
    return app


# =========================================================================
# APPLICATION ENTRY POINT
# =========================================================================

if __name__ == '__main__':
    config_name = os.environ.get('FLASK_ENV', 'development')
    app = create_app(config_name)

    # Start periodic automation only in the real application process
    if not app.debug or os.environ.get('WERKZEUG_RUN_MAIN') == 'true':
        automation_thread = threading.Thread(target=run_periodic_automation, daemon=True)
        automation_thread.start()

    app.run(host='127.0.0.1', port=5000, debug=config_name == 'development')
