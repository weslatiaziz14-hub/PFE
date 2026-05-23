# =============================================================================
# KPI DASHBOARD - SETUP & DEPLOYMENT GUIDE
# =============================================================================

## Project Structure

```
KPI_Digitalisation/
├── kpi_automation_system.py          # Modified to export to MySQL
├── seed_users.py                      # User seeding script
├── requirements.txt                   # Python dependencies for main script
├── database_schema.sql                # MySQL database schema
│
├── backend/                           # Flask backend
│   ├── __init__.py
│   ├── app.py                         # Main Flask application
│   ├── config.py                      # Configuration (DB, Secret key, etc.)
│   ├── db.py                          # SQLAlchemy initialization
│   ├── models.py                      # Database models (User)
│   └── requirements.txt               # Backend dependencies
│
├── frontend/                          # Static files
│   ├── templates/
│   │   ├── base.html                  # Base template (navbar, footer)
│   │   ├── login.html                 # Login page
│   │   ├── register.html              # Registration page (with dropdown)
│   │   └── dashboard.html             # Main dashboard page
│   └── static/
│       ├── css/
│       │   └── style.css              # Main stylesheet
│       └── js/
│           └── app.js                 # JavaScript utilities
│
├── data/                              # Excel source files (input)
│   ├── Suivi_defauts_CNC.xlsx
│   ├── Saisie_controle_Finale.xlsx
│   └── 01-Rapport_quantite_MMC.xlsx
│
└── output/                            # KPI Output (Excel, if using Excel mode)
```

---

## QUICK START

### 1. Setup Database

```bash
# Create MySQL database from schema
mysql -u root -p < database_schema.sql      # Enter password when prompted
```

**Update MySQL credentials in:**
- `kpi_automation_system.py` — lines ~50-56 (MYSQL_CONFIG)
- `backend/config.py` — lines ~29-34 (DB_HOST, DB_USER, DB_PASSWORD, DB_PORT)
- `seed_users.py` — lines ~35-41 (MYSQL_CONFIG)

### 2. Install Python Dependencies

```bash
# For main script
pip install -r requirements.txt

# For backend Flask app
cd backend
pip install -r requirements.txt
cd ..
```

### 3. Extract & Seed Users from Excel Files

```bash
python seed_users.py
```

**Output:**
- Extracts unique operator names from 3 Excel files
- Populates `users` table with unregistered employees
- Creates default manager account (username: "manager", password: "manager123")

### 4. Export KPI Data to MySQL

```bash
python kpi_automation_system.py
```

**Configuration:** Set `EXPORT_MODE = 'mysql'` in `kpi_automation_system.py` (line ~63)

**Output:**
- All 22 KPI tables populated in MySQL database `kpi_dashboard`
- Check with MySQL Workbench to verify

### 5. Start Flask Backend

```bash
cd backend
python app.py
```

**Output:**
```
 * Running on http://127.0.0.1:5000/ (Press CTRL+C to quit)
```

### 6. Access Dashboard

Open in browser: **http://localhost:5000**

---

## USER FLOW

### Employee Registration

1. Navigate to `/register`
2. **Dropdown:** Select your username (auto-populated from unregistered users)
3. **Password:** Set password (min 6 characters)
4. **Register:** Click button
5. ✅ Username removed from dropdown
6. Go to `/login` and login with credentials

### Employee Dashboard

1. Login with credentials
2. See personal KPI data (filtered by your name)
3. Browse tabs: Overview, CNC Quality, Final Control, MMC Production
4. View charts & tables with your metrics only

### Manager Dashboard

1. Login with manager account
2. **Username:** manager (default, can set in `seed_users.py`)
3. **Password:** manager123 (default, change after first login)
4. See **ALL** employees' data (no filtering)
5. Browse same tabs but access all data

---

## CONFIGURATION

### MySQL Connection

**File:** `backend/config.py`

```python
DB_HOST = 'localhost'        # Change if MySQL on different server
DB_USER = 'root'             # Your MySQL username
DB_PASSWORD = ''             # Your MySQL password (empty if none)
DB_NAME = 'kpi_dashboard'    # Database name (matches schema)
DB_PORT = 3306               # Standard MySQL port
```

### Flask Settings

**File:** `backend/config.py`

```python
SECRET_KEY = 'dev-secret-key-change-in-production'  # Change in production
DEBUG = True                                         # Set to False in production
SESSION_COOKIE_SECURE = False                       # Set to True with HTTPS
```

### Export Mode

**File:** `kpi_automation_system.py` (line ~63)

```python
EXPORT_MODE = 'mysql'    # Export to MySQL database
# EXPORT_MODE = 'excel'  # Export to Excel file (original)
```

---

## DATA FILTERING

### Employee Access

- Sees data **only where their name appears** in:
  - CNC: `kpi_suivi_journalier.controler_par`
  - Final Control: `kpi_cf_par_technicien.technicien`
  - MMC: `kpi_mmc_par_operateur.operateur`

### Manager Access

- Sees **ALL** data across all tables
- No name-based filtering applied
- Can view any employee's metrics

---

## CHARTS & DATA

### Available Charts

1. **RFT Trends** (Line) — Right First Time over time
2. **Error Rates** (Bar) — CNC & Assembly error rates
3. **Pareto Charts** (Bar) — Top defect types & corrective actions
4. **Production Trends** (Line) — Daily/monthly production
5. **Client Distribution** (Pie/Doughnut) — Production by client
6. **Scrap/Rework** (Stacked Bar) — Action classifications

### Data Tables

- Monthly KPI summaries
- Operator/Technician performance
- Test results
- All filterable by tabs

---

## TROUBLESHOOTING

### "MySQL Connection Failed"

1. Check MySQL is running locally
2. Verify credentials in `config.py`
3. Test connection:
   ```bash
   mysql -u root -p -e "SHOW DATABASES;"
   ```

### "No data in dashboard"

1. Run `python kpi_automation_system.py` to populate DB
2. Check Excel files in `data/` folder
3. Verify database tables:
   ```bash
   mysql kpi_dashboard -u root -p -e "SHOW TABLES;"
   ```

### "Can't login"

1. Check user is registered: `mysql kpi_dashboard -u root -p -e "SELECT * FROM users;"`
2. Check password is hashed (20+ characters in `password_hash`)
3. Verify `is_registered = 1` for your user

### "Dropdown is empty"

1. Run `python seed_users.py` again
2. Check Excel files contain operator names in expected columns:
   - CNC: `Controler par`, `detecter par`
   - Final Control: `Technicien QM`
   - MMC: `Opérateur`

### "Charts not showing"

1. Open browser console (F12) → Console tab
2. Check for JavaScript errors
3. Verify API endpoint `/api/dashboard-data` returns data:
   ```bash
   curl http://localhost:5000/api/dashboard-data
   ```

---

## SECURITY NOTES

⚠️ **For Development Use Only**

- Default manager password should be changed after first login
- Add `.env` file with secrets in production
- Use HTTPS in production (set `SESSION_COOKIE_SECURE = True`)
- Add `FLASK_ENV=production` environment variable
- Implement rate limiting on login/registration
- Consider adding 2FA for manager accounts

---

## DATABASE BACKUP

```bash
# Backup database
mysqldump -u root -p kpi_dashboard > kpi_backup.sql

# Restore database
mysql -u root -p kpi_dashboard < kpi_backup.sql
```

---

## NEXT STEPS (OPTIONAL ENHANCEMENTS)

- [ ] Add date range filters to charts
- [ ] Implement password reset functionality
- [ ] Add drill-down capability (click chart to see details)
- [ ] Implement real-time data sync
- [ ] Add data export to PDF/CSV
- [ ] Mobile-responsive UI improvements
- [ ] Add anomaly detection on KPI thresholds
- [ ] Deploy to production server (Heroku, AWS, etc.)

---

## SUPPORT

For issues, check:
1. MySQL Workbench connection status
2. Flask app logs (terminal output)
3. Browser console (F12) for JavaScript errors
4. Database query results for data verification

---

**Created:** March 2026
**Version:** 1.0
**Status:** Development Ready ✅
