# -*- coding: utf-8 -*-
# =============================================================================
# KPI TARGETS INITIALIZATION SCRIPT
# Initializes default KPI targets in the database using direct MySQL connection
# =============================================================================

import sys
import os
import mysql.connector
from mysql.connector import Error as MySQLError

# MySQL Configuration (should match backend config)
MYSQL_CONFIG = {
    'host': 'localhost',
    'user': 'root',
    'password': 'root',
    'database': 'kpi_dashboard',
    'port': 3306,
}

# Default targets
DEFAULT_TARGETS = [
    {
        'target_name': 'TARGET_FINISH_GOOD',
        'target_value': 40,
        'description': 'Finish Good pieces per day target',
        'category': 'MMC',
        'unit': 'pieces/day'
    },
    {
        'target_name': 'TARGET_STRETCH_FINISH_GOOD',
        'target_value': 45,
        'description': 'Finish Good stretch target',
        'category': 'MMC',
        'unit': 'pieces/day'
    },
    {
        'target_name': 'TARGET_INTERVENTION_LIMIT',
        'target_value': 35,
        'description': 'Intervention limit threshold',
        'category': 'MMC',
        'unit': 'pieces/day'
    },
    {
        'target_name': 'TARGET_ERROR_RATE_CNC',
        'target_value': 0.06,
        'description': 'Maximum error rate for CNC',
        'category': 'CNC',
        'unit': '%'
    },
    {
        'target_name': 'TARGET_ERROR_RATE_ASSEMBLY',
        'target_value': 0.10,
        'description': 'Maximum error rate for Assembly',
        'category': 'Assembly',
        'unit': '%'
    },
    {
        'target_name': 'TARGET_RFT_CNC',
        'target_value': 0.94,
        'description': 'Right First Time target for CNC',
        'category': 'CNC',
        'unit': '%'
    },
    {
        'target_name': 'TARGET_RFT_ASSEMBLY',
        'target_value': 0.90,
        'description': 'Right First Time target for Assembly',
        'category': 'Assembly',
        'unit': '%'
    },
    {
        'target_name': 'TARGET_RFT_MMC',
        'target_value': 0.80,
        'description': 'Right First Time target for Global/MMC',
        'category': 'MMC',
        'unit': '%'
    },
]


def init_targets():
    """Initialize default targets in database."""
    try:
        # Connect to database
        conn = mysql.connector.connect(**MYSQL_CONFIG)
        cursor = conn.cursor()
        
        # Create kpi_targets table if not exists
        create_table_sql = """
        CREATE TABLE IF NOT EXISTS kpi_targets (
            id INT PRIMARY KEY AUTO_INCREMENT,
            target_name VARCHAR(100) UNIQUE NOT NULL,
            target_value FLOAT NOT NULL,
            description VARCHAR(255),
            category VARCHAR(50),
            unit VARCHAR(50),
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
            INDEX idx_target_name (target_name),
            INDEX idx_category (category)
        )
        """
        cursor.execute(create_table_sql)
        print("[OK] kpi_targets table created/verified")
        
        # Check existing targets so we can insert only missing defaults
        cursor.execute("SELECT target_name FROM kpi_targets")
        existing_targets = {row[0] for row in cursor.fetchall()}
        
        # Insert default targets or missing ones
        insert_sql = """
        INSERT INTO kpi_targets (target_name, target_value, description, category, unit, created_at, updated_at)
        VALUES (%s, %s, %s, %s, %s, NOW(), NOW())
        """
        
        inserted_count = 0
        for target in DEFAULT_TARGETS:
            if target['target_name'] in existing_targets:
                continue
            cursor.execute(insert_sql, (
                target['target_name'],
                target['target_value'],
                target['description'],
                target['category'],
                target['unit']
            ))
            print(f"[+] Added: {target['target_name']} = {target['target_value']} {target['unit']}")
            inserted_count += 1
        
        if inserted_count > 0:
            conn.commit()
            print(f"\n[SUCCESS] Added {inserted_count} missing KPI target(s) to the database")
        else:
            print(f"[INFO] All default KPI targets already exist in the database.")
        
        cursor.close()
        conn.close()
        
    except MySQLError as err:
        print(f"[ERROR] MySQL error: {err}")
        sys.exit(1)
    except Exception as err:
        print(f"[ERROR] Failed to initialize targets: {err}")
        sys.exit(1)


if __name__ == '__main__':
    init_targets()
