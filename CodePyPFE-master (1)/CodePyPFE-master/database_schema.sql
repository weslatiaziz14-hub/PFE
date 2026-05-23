-- =============================================================================
-- KPI AUTOMATION SYSTEM - MySQL Schema
-- Created: 2026
-- =============================================================================
-- This schema stores all 22 KPI sheets from Excel output + user authentication
-- =============================================================================

-- Create database
CREATE DATABASE IF NOT EXISTS kpi_dashboard;
USE kpi_dashboard;

-- =============================================================================
-- USERS TABLE (Authentication & Authorization)
-- =============================================================================
CREATE TABLE users (
    id INT PRIMARY KEY AUTO_INCREMENT,
    username VARCHAR(100) UNIQUE NOT NULL,
    password_hash VARCHAR(255) NOT NULL,
    role ENUM('employee', 'manager') DEFAULT 'employee',
    is_registered BOOLEAN DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    INDEX idx_username (username),
    INDEX idx_role (role),
    INDEX idx_is_registered (is_registered)
);

-- =============================================================================
-- MODULE 1: SUIVI JOURNALIER (Daily Log)
-- =============================================================================
CREATE TABLE kpi_suivi_journalier (
    id INT PRIMARY KEY AUTO_INCREMENT,
    kw INT,
    mois VARCHAR(20),
    date DATE NOT NULL,
    detecter_par VARCHAR(100),
    controler_par VARCHAR(100),
    numero_order VARCHAR(50),
    statu VARCHAR(50),
    nombre_pieces_controlees INT,
    nombre_defauts INT,
    piece_nio VARCHAR(100),
    type_technique_defaut VARCHAR(100),
    description_defaut TEXT,
    numero_qrqc VARCHAR(50),
    action_correction TEXT,
    origine_defaut VARCHAR(50),
    etat VARCHAR(50),
    INDEX idx_date (date),
    INDEX idx_controler_par (controler_par),
    FOREIGN KEY (controler_par) REFERENCES users(username) ON DELETE SET NULL
);

-- =============================================================================
-- MODULE 2: CNC KPIs
-- =============================================================================
CREATE TABLE kpi_cnc_mensuel (
    id INT PRIMARY KEY AUTO_INCREMENT,
    annee INT NOT NULL,
    mois_n INT NOT NULL,
    mois VARCHAR(20),
    pieces_controlees INT,
    nombre_defauts INT,
    nb_ordres INT,
    nb_rft INT,
    taux_erreur_cnc_percent DECIMAL(10, 2),
    objectif_percent_pieces DECIMAL(10, 2),
    rft_cnc DECIMAL(10, 4),
    objectif_rft_cnc DECIMAL(10, 4),
    efficacite_cnc DECIMAL(10, 4),
    objectif_efficacite DECIMAL(10, 4),
    UNIQUE KEY uk_annee_mois_n (annee, mois_n),
    INDEX idx_annee (annee),
    INDEX idx_mois_n (mois_n)
);

CREATE TABLE kpi_cnc_rft_hebdomadaire (
    id INT PRIMARY KEY AUTO_INCREMENT,
    annee INT NOT NULL,
    kw INT NOT NULL,
    mois VARCHAR(20),
    pieces_controlees INT,
    nombre_defauts INT,
    nb_ordres INT,
    nb_rft INT,
    taux_erreur_percent DECIMAL(10, 2),
    rft_cnc DECIMAL(10, 4),
    objectif_rft DECIMAL(10, 4),
    UNIQUE KEY uk_annee_kw (annee, kw),
    INDEX idx_annee (annee),
    INDEX idx_kw (kw)
);

CREATE TABLE kpi_cnc_pareto_defauts (
    id INT PRIMARY KEY AUTO_INCREMENT,
    description_defaut VARCHAR(255) NOT NULL,
    nombre_defauts INT,
    cumul_percent DECIMAL(10, 1),
    INDEX idx_nombre_defauts (nombre_defauts)
);

CREATE TABLE kpi_cnc_pareto_actions (
    id INT PRIMARY KEY AUTO_INCREMENT,
    action_correction VARCHAR(255) NOT NULL,
    nombre_defauts INT,
    cumul_percent DECIMAL(10, 1),
    INDEX idx_nombre_defauts (nombre_defauts)
);

CREATE TABLE kpi_cnc_par_operateur (
    id INT PRIMARY KEY AUTO_INCREMENT,
    operateur VARCHAR(100) NOT NULL,
    pieces_controlees INT,
    nombre_defauts INT,
    taux_defauts_percent DECIMAL(10, 2),
    INDEX idx_operateur (operateur),
    FOREIGN KEY (operateur) REFERENCES users(username) ON DELETE SET NULL
);

CREATE TABLE kpi_cnc_par_statut (
    id INT PRIMARY KEY AUTO_INCREMENT,
    statut VARCHAR(50) NOT NULL,
    pieces_controlees INT,
    nombre_defauts INT,
    taux_defauts_percent DECIMAL(10, 2),
    INDEX idx_statut (statut)
);

CREATE TABLE kpi_cnc_par_piece (
    id INT PRIMARY KEY AUTO_INCREMENT,
    nom_piece VARCHAR(255) NOT NULL,
    pieces_controlees INT,
    nombre_defauts INT,
    taux_defauts_percent DECIMAL(10, 2),
    INDEX idx_nom_piece (nom_piece),
    INDEX idx_nombre_defauts (nombre_defauts)
);

CREATE TABLE kpi_cnc_par_origine (
    id INT PRIMARY KEY AUTO_INCREMENT,
    origine VARCHAR(100) NOT NULL,
    nombre_defauts INT,
    INDEX idx_nombre_defauts (nombre_defauts)
);

-- =============================================================================
-- MODULE 3: FINAL CONTROL KPIs
-- =============================================================================
CREATE TABLE kpi_cf_mensuel (
    id INT PRIMARY KEY AUTO_INCREMENT,
    annee INT NOT NULL,
    mois_n INT NOT NULL,
    mois VARCHAR(20),
    quantite_inspectee INT,
    nb_non_conformites INT,
    nb_conforme INT,
    nb_total INT,
    quantite_nio INT,
    nombre_non_conformites INT,
    taux_erreur_assembly_percent DECIMAL(10, 2),
    objectif_percent_module DECIMAL(10, 2),
    rft_assembly DECIMAL(10, 4),
    objectif_rft_assembly DECIMAL(10, 4),
    UNIQUE KEY uk_annee_mois_n (annee, mois_n),
    INDEX idx_annee (annee),
    INDEX idx_mois_n (mois_n)
);

CREATE TABLE kpi_cf_scrap_rework (
    id INT PRIMARY KEY AUTO_INCREMENT,
    annee INT NOT NULL,
    mois_n INT NOT NULL,
    mois VARCHAR(20),
    type_action VARCHAR(100),
    total INT,
    INDEX idx_annee_mois_n (annee, mois_n),
    INDEX idx_type_action (type_action)
);

CREATE TABLE kpi_cf_pareto_actions (
    id INT PRIMARY KEY AUTO_INCREMENT,
    action_correction VARCHAR(255) NOT NULL,
    quantite INT,
    cumul_percent DECIMAL(10, 1),
    INDEX idx_quantite (quantite)
);

CREATE TABLE kpi_cf_par_origine (
    id INT PRIMARY KEY AUTO_INCREMENT,
    origine VARCHAR(100) NOT NULL,
    nb_non_conformites INT,
    INDEX idx_nb_non_conformites (nb_non_conformites)
);

CREATE TABLE kpi_cf_par_technicien (
    id INT PRIMARY KEY AUTO_INCREMENT,
    technicien VARCHAR(100) NOT NULL,
    nb_non_conformites INT,
    nb_total INT,
    taux_nc_percent DECIMAL(10, 2),
    INDEX idx_technicien (technicien),
    FOREIGN KEY (technicien) REFERENCES users(username) ON DELETE SET NULL
);

-- =============================================================================
-- MODULE 4: MMC PRODUCTION KPIs
-- =============================================================================
CREATE TABLE kpi_mmc_journalier (
    id INT PRIMARY KEY AUTO_INCREMENT,
    date DATE NOT NULL,
    quantite_produite INT,
    nb_finish_good INT,
    nb_prototype INT,
    taux_prototypage DECIMAL(10, 4),
    tested_good INT,
    INDEX idx_date (date)
);

CREATE TABLE kpi_mmc_finishgood_mensuel (
    id INT PRIMARY KEY AUTO_INCREMENT,
    annee INT NOT NULL,
    mois_n INT NOT NULL,
    mois VARCHAR(20),
    quantite_produite INT,
    nb_prototype INT,
    nb_finish_good INT,
    tested_good INT,
    taux_prototypage DECIMAL(10, 4),
    quantity_target INT,
    quantity_stretch_target INT,
    quantity_intervention INT,
    objectif_rft_mmc DECIMAL(10, 4),
    UNIQUE KEY uk_annee_mois_n (annee, mois_n),
    INDEX idx_annee (annee),
    INDEX idx_mois_n (mois_n)
);

CREATE TABLE kpi_mmc_par_client (
    id INT PRIMARY KEY AUTO_INCREMENT,
    client VARCHAR(100) NOT NULL,
    quantite_produite INT,
    INDEX idx_quantite_produite (quantite_produite)
);

CREATE TABLE kpi_mmc_par_operateur (
    id INT PRIMARY KEY AUTO_INCREMENT,
    operateur VARCHAR(100) NOT NULL,
    quantite_produite INT,
    INDEX idx_operateur (operateur),
    FOREIGN KEY (operateur) REFERENCES users(username) ON DELETE SET NULL
);

CREATE TABLE kpi_mmc_tests (
    id INT PRIMARY KEY AUTO_INCREMENT,
    test VARCHAR(255) NOT NULL,
    total_pieces INT,
    reussi INT,
    non_teste INT,
    taux_reussite_percent DECIMAL(10, 2),
    objectif_percent DECIMAL(10, 2),
    INDEX idx_test (test)
);

-- =============================================================================
-- MODULE 5: GLOBAL KPIs
-- =============================================================================
CREATE TABLE kpi_kpi_mensuel (
    id INT PRIMARY KEY AUTO_INCREMENT,
    annee INT NOT NULL,
    mois_n INT NOT NULL,
    mois VARCHAR(20),
    -- CNC columns
    pieces_controlees_cnc INT,
    nombre_defauts_cnc INT,
    nb_ordres_cnc INT,
    nb_rft_cnc INT,
    taux_erreur_cnc_percent DECIMAL(10, 2),
    rft_cnc DECIMAL(10, 4),
    -- Final Control columns
    quantite_inspectee INT,
    nb_non_conformites INT,
    nb_conforme INT,
    nb_total INT,
    taux_erreur_assembly_percent DECIMAL(10, 2),
    rft_assembly DECIMAL(10, 4),
    -- MMC columns
    quantite_produite INT,
    nb_prototype INT,
    nb_finish_good INT,
    tested_good INT,
    -- Global metrics
    taux_defauts_global_percent DECIMAL(10, 2),
    UNIQUE KEY uk_annee_mois_n (annee, mois_n),
    INDEX idx_annee (annee),
    INDEX idx_mois_n (mois_n)
);

CREATE TABLE kpi_rft_global (
    id INT PRIMARY KEY AUTO_INCREMENT,
    annee INT NOT NULL,
    mois_n INT NOT NULL,
    mois VARCHAR(20),
    rft_cnc DECIMAL(10, 4),
    objectif_rft_cnc DECIMAL(10, 4),
    rft_assembly DECIMAL(10, 4),
    objectif_rft_assembly DECIMAL(10, 4),
    rft_global DECIMAL(10, 4),
    objectif_rft_mmc DECIMAL(10, 4),
    UNIQUE KEY uk_annee_mois_n (annee, mois_n),
    INDEX idx_annee (annee),
    INDEX idx_mois_n (mois_n)
);

CREATE TABLE kpi_kpi_global (
    id INT PRIMARY KEY AUTO_INCREMENT,
    annee INT NOT NULL,
    quantite_produite INT,
    pieces_controlees INT,
    nombre_defauts INT,
    nb_non_conformites INT,
    taux_defauts_percent DECIMAL(10, 2),
    taux_nc_percent DECIMAL(10, 2),
    UNIQUE KEY uk_annee (annee),
    INDEX idx_annee (annee)
);

-- =============================================================================
-- Create base users for testing (optional - can be done via seed script)
-- =============================================================================
-- INSERT INTO users (username, password_hash, role, is_registered) 
-- VALUES ('manager', 'hashed_password_here', 'manager', 1);
