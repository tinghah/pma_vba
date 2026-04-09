-- PostgreSQL Combined Script
-- This script creates three independent tables with all columns included from your sheets.
-- Composite Primary Keys (id_no + date) are used to allow yearly data and daily uploads.

-- 1. Table: leave
CREATE TABLE IF NOT EXISTS leave (
    id_no INTEGER,
    english_name VARCHAR(255),
    onboard_date DATE,
    resign_date DATE,
    factory VARCHAR(50),
    group_code VARCHAR(50),
    department VARCHAR(100),
    leave_date DATE,
    leave_hour NUMERIC(5, 2),
    total NUMERIC(10, 2),
    casual NUMERIC(5, 2),
    unpaid NUMERIC(5, 2),
    ssb_sick NUMERIC(5, 2),
    sick NUMERIC(5, 2),
    earned NUMERIC(5, 2),
    og NUMERIC(5, 2),
    maternity NUMERIC(5, 2),
    paternity NUMERIC(5, 2),
    official NUMERIC(5, 2),
    paid_injury NUMERIC(5, 2),
    unpaid_injury NUMERIC(5, 2),
    other NUMERIC(5, 2),
    business NUMERIC(5, 2),
    absent NUMERIC(5, 2),
    PRIMARY KEY (id_no, leave_date)
);

-- 2. Table: ot
CREATE TABLE IF NOT EXISTS ot (
    id_no INTEGER,
    english_name VARCHAR(255),
    onboard_date DATE,
    resign_date DATE,
    factory VARCHAR(50),
    group_code VARCHAR(50),
    department VARCHAR(100),
    ot_date DATE,
    ot_hour NUMERIC(5, 2),
    total NUMERIC(10, 2),
    PRIMARY KEY (id_no, ot_date)
);

-- 3. Table: late
CREATE TABLE IF NOT EXISTS late (
    id_no INTEGER,
    english_name VARCHAR(255),
    onboard_date DATE,
    resign_date DATE,
    factory VARCHAR(50),
    group_code VARCHAR(50),
    department VARCHAR(100),
    late_date DATE,
    late_value NUMERIC(5, 2),
    total NUMERIC(10, 2),
    PRIMARY KEY (id_no, late_date)
);

-- Performance Indexes for Yearly Queries
CREATE INDEX idx_leave_date ON leave(leave_date);
CREATE INDEX idx_ot_date ON ot(ot_date);
CREATE INDEX idx_late_date ON late(late_date);