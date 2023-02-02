*** Please follow the following instructions ***
1. Install postgresql v15.0 and above (Installation file is present in the directory: Installation/Postgres_Installation_Files)
2. While Installing keep the settings to default.
3. Important *** While installing please set the password for postgres as "9729" ***
4. After installing postgres run the following commands one by one in "psql prompt"
5. Note : psql prompt can be found by typing psql in start menu (Windows)
6. Important *** After copying and pasting each command hit enter ***

Commands:

1.

-- Database: po_nal_db

-- DROP DATABASE IF EXISTS po_nal_db;

CREATE DATABASE po_nal_db
    WITH
    OWNER = postgres
    ENCODING = 'UTF8'
    LC_COLLATE = 'English_India.1252'
    LC_CTYPE = 'English_India.1252'
    TABLESPACE = pg_default
    CONNECTION LIMIT = -1
    IS_TEMPLATE = False;

2.

-- SCHEMA: public

-- DROP SCHEMA IF EXISTS public ;

CREATE SCHEMA IF NOT EXISTS public
    AUTHORIZATION pg_database_owner;

COMMENT ON SCHEMA public
    IS 'standard public schema';

GRANT USAGE ON SCHEMA public TO PUBLIC;

GRANT ALL ON SCHEMA public TO pg_database_owner;

							***** Important: Please enter the following command before pasting next command in psql prompt : "\c po_nal_db" *****

3.

-- Table: public.purchase_details_tbl

-- DROP TABLE IF EXISTS public.purchase_details_tbl;

CREATE TABLE IF NOT EXISTS public.purchase_details_tbl
(
    indent_no character varying COLLATE pg_catalog."default" NOT NULL,
    date_raised date,
    item_descrp character varying COLLATE pg_catalog."default",
    division character varying COLLATE pg_catalog."default",
    indentor_name character varying COLLATE pg_catalog."default",
    delivery_date date,
    mode_of_procurement character varying COLLATE pg_catalog."default",
    specs bytea,
    amount_estimate numeric(15,2),
    status character varying COLLATE pg_catalog."default",
    actual_amount numeric(15,2),
    additional_info text COLLATE pg_catalog."default",
    extension character varying COLLATE pg_catalog."default",
    CONSTRAINT purchase_details_tbl_pkey PRIMARY KEY (indent_no)
)

TABLESPACE pg_default;

ALTER TABLE IF EXISTS public.purchase_details_tbl
    OWNER to postgres;


							***** Close the terminal and run the application: Installation finished *****


