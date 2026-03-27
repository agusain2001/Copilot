BEGIN TRANSACTION;
CREATE TABLE exchange_rates (
	id UUID NOT NULL, 
	from_currency VARCHAR(10) NOT NULL, 
	to_currency VARCHAR(10) NOT NULL, 
	rate NUMERIC(18, 6) NOT NULL, 
	effective_date DATE NOT NULL, 
	PRIMARY KEY (id)
);
CREATE TABLE report_sequences (
	id INTEGER NOT NULL, 
	created_at DATETIME DEFAULT (CURRENT_TIMESTAMP) NOT NULL, 
	PRIMARY KEY (id)
);
INSERT INTO "report_sequences" VALUES(1,'2026-03-20 08:26:14.930407');
INSERT INTO "report_sequences" VALUES(2,'2026-03-20 08:27:27.795499');
INSERT INTO "report_sequences" VALUES(3,'2026-03-20 08:51:54.411087');
CREATE TABLE report_types (
	id UUID NOT NULL, 
	name VARCHAR(50) NOT NULL, 
	label VARCHAR(100) NOT NULL, 
	sort_order INTEGER NOT NULL, 
	is_active BOOLEAN NOT NULL, 
	PRIMARY KEY (id), 
	UNIQUE (name)
);
INSERT INTO "report_types" VALUES('a972dc2c-10e3-4804-b8a7-ec50c62f197a','alpha','Alpha',1,1);
INSERT INTO "report_types" VALUES('0c9e7197-7179-4981-b237-5511fd732ce7','beta','Beta',2,1);
INSERT INTO "report_types" VALUES('079f7dd4-e603-478f-b5b8-35aba5b33e2d','gamma','Gamma',3,1);
INSERT INTO "report_types" VALUES('5684f12c-6a57-4992-9976-385e4ca2db2b','theta','Theta',4,1);
CREATE TABLE reports (
	id UUID NOT NULL, 
	name VARCHAR(300) NOT NULL, 
	type_id UUID NOT NULL, 
	file_url VARCHAR NOT NULL, 
	original_filename VARCHAR(300) NOT NULL, 
	file_size_bytes BIGINT, 
	uploaded_by UUID NOT NULL, 
	uploaded_at DATETIME DEFAULT (CURRENT_TIMESTAMP) NOT NULL, 
	is_deleted BOOLEAN NOT NULL, 
	deleted_at DATETIME, 
	sequence_id INTEGER, 
	file_type VARCHAR(10), 
	stored_filename VARCHAR(400), 
	folder_path VARCHAR(700), 
	PRIMARY KEY (id), 
	FOREIGN KEY(type_id) REFERENCES report_types (id) ON DELETE RESTRICT, 
	FOREIGN KEY(uploaded_by) REFERENCES users (id) ON DELETE CASCADE, 
	FOREIGN KEY(sequence_id) REFERENCES report_sequences (id) ON DELETE SET NULL
);
INSERT INTO "reports" VALUES('4c1a34fe-147e-4435-ab8c-641f00d0653b','project_tracker_conditional_formatting','a972dc2c-10e3-4804-b8a7-ec50c62f197a','uploads\reports\f0e25e69-1f86-4f23-b612-e1d3f1a2619c.xlsx','project_tracker_conditional_formatting.xlsx',8149,'9c47c0b2-8c37-4e05-a927-1ec805cb8623','2026-03-20 07:41:24.721902',1,'2026-03-20 07:57:03.579112',NULL,NULL,NULL,NULL);
INSERT INTO "reports" VALUES('ea3267a8-e232-45b9-a4eb-096636706465','report','a972dc2c-10e3-4804-b8a7-ec50c62f197a','uploads\reports\de1c418b-71a3-4745-850d-753e0fd488c2.xlsx','project_tracker_conditional_formatting.xlsx',8149,'9c47c0b2-8c37-4e05-a927-1ec805cb8623','2026-03-20 07:56:15.305443',1,'2026-03-20 07:57:03.579112',NULL,NULL,NULL,NULL);
INSERT INTO "reports" VALUES('3adb540e-f73e-414d-b3e8-1b376648ecae','report','a972dc2c-10e3-4804-b8a7-ec50c62f197a','uploads\reports\1\inputs\20032026082617_project_tracker_conditional_formatting.xlsx','project_tracker_conditional_formatting.xlsx',8149,'9c47c0b2-8c37-4e05-a927-1ec805cb8623','2026-03-20 08:26:14.930407',0,NULL,1,'input','20032026082617_project_tracker_conditional_formatting.xlsx','uploads\reports\1\inputs\20032026082617_project_tracker_conditional_formatting.xlsx');
INSERT INTO "reports" VALUES('551f5d52-dbb3-4d1a-983c-0ec601537425','reports','a972dc2c-10e3-4804-b8a7-ec50c62f197a','uploads\reports\2\inputs\20032026082728_project_tracker_conditional_formatting.xlsx','project_tracker_conditional_formatting.xlsx',8149,'9c47c0b2-8c37-4e05-a927-1ec805cb8623','2026-03-20 08:27:27.795499',0,NULL,2,'input','20032026082728_project_tracker_conditional_formatting.xlsx','uploads\reports\2\inputs\20032026082728_project_tracker_conditional_formatting.xlsx');
INSERT INTO "reports" VALUES('5ad1c8ff-b979-449a-b400-114c6e645a16','report','a972dc2c-10e3-4804-b8a7-ec50c62f197a','uploads\reports\3\inputs\20032026142155_report.xlsx','report.xlsx',8149,'9c47c0b2-8c37-4e05-a927-1ec805cb8623','2026-03-20 08:51:54.411087',0,NULL,3,'input','20032026142155_report.xlsx','uploads\reports\3\inputs\20032026142155_report.xlsx');
CREATE TABLE shops (
	id UUID NOT NULL, 
	name VARCHAR(200) NOT NULL, 
	country VARCHAR(100) NOT NULL, 
	currency_code VARCHAR(10) NOT NULL, 
	is_active BOOLEAN NOT NULL, 
	created_at DATETIME DEFAULT (CURRENT_TIMESTAMP) NOT NULL, 
	PRIMARY KEY (id)
);
CREATE TABLE user_preferences (
	id UUID NOT NULL, 
	user_id UUID NOT NULL, 
	theme VARCHAR(10) NOT NULL, 
	updated_at DATETIME DEFAULT (CURRENT_TIMESTAMP) NOT NULL, 
	PRIMARY KEY (id), 
	UNIQUE (user_id), 
	FOREIGN KEY(user_id) REFERENCES users (id) ON DELETE CASCADE
);
INSERT INTO "user_preferences" VALUES('87f08890-e036-48e1-b157-fb431196f6b9','9c47c0b2-8c37-4e05-a927-1ec805cb8623','light','2026-03-20 06:03:12.008539');
CREATE TABLE users (
	id UUID NOT NULL, 
	username VARCHAR(100) NOT NULL, 
	full_name VARCHAR(200) NOT NULL, 
	employee_id VARCHAR(50) NOT NULL, 
	email VARCHAR(200) NOT NULL, 
	phone VARCHAR(20), 
	password_hash VARCHAR NOT NULL, 
	profile_photo_url VARCHAR, 
	password_changed_at DATETIME, 
	is_active BOOLEAN NOT NULL, 
	created_at DATETIME DEFAULT (CURRENT_TIMESTAMP) NOT NULL, 
	updated_at DATETIME DEFAULT (CURRENT_TIMESTAMP) NOT NULL, 
	PRIMARY KEY (id)
);
INSERT INTO "users" VALUES('9c47c0b2-8c37-4e05-a927-1ec805cb8623','admin','Admin User','EMP-0001','admin@fccs.com','+91 XXXXX XXXXX','$2b$12$dBIK1EfHTzr6YreZHgM82ewAmKVSctDjc6jH9s7Kdu.2LFA14Iina',NULL,NULL,1,'2026-03-20 05:39:46.711851','2026-03-20 05:39:46.711851');
CREATE UNIQUE INDEX ix_users_email ON users (email);
CREATE UNIQUE INDEX ix_users_username ON users (username);
CREATE UNIQUE INDEX ix_users_employee_id ON users (employee_id);
CREATE INDEX ix_exchange_rates_from_currency ON exchange_rates (from_currency);
CREATE INDEX ix_exchange_rates_to_currency ON exchange_rates (to_currency);
CREATE INDEX ix_exchange_rates_effective_date ON exchange_rates (effective_date);
CREATE INDEX ix_reports_name ON reports (name);
CREATE INDEX ix_reports_sequence_id ON reports (sequence_id);
CREATE INDEX ix_reports_is_deleted ON reports (is_deleted);
CREATE INDEX ix_reports_uploaded_at ON reports (uploaded_at);
COMMIT;
