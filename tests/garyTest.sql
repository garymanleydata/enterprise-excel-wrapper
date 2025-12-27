-- 1. Clear previous 'GaryTest' entries to avoid duplicates
DELETE FROM report_config 
WHERE profile_name = 'GaryTest';

-- 2. Insert Global Component Settings
-- These set the defaults for the entire workbook
INSERT INTO report_config (profile_name, component, setting_key, setting_value) 
VALUES 
    ('GaryTest', 'Global', 'primary_colour', '#4B0082'),   -- Indigo (Main Theme)
    ('GaryTest', 'Global', 'secondary_colour', '#E6E6FA'), -- Lavender (Accents)
    ('GaryTest', 'Global', 'title_prefix', 'GARY: '),
    ('GaryTest', 'Global', 'show_watermark', 'True'),
    ('GaryTest', 'Global', 'default_font_family', 'Arial'),
    ('GaryTest', 'Global', 'default_font_size', '10');

-- 3. Insert Header Component Settings
-- Controls the main page titles (fAddTitle)
INSERT INTO report_config (profile_name, component, setting_key, setting_value) 
VALUES 
    ('GaryTest', 'Header', 'font_size', '20'),
    ('GaryTest', 'Header', 'font_colour', '#4B0082'),
    ('GaryTest', 'Header', 'bg_colour', '#FFFFFF'); -- Transparent/White background

-- 4. Insert Logo Component Settings
-- Controls the logo placement and size (fAddLogo)
INSERT INTO report_config (profile_name, component, setting_key, setting_value) 
VALUES 
    ('GaryTest', 'Logo', 'path', 'assets/logo_gary.png'),
    ('GaryTest', 'Logo', 'width_scale', '0.6'),
    ('GaryTest', 'Logo', 'position', 'A1');

-- 5. Insert DataFrame Component Settings
-- Controls the main data tables (fWriteDataframe)
INSERT INTO report_config (profile_name, component, setting_key, setting_value) 
VALUES 
    ('GaryTest', 'DataFrame', 'header_bg_colour', '#4B0082'), -- Matches Global Primary
    ('GaryTest', 'DataFrame', 'header_font_colour', '#FFFFFF'),
    ('GaryTest', 'DataFrame', 'header_font_size', '11'),
    ('GaryTest', 'DataFrame', 'border_colour', '#CCCCCC'),    -- Grey grid lines
    ('GaryTest', 'DataFrame', 'stripe_rows', 'True'),         -- Enable zebra striping
    ('GaryTest', 'DataFrame', 'stripe_colour', '#F8F8FF');    -- Ghost White for stripes

-- 6. Insert KPI Component Settings
-- Controls the metric cards (fAddKpiRow)
INSERT INTO report_config (profile_name, component, setting_key, setting_value) 
VALUES 
    ('GaryTest', 'KPI', 'border_colour', '#4B0082'),
    ('GaryTest', 'KPI', 'label_font_size', '9'),
    ('GaryTest', 'KPI', 'value_font_size', '14'),
    ('GaryTest', 'KPI', 'value_font_colour', '#4B0082');

-- 7. Insert Data Dictionary Component Settings
-- Controls the appendix table (fAddDataDictionary)
INSERT INTO report_config (profile_name, component, setting_key, setting_value) 
VALUES 
    ('GaryTest', 'DataDict', 'header_bg_colour', '#000000'), -- Distinct Black Header
    ('GaryTest', 'DataDict', 'header_font_colour', '#FFFFFF');
	
COMMIT;

-- 1. Global: Hide Gridlines (2 = Hide on screen and print)
INSERT INTO report_config (profile_name, component, setting_key, setting_value)
VALUES ('GaryTest', 'Global', 'hide_gridlines', '2');

-- 2. Global: Default Date Format (UK Standard)
INSERT INTO report_config (profile_name, component, setting_key, setting_value)
VALUES ('GaryTest', 'Global', 'default_date_format', 'dd/mm/yyyy');

-- 1. Guidance Component (Standard Grey Background for definitions)
INSERT INTO report_config (profile_name, component, setting_key, setting_value)
VALUES ('GaryTest', 'Guidance', 'bg_colour', '#E8EDEE');

-- 2. Warning Component (Blue Background, White Text for Confidential Banners)
INSERT INTO report_config (profile_name, component, setting_key, setting_value)
VALUES 
    ('GaryTest', 'Warning', 'bg_colour', '#0091C9'),
    ('GaryTest', 'Warning', 'font_colour', '#FFFFFF');