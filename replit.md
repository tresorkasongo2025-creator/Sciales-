# E-SCIALES UNILU + Application

## Overview
This web application, "E-SCIALES UNILU +", is designed for the Faculty of Social, Political, and Administrative Sciences at the University of Lubumbashi (UNILU). Its primary purpose is to manage student registrations and control attendance efficiently. The project aims to streamline administrative tasks, provide robust student management features, and offer a user-friendly interface for both students and faculty. Key capabilities include student enrollment with automatic QR code generation, secure administrative dashboards, real-time attendance tracking via QR code scanning, course and timetable management, news publication, and a unique hourly charge letter generator for professors. The application uses "CODE-ID FAC" universally in place of "Matricule" and visually incorporates a red medical cross (⊕) as part of its logo.

## User Preferences
- The application interface should be 100% in French.
- Use "CODE-ID FAC" terminology instead of "Matricule" throughout the interface.
- Ensure the application name is presented as "E-SCIALES UNILU +" with a red medical cross (⊕) symbol.
- QR codes should contain all student/professor information in JSON format.
- Matricules (CODE-ID FAC) should be generated automatically and sequentially (e.g., ETU0001, ETU0002).
- The base database creation should be automatic upon application startup.
- All dates and times should use local time.
- The 5 most recent news articles should be displayed on the homepage.
- The public attendance scanner should be accessible directly from the homepage.

## System Architecture

### UI/UX Decisions
The application features a professional and responsive design, optimized for mobile viewing, particularly on the homepage with compact texts. Sections are clearly delineated with colored cards for departments and a dedicated "À PROPOS" page with colored sections. The navigation includes "À PROPOS" between "Horaires" and "DECANAT", and a "Professeur" menu link.

### Technical Implementations
- **Student Registration:** Full form with photo upload, unique 4-digit "CODE-ID FAC" generation (format: ETUXXXX), personalized QR code creation (PNG), and PostgreSQL storage.
- **DECANAT (Administration) Space:** Password-protected access (default: `DECANAT2026`), student list viewing (by department/promotion), PDF/Excel export, and attendance system management.
- **Student Search:** Homepage search by "CODE-ID FAC", displaying full student info, photo, and QR code (downloadable).
- **Attendance Control:** QR code scanning via camera/tablet (30 FPS optimized), automatic entry time recording, support for students and professors, real-time attendance lists.
- **Course Management (DECANAT):** Manual/bulk course addition (code, name, department, promotion), course deletion, and integration with the attendance scanner.
- **Timetable Management (DECANAT):** Publication of normal/supplementary timetables via PDF upload, organized by department/promotion, deletion with file cleanup, and public consultation with filters.
- **News and Announcements (DECANAT):** Publication of official news/communiqués with optional images, display of 5 most recent on homepage, deletion with image cleanup.
- **Public Attendance Scanner:** Accessible from homepage without authentication, mandatory course selection, 1-hour anti-multiple scan protection with countdown, automatic presence recording.
- **"À PROPOS" Page:** Comprehensive faculty information (mission, structure, personnel, contacts).
- **"DÉPARTEMENTS" Page:** Detailed presentation of 4 departments (Sociology, Anthropology, International Relations, Political and Administrative Sciences) with training info and career opportunities.
- **Professor Menu:** Password-protected access (default: `PROF2026`), dedicated login, presence list management, Excel/PDF export of lists with date range filters.
- **Page Management (DECANAT):** Editing "À PROPOS" and "DÉPARTEMENTS" content (texts, facade photo, contacts), using JSON for database storage.
- **Hourly Charge Letter Generator (DECANAT):** Upload Excel file (titulaire, courses, promotion, hours, credits) and Word template, automatic generation of personalized Word letters (.docx) per titulaire with all their courses grouped on a single letter, formatted table insertion (5 columns: N°, Intitulé, Promotion, Heures, Crédits), total hours/credits calculation, automatic replacement of placeholder text ("………………………" or "LUBUYA DIAMBILA") with titulaire name, removal of existing table and replacement with new styled table at same position, and ZIP/individual download. Only courses with a titulaire are processed. Files named as: Lettre_TITULAIRE.docx.

### System Design Choices
- **Security:** Flask session for authentication, file upload size limit (16MB), restricted image (PNG, JPG, JPEG) and PDF formats, secured filenames (werkzeug.secure_filename), protected DECANAT routes, public scanner with anti-multiple scan protection (1-hour delay), HTTPS recommended for production.
- **File Storage:** Photos in `static/uploads/`, QR codes in `static/qrcodes/`, timetables in `static/horaires/`, news images in `static/actualites/`.
- **Project Structure:** Clear separation of `app.py` (main Flask app), `templates/` (HTML files), and `static/` (CSS, uploads, qrcodes, horaires, actualites).

## External Dependencies

### Backend
-   **Flask:** Python web framework.
-   **Flask-SQLAlchemy:** ORM for database interaction.
-   **PostgreSQL:** Relational database management system (provided by Replit).
-   **Werkzeug:** Utility library for handling file uploads.

### Frontend
-   **HTML5/CSS3:** For structuring and styling web pages.
-   **JavaScript:** For interactive elements.
-   **html5-qrcode:** Library for QR code scanning functionality.

### Document Generation & Processing
-   **qrcode:** Python library for generating QR codes.
-   **Pillow:** Python imaging library for image processing.
-   **openpyxl:** Library for reading/writing Excel 2010 xlsx/xlsm/xltx/xltm files.
-   **reportlab:** Library for generating PDF documents.
-   **python-docx:** Library for creating and updating Microsoft Word .docx files.