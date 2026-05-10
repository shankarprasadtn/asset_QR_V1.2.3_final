# Asset QR Management System – v1.2.3

A Flask-based Asset QR Management System designed to manage IT assets with QR code generation, asset tracking, hotspare management, Excel import/export, and user authentication.

## Features

- Asset inventory management
- QR code generation for assets
- Asset search and filtering
- Hotspare asset tracking
- Excel import and export support
- User authentication and role management
- Dashboard for quick asset overview
- Flask-based lightweight web application
- SQLite database support
- Responsive UI templates

---

## Project Structure

```bash
asset_QR_V1.2.3_final/
│
├── app.py
├── models.py
├── import_from_excel.py
├── requirements.txt
├── users.db
├── static/
├── templates/
│   ├── index.html
│   ├── hotspare.html
│   └── ...
└── uploads/
```

---

## Prerequisites

Ensure the following are installed:

- Python 3.10+
- pip
- Git

---

## Installation

### 1. Clone the Repository

```bash
git clone https://github.com/shankarprasadtn/asset_QR_V1.2.3_final.git
cd asset_QR_V1.2.3_final
```

### 2. Create Virtual Environment

#### Windows

```bash
python -m venv venv
venv\Scripts\activate
```

#### Linux/macOS

```bash
python3 -m venv venv
source venv/bin/activate
```

---

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

---

## Running the Application

```bash
python app.py
```

Application will start on:

```bash
http://127.0.0.1:5000
```

---

## Default Login Credentials

```text
Username: admin
Password: Admin@123
```

> Change the default password after first login.

---

## Import Assets from Excel

Use the Excel import utility:

```bash
python import_from_excel.py
```

Supported features:

- Bulk asset upload
- Excel-based inventory update
- Platform-wise asset categorization

---

## Export Features

The application supports exporting:

- Full asset inventory
- Filtered asset reports
- Hotspare asset reports
- Platform-specific reports

---

## Tech Stack

- Backend: Flask
- Frontend: HTML, CSS, Bootstrap
- Database: SQLite
- QR Generation: Python QR libraries
- Excel Handling: Pandas / OpenPyXL

---

## Screenshots

Add screenshots inside a `screenshots/` folder and update this section.

Example:

```markdown
![Dashboard](screenshots/dashboard.png)
```

---

## Security Notes

- Change default admin credentials
- Restrict database file access
- Use HTTPS in production
- Regularly back up the database

---

## Future Enhancements

- LDAP/AD Integration
- Email notifications
- Asset lifecycle tracking
- REST API support
- Docker deployment
- Multi-user audit logs

---

## License

This project is intended for internal asset management and learning purposes.

---

## Author

Shankar Prasad T N

GitHub:
https://github.com/shankarprasadtn
