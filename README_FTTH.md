# FTTH Evidence Management System

A web-based system for managing **Fiber To The Home (FTTH)** field documentation and evidence, developed during internship at **PT Telkom Akses**.

---

## Overview

This system automates the documentation process for FTTH installation field work, replacing manual reporting with a structured digital workflow. Field technicians can upload photo evidence, and the system automatically generates standardized reports.

---

## Features

- Automated document generation (`.docx`) using template-based approach
- Image upload system for field evidence validation
- Structured data management for FTTH documentation
- Stateless architecture for efficient request handling
- Cloud deployment via Railway

---

## Tech Stack

| Component | Technology |
|---|---|
| Backend | Python, Flask |
| Document Generation | python-docx (template-based) |
| Database | MySQL |
| Deployment | Railway |
| Frontend | HTML, CSS, JavaScript |

---

## Project Structure

```
ftth-evidence/
├── templates/          # HTML frontend templates
├── app.py              # Main Flask application
├── requirements.txt    # Python dependencies
├── Procfile            # Railway deployment config
└── railway.toml        # Railway configuration
```

---

## Getting Started

### Prerequisites
- Python 3.10+
- MySQL

### Installation

```bash
# Clone the repository
git clone https://github.com/zahrotulcm-cloud/Evidence.git
cd Evidence

# Create virtual environment
python -m venv venv
venv\Scripts\activate  # Windows

# Install dependencies
pip install -r requirements.txt

# Setup environment variables
cp .env.example .env
# Edit .env with your database configuration

# Run the application
python app.py
```

---

## Deployment

This project is deployed using **Railway** cloud platform.

Configuration files:
- `Procfile` — defines the web process
- `railway.toml` — Railway-specific settings

---

## Author

**Zahrotul Camelia**
- zahrotulcm@gmail.com
- GitHub: [@zahrotulcm-cloud](https://github.com/zahrotulcm-cloud)

---

## License

This project was developed as an internship project at **PT Telkom Akses** through the **State Polytechnic of Malang** internship program.
