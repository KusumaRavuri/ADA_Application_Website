# ADA Application Portal — Setup Guide

## Quick Start

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the App
```bash
python app.py
```

### 3. Open in Browser
```
http://127.0.0.1:5000/
```

---

## Project Structure

```
ada_portal/
├── app.py                  # Flask backend (all routes + Excel logic)
├── requirements.txt        # Python dependencies
├── APPLICATION.xlsx        # Auto-created on first run
│
├── templates/
│   ├── base.html           # Base layout with nav + glassmorphism bg
│   ├── index.html          # Landing/home page
│   ├── apply.html          # 9-step application form
│   ├── success.html        # Submission success page
│   └── admin.html          # Admin dashboard
│
├── static/
│   ├── css/style.css       # Full premium glassmorphism CSS
│   └── js/apply.js         # Multi-step wizard JS logic
│
└── uploads/
    ├── pdfs/               # Uploaded undertaking PDFs stored here
    ├── photos/             # Uploaded passport photos stored here
    └── samples/            # Place undertaking_sample.pdf here
```

---

## Features

### Applicant Portal
- 9-step wizard with animated progress bar
- Step-by-step validation (required fields checked before proceeding)
- Drag-and-drop file upload with real-time preview
- Image preview for passport photo
- Scrollable undertaking declaration with 4 checkboxes
- Full review page before final submit
- Duplicate email detection
- File type validation (PDF-only, JPG-only)

### Admin Dashboard (`/admin`)
- Total applications count card
- Search by applicant name
- Filter by branch
- Full sortable table view
- Direct links to view uploaded PDF and photo
- Download APPLICATION.xlsx button

### Excel Storage
- `APPLICATION.xlsx` auto-created with formatted headers
- **Applications** sheet — full details (44 columns)
- **Admin Summary** sheet — 10-column summary
- Alternating row colors, auto-column width

---

## Adding the Undertaking Sample

Place your `undertaking_sample.pdf` file in:
```
uploads/samples/undertaking_sample.pdf
```

Then update the download link in `templates/apply.html` (Step 7):
```html
<a href="/uploads/samples/undertaking_sample.pdf" ...>
```

---

## Customization

- **Organization name / branding**: Edit `base.html` nav brand text
- **Form fields**: Add/remove inputs in `apply.html` and update `app.py`'s `HEADERS_FULL` list and `save_to_excel()` function
- **Colors**: Edit CSS variables at the top of `static/css/style.css`

---

## Tech Stack

| Component | Technology |
|-----------|-----------|
| Backend   | Python Flask 3.0 |
| Frontend  | Bootstrap 5.3 + Custom CSS |
| Storage   | openpyxl (Excel) |
| Uploads   | Werkzeug secure_filename |
| Fonts     | Plus Jakarta Sans + Syne (Google Fonts) |
