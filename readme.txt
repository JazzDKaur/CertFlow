# 🎓 CertFLow

> **Automated Certificate Generation using Excel, PowerPoint, and Streamlit**

CertFlow is a Python-based Streamlit application that automates the creation of personalized certificates from an Excel spreadsheet and a PowerPoint template. It enables organizations, universities, training institutes, and event organizers to generate hundreds of certificates within minutes.

---

## ✨ Features

- 📊 Upload participant data from an Excel file
- 🎨 Use custom PowerPoint certificate templates
- 🔄 Dynamic placeholder replacement
- 📄 Generate editable PowerPoint certificates
- 📑 Convert certificates to PDF (Optional)
- 📦 Download all certificates as a ZIP file
- ⚡ Fast bulk certificate generation
- 🖥️ User-friendly Streamlit interface
- 🎯 Skip participants based on custom conditions (e.g., Grade F)

---

## 📷 Application Workflow

```text
Upload Excel File
        │
        ▼
Upload PowerPoint Template
        │
        ▼
Read Participant Data
        │
        ▼
Replace Placeholders
        │
        ▼
Generate PPT Certificates
        │
        ▼
(Optional) Convert to PDF
        │
        ▼
Create ZIP Archive
        │
        ▼
Download Certificates
```

---

# 📋 Input Excel Format

The Excel file should contain participant information.

Example:

| Name | Course | Grade | Date |
|------|---------|-------|------|
| John Doe | Python | A | 20 June 2026 |
| Jane Smith | AI | B | 20 June 2026 |

---

# 🎨 PowerPoint Template

Create placeholders inside your PowerPoint template.

Example

```
{{NAME}}

{{COURSE}}

{{DATE}}

{{GRADE}}
```

During execution these placeholders are automatically replaced with participant information.

---

# 📦 Output

The application generates

```
Generated Certificates/
│
├── PPT/
│   ├── John Doe.pptx
│   ├── Jane Smith.pptx
│
├── PDF/
│   ├── John Doe.pdf
│   ├── Jane Smith.pdf
│
└── Certificates.zip
```

---

# ⚙️ Technologies Used

- Python
- Streamlit
- Pandas
- OpenPyXL
- python-pptx
- LibreOffice (Optional for PDF conversion)
- ZipFile

---

# 💼 Use Cases

- Universities
- Schools
- Corporate Training
- Online Courses
- Workshops
- Conferences
- Webinars
- Employee Recognition
- Certification Programs

---

# 📌 Requirements

- Python 3.10+
- Streamlit
- Microsoft PowerPoint Template (.pptx)
- Excel (.xlsx)

Optional

- LibreOffice (for automatic PDF conversion)

---

# 🛠 Future Enhancements

- QR Code Verification
- Digital Signature Support
- Email Certificates Automatically
- Multiple Certificate Templates
- Certificate Preview
- Cloud Storage Integration
- Batch History
- Certificate Verification Portal
- Custom Fonts and Themes
- Database Integration

---

# 🤝 Contributing

Contributions are welcome.

1. Fork the repository
2. Create a new branch

```bash
git checkout -b feature-name
```

3. Commit your changes

```bash
git commit -m "Added new feature"
```

4. Push

```bash
git push origin feature-name
```

5. Open a Pull Request

---

# 👨‍💻 Author

Developed with ❤️ using Python and Streamlit.

If you found this project helpful, consider giving it a ⭐ on GitHub.