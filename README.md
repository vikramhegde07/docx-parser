## 📄 DOCX Parser Microservice

This project is a lightweight **Flask-based microservice** that parses `.docx` files and extracts clean, structured HTML while preserving:

- ✅ Paragraphs with indentation and spacing
- ✅ Headings, bold/italic, and inline styles
- ✅ Tables and list structures
- ✅ Images (extracted and uploaded to S3-compatible storage)

It’s designed to be integrated with publishing platforms or content editors like Tiptap — enabling accurate conversion from Word documents to rich HTML without losing layout fidelity.

---

### ⚙️ Tech Stack
- Python 3.x
- Flask
- `python-docx` / `lxml` / `Pillow`
- Custom styling rules

---

### 🚀 Use Case
Use this service as a backend parser in:
- Article publishing workflows
- Content migration tools
- CMS importers for `.docx` files
