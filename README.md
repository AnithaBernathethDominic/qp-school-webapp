# School Question Paper Topic Weightage Web App

A school-ready Flask web app where teachers can log in, drag-drop upload a question paper PDF and a syllabus PDF, and download Word/Excel reports with:

- question number
- parsed question text
- mapped topic/chapter/subtopic from the syllabus
- confidence and keyword evidence
- chapter-wise and subtopic-wise weightage
- dashboard charts and recent analysis history

## Default login

- Username: `admin`
- Password: `admin123`

Change these in `app.py` before using in production.

## Install and run

```bash
cd qp_school_webapp
pip install -r requirements.txt
python app.py
```

Open:

```text
http://127.0.0.1:5000
```

## How teachers use it

1. Login.
2. Go to Dashboard.
3. Drag and drop the question paper PDF and syllabus PDF.
4. Click **Analyze Paper**.
5. View question-wise mapping and topic weightage.
6. Download the Word or Excel report.

## Notes

The topic detector is transparent and rule-based. Improve accuracy by editing the `RULES` list in `analyzer.py`.
