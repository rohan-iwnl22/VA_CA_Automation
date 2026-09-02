# Task 08: Deploy to Render

## Objective
Deploy the full application (FastAPI + frontend) to Render's free tier.

## Prerequisites
- GitHub repository with `UI-test` branch pushed
- Neon DB database created (free tier)
- All previous tasks completed and tested locally

## Environment Variables (Render Dashboard)

| Variable | Value | Notes |
|---|---|---|
| `DATABASE_URL` | `postgresql://user:pass@ep-xxx.us-east-2.aws.neon.tech/neondb?sslmode=require` | From Neon DB dashboard |
| `ADMIN_USERNAME` | `admin` | Default admin username |
| `ADMIN_PASSWORD` | `<secure_password>` | Change from default |
| `JWT_SECRET` | `<random_32_char_string>` | Generate with: `python -c "import secrets; print(secrets.token_urlsafe(32))"` |

## Render Configuration

### Option A: Web Service (Recommended)

1. Go to [render.com](https://render.com)
2. Click "New" → "Web Service"
3. Connect GitHub repo
4. Configure:
   - **Name:** `va-ca-automation`
   - **Region:** Oregon (or closest to users)
   - **Branch:** `UI-test`
   - **Runtime:** Python 3
   - **Build Command:**
     ```
     pip install -e .
     ```
   - **Start Command:**
     ```
     uvicorn va_ca_automation.api.main:app --host 0.0.0.0 --port $PORT
     ```
   - **Plan:** Free

5. Add environment variables in the "Environment" tab

6. Click "Create Web Service"

### Option B: Docker (Alternative)

If you prefer Docker:

1. Create `Dockerfile` in project root:
```dockerfile
FROM python:3.11-slim

WORKDIR /app
COPY . .
RUN pip install -e .

EXPOSE 8000
CMD ["uvicorn", "va_ca_automation.api.main:app", "--host", "0.0.0.0", "--port", "8000"]
```

2. In Render, choose "New" → "Web Service" → "Deploy an existing image" or use Docker runtime.

## Post-Deployment Steps

1. **Seed admin user:** On first run, `init_db()` creates the admin user with credentials from env vars.

2. **Test the endpoints:**
   - `GET /login` → Login page loads
   - `POST /api/login` → Returns JWT
   - `POST /api/merge-csv` → Merges files
   - `POST /api/report` → Generates reports
   - `POST /api/word` → Generates Word doc

3. **Update Word template (if needed):** Ensure `templates/Word file.docx` is in the repo.

## Free Tier Limitations

| Limitation | Impact | Mitigation |
|---|---|---|
| Spins down after 15 min idle | First request after idle takes ~30s to wake | Acceptable for internal tool |
| 750 hours/month | Enough for testing/low usage | Monitor usage |
| 512 MB RAM | May be tight for large file processing | Optimize memory usage |

## Files to Ensure Are in Repo

```
├── pyproject.toml              (with all new dependencies)
├── src/va_ca_automation/
│   ├── api/                    (all new API code)
│   ├── pipelines/              (existing)
│   ├── word_writer/            (existing)
│   ├── metadata/               (extended)
│   └── ...
├── static/                     (new frontend files)
├── templates/
│   ├── va_report_template.xlsx
│   ├── ca_report_template.xlsx
│   └── Word file.docx
├── config/
│   ├── business_rules.yaml
│   ├── column_mappings.yaml
│   └── naming_convention.yaml
├── tasks/                      (markdown task files)
└── Dockerfile                  (if using Docker)
```

## Acceptance Criteria
- [ ] Render service deployed successfully
- [ ] Neon DB connection working
- [ ] Login page accessible at `https://<service>.onrender.com/login`
- [ ] Admin can login with seeded credentials
- [ ] File upload and merge works
- [ ] Excel report generation works
- [ ] Word report generation works
- [ ] All files download correctly
- [ ] No errors in Render logs
