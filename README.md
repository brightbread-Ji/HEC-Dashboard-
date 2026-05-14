# HEC Team Financial Dashboard

Lightweight Streamlit dashboard for HEC team financial and project previews.

## Quick Start

```powershell
pip install -r requirements.txt
Copy-Item config/auth.example.json config/auth.json
Copy-Item data/team_targets.example.json data/team_targets.json
streamlit run app.py
```

Place the monthly Excel source file in the project root as:

```text
HEC AOT by team.xlsx
```

Optional local files:

- `ipsos-logo.png.png` for the sidebar/login logo.
- `config/auth.json` for real BP/team accounts.
- `data/team_targets.json` for local Team Target AOT values.
- `data/uploads/` for BP-uploaded monthly Excel files.

These local data/config files are intentionally ignored by Git because they may contain company data, passwords, or runtime-only values.

## App Pages

- `财务信息预览`: AOT, AOGM%, target completion, monthly movement, client and product mix.
- `项目信息预览`: TO metrics, selected project count, searchable project list, and control/compliance checks.
- `设置`: BP-only monthly upload and Team Target AOT maintenance.
