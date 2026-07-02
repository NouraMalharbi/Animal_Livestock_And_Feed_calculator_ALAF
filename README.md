# Alaf
## علف 
 حاسبة رقمية لدعم قرارات تغذية الماشية من خلال توصيات مبنية على الاحتياج الغذائي، و تقدير 
 الكميات والتكاليف وإعداد تقارير التغذية
.

احدى حلول البرنامج الوطني لتطوير قطاع الثروة الحيوانية والسمكية



Arabic Streamlit application for building livestock feed recommendations, estimating feed quantities, and generating printable feeding-plan reports for farms.

## Project Scope

This repository contains the standalone `alaf` app in two UI variants:

- `app_desktop.py`
  Main desktop-oriented Streamlit version.
- `app_phone.py`
  Mobile-oriented Streamlit version with phone-specific layout adjustments.
- `data/Livestock_NRC_v2.xlsx`
  Nutrition and feed-bank workbook used by the calculator.
- `assets/white_logo_nlfdp.png`
  Web app logo.
- `assets/NLFDP_full.png`
  Printable report logo.
- `requirements.txt`
  Python dependencies for both app variants.

## Features

- Build feed recommendations for sheep, goats, and camels
- Support multiple animal groups in one farm plan
- Use market feed prices when available, with manual override support
- Estimate bag counts, cost per head, and whole-farm cost
- Generate printable feeding-plan reports
- Provide separate desktop and phone app variants

## Requirements

Dependencies listed in `requirements.txt`:

- `streamlit`
- `pandas`
- `numpy`
- `scipy`
- `openpyxl`

## Installation

From the repository root:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Running The App

Desktop version:

```bash
streamlit run app_desktop.py
```

Phone version:

```bash
streamlit run app_phone.py
```

## Repository Structure

```text
alaf/
├── app_desktop.py
├── app_phone.py
├── requirements.txt
├── README.md
├── assets/
│   ├── NLFDP_full.png
│   └── white_logo_nlfdp.png
└── data/
    └── Livestock_NRC_v2.xlsx
```

## Data And Assets

The app depends on local relative paths inside this repository:

- `data/Livestock_NRC_v2.xlsx`
- `assets/white_logo_nlfdp.png`
- `assets/NLFDP_full.png`

If you move files, update the paths in both app scripts.

## Notes

- The UI is primarily in Arabic.
- The desktop and phone apps share the same calculation logic, with different layout tuning.
- The printable report uses a different logo asset from the web app.
