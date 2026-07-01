
# Livestock Feed Calculator

Arabic Streamlit application for building livestock feed recommendations, estimating feed quantities, and generating printable feeding-plan reports for farms.

## Project Scope

This project contains multiple working versions of the calculator and related utilities:

- `branded_feed_calulatr_V7.py`
  Main desktop-focused Streamlit version.
- `phone-display.py`
  Mobile-focused Streamlit version with responsive layout overrides for smaller screens.
- `Livestock_feed_calculator_pkg/app.py`
  Packaged app entry point.
- `Livestock_NRC_v2 _final.xlsx`
  Source nutrition and feed-bank workbook used by the calculator.
- `white_logo_nlfdp.png`
  Web app logo.
- `NLFDP.png`
  Printable report logo.

## Features

- Build feed recommendations for sheep, goats, and camels
- Support multiple animal groups in one farm plan
- Use market feed prices when available, with manual override support
- Estimate bag counts, cost per head, and whole-farm cost
- Generate printable feeding-plan reports
- Separate mobile-oriented app version for phone viewing

## Requirements

The packaged app requirements file is located at:

`Livestock_feed_calculator_pkg/requirements.txt`

Dependencies:

- `streamlit`
- `pandas`
- `numpy`
- `scipy`
- `openpyxl`

## Installation

From the project folder:

```bash
cd transit/Norah/Livstock_Feed_calulator
python -m venv .venv
source .venv/bin/activate
pip install -r Livestock_feed_calculator_pkg/requirements.txt
```

## Running The App

Desktop-oriented version:

```bash
streamlit run branded_feed_calulatr_V7.py
```

Phone-oriented version:

```bash
streamlit run phone-display.py
```

Packaged app version:

```bash
streamlit run Livestock_feed_calculator_pkg/app.py
```

## Data Files

The calculator depends on local project files, especially:

- `Livestock_NRC_v2 _final.xlsx`
- logo image files in the project root

If you move the app to a new machine or repository, keep these files in the same relative locations unless you also update the file paths in the scripts.

## Repository Structure



## Notes

- The app interface is primarily in Arabic.
- The printable report uses a different logo asset from the web app.
- Market-price extraction and comparison helper scripts are included in the repository, but they are not required to run the main Streamlit app.


```
# Animal_Livestock_And_Feed_calculator_ALAF
Smart livestock nutrition calculator designed to formulate balanced feed rations and optimize feed costs
>>>>>>> d5c8a77ddc3f580bcc67b04e3998e07f73f27f5d

