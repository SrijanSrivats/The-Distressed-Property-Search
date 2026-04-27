# Distressed Property Search (Streamlit App)

## Project Overview

The Distressed Property Search is an advanced interactive analytics application designed to identify potentially undervalued rental markets across the United States using HUD Fair Market Rent (FMR) data.

This project focuses on combining data engineering, statistical analysis, geospatial visualization, and AI-generated insights to support real-world decision-making in real estate investment. The application enables users to explore patterns, detect anomalies, and evaluate investment opportunities using structured and explainable metrics.

---

## Live Application

👉 **Streamlit App:**  
[PASTE YOUR DEPLOYED STREAMLIT LINK HERE]

---

## GitHub Repository

👉 **Repository:**  
https://github.com/YOUR-USERNAME/The-Distressed-Property-Search

---

## Key Features

### 1. Outlier Detection
- Identifies counties where rental prices are significantly lower than comparable regions
- Uses configurable thresholds (e.g., 25% below median)
- Helps detect undervalued markets

### 2. Choropleth Heatmap Visualization
- County-level interactive U.S. map
- High-contrast color scaling for clear interpretation
- Quickly identifies regional rent disparities

### 3. Neighbor-Based Comparison
- Compares counties against nearby regions
- Detects localized anomalies
- Improves contextual accuracy of “distress” signals

### 4. Multi-Year Trend Analysis
- Uses data from FY2021–FY2026
- Tracks rent growth and recovery signals
- Distinguishes between:
  - consistently low-performing markets
  - recovering investment opportunities

### 5. Data Quality & Anomaly Signals
- Detects irregularities in rent patterns
- Identifies potential inconsistencies in data trends
- Supports analytical decision-making

### 6. AI-Generated Investment Insights
- Uses Gemini API
- Generates:
  - summary
  - strengths
  - risks
  - recommendation
- Structured, data-driven explanations

---

## Dataset

The project uses Fair Market Rent (FMR) datasets from:

- HUD Office of Policy Development and Research

Required files:

data/
- FY26_FMRs.xlsx
- FY2021_FMRs.xlsx
- FY2022_FMRs.xlsx
- FY2023_FMRs.xlsx
- FY2024_FMRs.xlsx
- FY2025_FMRs.xlsx

---

## Project Structure

The-Distressed-Property-Search/
│
├── app.py                  # Main Streamlit application
├── requirements.txt       # Dependencies for deployment
├── pyproject.toml         # uv-based project configuration
├── README.md              # Project documentation
│
└── data/                  # Input datasets
    ├── FY26_FMRs.xlsx
    ├── FY2021_FMRs.xlsx
    ├── FY2022_FMRs.xlsx
    ├── FY2023_FMRs.xlsx
    ├── FY2024_FMRs.xlsx
    └── FY2025_FMRs.xlsx

---

## Technology Stack

### Core
- Python
- Streamlit

### Data Processing
- Pandas
- NumPy

### Visualization
- Plotly

### Statistical Analysis
- SciPy

### Geospatial
- Shapely

### AI Integration
- Google Gemini (google-genai)

---

## Local Setup Instructions (uv-based)

### Step 1: Clone the Repository

git clone https://github.com/YOUR-USERNAME/The-Distressed-Property-Search.git  
cd The-Distressed-Property-Search

---

### Step 2: Install uv

pip install uv

---

### Step 3: Install Dependencies

uv sync

---

### Step 4: Configure Secrets

Create folder:

.streamlit/

Create file:

.streamlit/secrets.toml

Add:

ai_backend = "gemini"  
gemini_api_key = "YOUR_API_KEY"  
gemini_model = "gemini-2.5-flash"

---

### Step 5: Run Application

uv run streamlit run app.py

Open in browser:

http://localhost:8501

---

## Deployment Instructions (Streamlit Cloud)

1. Go to https://share.streamlit.io  
2. Login with GitHub  
3. Select your repository  
4. Set main file:

app.py

5. Add Secrets:

ai_backend = "gemini"  
gemini_api_key = "YOUR_API_KEY"  
gemini_model = "gemini-2.5-flash"

6. Click Deploy  

---

## Important Notes

- Ensure all Excel files are inside the data/ folder
- Do NOT upload secrets.toml to GitHub
- File names must match exactly
- Gemini API key is required for AI functionality

---

## Common Issues & Fixes

### App not loading data
- Check if data/ folder exists
- Verify all Excel files are present

### AI not working
- Ensure API key is set in Streamlit secrets

### Deployment failure
- Check logs in Streamlit Cloud
- Verify requirements.txt dependencies

---

## Team

- Aditya Pola  
- Srijan Srivatsava Ganji  
- Raviteja Thode  

---

## Course

Visual Analytics  
Spring 2025

---

## Objective

This project demonstrates how combining data visualization, statistical analysis, geospatial intelligence, and AI-driven insights can support real-world decision-making in real estate investment.

