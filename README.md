# Event Insights – Sentiment & Data Report Generator

This project analyzes feedback from volunteers and donators to generate a detailed sentiment and participation report using Natural Language Processing (NLP) and data visualization techniques.

## Features

- Sentiment analysis using the `nlptown/bert-base-multilingual-uncased-sentiment` model
- Separate insights for volunteers and donators
- Age distribution, participation, and sentiment graphs
- Summary tables for motivations, positive comments, and suggested improvements
- Automatic Excel report generation with embedded graphs
- Analysis timestamp and duration logging

## Technologies Used

- **Python**
- **Pandas** for data processing
- **Transformers (Hugging Face)** for sentiment analysis
- **Matplotlib & Seaborn** for visualizations
- **Openpyxl** for Excel report creation
- **Counter** for textual frequency analysis

## Input

CSV file with feedback including:
- Role (Volunteer/Donator)
- Age
- Motivations, experiences, and improvement suggestions

## Output

An Excel file named `EventInsights_<timestamp>.xlsx` containing:
- Two sheets for volunteer and donator insights
- A summary sheet with graphs and analysis metadata
- Charts and frequency tables

## Purpose

The tool was designed to support decision-making in community events by understanding participant motivations, experiences, and suggestions through AI-powered analysis.
