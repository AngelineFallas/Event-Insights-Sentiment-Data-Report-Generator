# 📊 EventInsights: Sentiment-Based Feedback Analyzer (Prototype)

This project is a **prototype** designed to analyze textual feedback using Natural Language Processing (NLP) and generate a structured Excel report. It groups comments by sentiment and user type (Volunteer or Donator), helping organizations understand motivations, positive highlights, and areas for improvement.

## 📥 Data Collection  
The data for this analysis was collected using **Google Forms**, and the responses were exported as a **CSV file**. This allowed for easy processing and analysis in Python using standard data science libraries.

## 🛆 Requirements & Installation

Before running the script, install the required libraries using:

```bash
pip install pandas matplotlib seaborn openpyxl textblob nltk
```

You also need to download the `vader_lexicon` for TextBlob's sentiment analysis to work correctly:

```python
import nltk
nltk.download('vader_lexicon')
```

## 🧠 Why These Libraries?

* **Pandas**: Efficient data manipulation and Excel exportation.
* **TextBlob + NLTK**: Lightweight NLP tools suitable for basic sentiment classification.
* **Matplotlib & Seaborn**: For generating clear and professional visualizations.
* **OpenPyXL**: To manage multi-sheet Excel reports with formatting.

These were chosen for their simplicity and ease of integration in a lightweight prototype. More advanced models (e.g., from Hugging Face) may be used for production-grade implementations.

## ⚙️ How It Works

1. The script reads a `.csv` file with user feedback including fields like role (Volunteer/Donator), age, and comment.
2. It performs sentiment analysis using TextBlob and classifies each comment as *Positive*, *Neutral*, or *Negative*.
3. It generates an Excel file with the following sheets:

   * **Execution Log**: Contains timestamp, number of comments processed, and runtime duration.
   * **Analysis Summary**: A combined overview, including total feedback and role distribution.
   * **Volunteers Report**: Insights specific to volunteers.
   * **Donators Report**: Insights specific to donators.

Each detailed report includes:

* Bar plots of age distribution.
* Comment counts per sentiment.
* Extracted summaries: *Motivations*, *Highlights*, and *Suggestions for Improvement*.

The Excel filename is auto-generated as:
`EventInsights_YYYY-MM-DD_HH-MM-SS.xlsx`

## 📝 Output Files in This Repository

* `FeelingAnalizer.py`: Main Python script.
* `volunteering_data(2).csv`: Sample dataset used in this analysis.
* `EventInsights_2025-04-03_19-36-52.xlsx`: Final generated report from the sample data.

## 💼 Application Potential

Although developed for a volunteer-donator feedback context, this prototype can be adapted to **customer service** or **product review analysis**. To do so:

* Replace the `volunteering_data(2).csv` with your own feedback dataset.
* Adjust role filters in the code:
  (Change `"Volunteer"`/`"Donator"` labels as needed.)
* Optionally customize the summary sections and visualizations for your context.

## 🚀 Future Improvements (if scaled)

* Use multilingual or transformer-based NLP models (e.g., BERT).
* Integrate with a front-end or dashboard for dynamic interaction.
* Automate email report delivery or real-time monitoring.

##
## 👤 About the Author
I'm a creative individual who believes in using technology to support meaningful causes. This repository is part of my personal portfolio, as someone motivated by the potential of AI to improve the world we live in.
