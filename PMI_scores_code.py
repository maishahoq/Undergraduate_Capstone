# import gspread
# from nltk.collocations import BigramCollocationFinder, BigramAssocMeasures
# from nltk.tokenize import word_tokenize
# from nltk.corpus import stopwords
# import string
# import nltk
# nltk.download('punkt')
# nltk.download('stopwords')


# # import library

# # connect to the service account
# gc = gspread.service_account(filename="creds.json")

# # connect to your sheet (between "" = the name of your G Sheet, keep it short)
# sh = gc.open("PMI Input and Output").sheet1


# # Provided text from FinGPT Forecaster's output
# finGPT_output = """
# "[Positive Developments]:
# 1. The company's acquisition of Hess Corporation in an all-stock deal worth $53 billion is a significant positive development. This deal is expected to boost the company's production and reserves, providing a strong growth trajectory.
# 2. The company has been paying dividends and repurchasing shares, demonstrating its commitment to shareholder value.
# 3. Chevron has been successful in outperforming its peers in terms of total returns to shareholders, which suggests strong investor confidence in the company.

# [Potential Concerns]:
# 1. The company's stock price has been volatile in recent weeks, with a decrease from 168.92 to 155.87 in the period from 2023-10-18 to 2023-10-25.
# 2. The company's asset turnover ratio, which measures how efficiently assets are being used to generate sales, is only 0.787, indicating a potential inefficiency in asset utilization.
# 3. The company's long-term debt to total asset ratio is 0.0762, which is relatively high, indicating a potential financial risk.

# [Prediction & Analysis]:
# Prediction: Up by 3-4%
# Analysis: Given the recent developments and the potential concerns, we predict that Chevron's stock price will increase by 3-4% in the upcoming week. The acquisition of Hess Corporation is expected to boost the company's production and reserves, which can positively impact its stock price. Additionally, the company's commitment to shareholder value through dividends and share repurchases can also drive investor confidence and increase the stock price.

# However, the company's recent volatility in stock price, as well as its relatively high long-term debt to total asset ratio, are potential concerns that could weigh on the stock price. Nevertheless, these concerns are not expected to significantly impact the stock price in the short term. Therefore, we predict a positive stock price movement for Chevron in the coming week."
# """

# # Load stop words and punctuation
# stop_words = set(stopwords.words('english'))
# punctuation = set(string.punctuation)

# # Tokenize and filter out stop words and punctuation
# tokens = [word for word in word_tokenize(
#     finGPT_output) if word.lower() not in stop_words and word not in punctuation]

# bigram_measures = BigramAssocMeasures()
# finder = BigramCollocationFinder.from_words(tokens)

# # Filter for bigrams that might be relevant for financial/political analysis
# finder.apply_freq_filter(2)  # Adjust this number based on your dataset

# # Print bigrams with their PMI scores
# for bigram, score in finder.score_ngrams(bigram_measures.pmi):
#     print(bigram, score)


# Requirements to run this script->
# Make sure to use a different env than the base environment
# Make sure to install python3
# Make sure to pip3 install nltk into the new environment before running the script.
# Run the script with command -> python3 PMI_scores_code.py


import nltk
import string
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.collocations import BigramCollocationFinder, BigramAssocMeasures
from oauth2client.service_account import ServiceAccountCredentials
import gspread


nltk.download('punkt')
nltk.download('stopwords')

# Set up the connection to Google Sheets
scope = ["https://spreadsheets.google.com/feeds",
         "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)
sheet = client.open("PMI Input and Output").sheet1

# Load stop words and punctuation
stop_words = set(stopwords.words('english'))
punctuation = set(string.punctuation)

bigram_measures = BigramAssocMeasures()

# Loop through the first 40 rows in column B
for row in range(2, 41):  # Starting from 2 because the first row is often headers
    finGPT_output = sheet.cell(row, 2).value  # Column B

    # Tokenize and filter out stop words and punctuation
    tokens = [word for word in word_tokenize(
        finGPT_output) if word.lower() not in stop_words and word not in punctuation]

    finder = BigramCollocationFinder.from_words(tokens)
    finder.apply_freq_filter(2)  # Adjust as needed

    # Compile bigrams and their scores into a single string
    bigrams_output = "\n".join(
        [f"{bigram}: {score:.2f}" for bigram, score in finder.score_ngrams(bigram_measures.pmi)])

    # Write the results back into column C
    sheet.update_cell(row, 3, bigrams_output)  # Column C

print('PMI Analysis complete and results updated in the sheet.')
