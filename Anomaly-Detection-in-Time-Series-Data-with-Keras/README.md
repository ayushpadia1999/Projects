This project was developed to create an anomaly detection model using deep learning.
Specifically, in this project I have designed and trained an LSTM autoencoder using the
Keras API with Tensorflow 2 as the backend to detect anomalies (sudden price
changes) in the S&P 500 index.
The Standard and Poor's 500, or simply the S&P 500, is a free-float weighted
measurement stock market index of 500 of the largest companies listed on stock
exchanges in the United States. Our DataSource is S&P 500 Daily Prices 1986 - 2018.
It is one of the most commonly followed equity indices.
Our Dataset has the following columns:
● Date : in format: yy-mm-dd
● Close: Closing price for the day.
The Dataset is a time series with daily un-adjusted closing prices for the S&P 500 Index.
The S&P 500 is a stock performance indicator for the Top 500 companies listed on the
stock exchange in the United States. It is considered as one of the best representations
of the United States stock market.
Source of Dataset: Yahoo Finance
The Project will follow the below steps:

● Import Libraries
● Load and Inspect the S&P 500 Index Data
● Data Preprocessing
● Temporalizing Data and Creating Training and Test Splits
● Build an LSTM Autoencoder
● Train the Autoencoder
● Plot Metrics and Evaluate the Model
● Detect Anomalies in the S&P 500 Index Data
● Visualize the Anomalies based on threshold

Anomaly detection has been widely studied in statistics and machine learning, where it
is also known as outlier detection, deviation detection, or novelty detection. The
burgeon of various successful outlier detection algorithms is now applied to detect
outliers in the stock market prices.
A direct application of outlier detection is to detect any market manipulation. Market
manipulation is a deliberate attempt to intervene in the market price in order to create
artificial, false, or misleading appearances with respect to the price of a security. Market
manipulation is detrimental because it distorts the prices and undermines the function of
the security market.This model can be used to detect such anomalies and
Manipulations in the stock markets saving investors money which might otherwise be
lost due to manipulators making illegal profits.
The required libraries:
● Numpy
● Tensorflow
● Keras
● Pandas
● Seaborn
● Matplotlib
● Plotly

I have also created interactive charts and plots using Plotly Python and Seaborn for
data visualization and displayed the results in Jupyter notebooks.
