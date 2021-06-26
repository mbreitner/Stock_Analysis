# **Stock_Analysis**
Columbia Bootcamp Module 2 Stock Analysis

## **Overview of Project**
The project is to utilize our refactoring skills to make our initial VBA script run faster and future proof to look across multiple years (Sheets) in our workbook.

## **Results**
The results of the analysis show that the 12 tickers we have selected overall performed better in 2017 than 2018. 2017 saw an average yearly return of 67.3% compared to 2018 which saw an average yearly return of -8.5%. Ticker ENPH and ticker RUN were the only two tickers that saw a positive return in both years. Steve should recommend these two tickers to his parents. 

The refactored code ran in .226 seconds for 2017. ![VBA_Challenge_2017](https://user-images.githubusercontent.com/84791455/123522679-a3a4c600-d673-11eb-985b-a30e31fb97ac.PNG) The original script ran in 1.31 seconds for 2017 so the refactored code was faster. 

The refactored code also ran in .226 seconds for 2018. ![VBA_Challenge_2018](https://user-images.githubusercontent.com/84791455/123522686-b5866900-d673-11eb-8eee-359e9b11f686.PNG) The original script ran in 1.34 seconds for 2018 so our refactored code was again faster. 

## **Summary**
- **What are the advantages or disadvantages of refactoring code?
1. Advantages:
	 - Refactoring code helps to make the scripts run faster.
	 - It future proofs the sheet to make it possible to run with added data or other worksheets
	 - It makes it easier for other people to run your code
2. Disadvantages
	 - It can be time consuming to refactor if the code already works
	 - It can introduce bugs if you mess up during the refactoring

- **How do these pros and cons apply to refactoring the original VBA script?
	 - The refactored script ran faster than the original for both years
	 - the original code made it possible to run across different sheets much faster so if Steve wanted to look at 2019 as well he could easily add in the raw data and run the script
	 - The creation of the button makes it easy for anyone to come in and run this script
	 - It took me more time to create the refactored code and make it run properly. (Had to do some google-fu)
	 - I had issues with the creation of the three output arrays and had to debug a few times before the code ran successfully. 
