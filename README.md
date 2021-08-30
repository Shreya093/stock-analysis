# Stock-Analysis

## Overview of Project

In this project we will be helping Steve to analyse the stock information for the year 2017 and 2018 and determine whether or not the stocks are worth investing. 

### Purpose

The purpose of this project was to refactor the VBA code to collect the stock information for the year 2017 and 2018. We have initially performed our stock-analysis by writing the VBA code and now we need to refactor our code to improve the code readabililty and reduced complexity.
The data available contains stock information on 12 different stocks. The stock information contains a ticker value, the date the stock was issued, the opening, closing and adjusted closing price, the highest and lowest price, and the volume of the stock. The goal is to retrieve the ticker, the total daily volume, and the return on each stock.
We created macros to execute our analysis and an interface was created to allow the user to perform this analysis/functions with the click of a button.

## Results

We have the results of all the stocks analysis based on the "Total Daily Volume" and "Return" for both the years 2017 and 2018. The "Total Daily Volume" is the number of trades of stock in a given day and in the "Return" column, the stock's price at the end of the year is divided by the price at the beginning of the year and it's value is shown as the percentage of gain or loss in the growth.
The profit or gain value for a particular stock is depicted in green color and loss is depicted in red color.
 
Below are the stock analyis information for the year 2017 and 2018 :

<img width="260" alt="VBA_Challenge_StockAnalysis_2017" src="https://user-images.githubusercontent.com/88418201/131408982-21699e85-4b25-4081-a685-ba988a0b590f.png">

<img width="260" alt="VBA_Challenge_StockAnalysis_2018" src="https://user-images.githubusercontent.com/88418201/131408991-533d0bc8-1fe8-451c-9c48-4b404699f164.png">

1. We can see that in the year 2017 the ticker "TERP" shows a negative stock return value of -7.2% and rest all the other tickers had a good return value. 
2. Ticker "SEDG" had the highest return value of around 185% in 2017. 
3. While in the year 2018 all the ticker's had a loss in their stock growth except "ENPH" (81.9%) and "RUN" (84%). Ticker "SEDG" which had the highest stock return value in 2017 suffered a loss in return value of -7.8% in 2018.

Before refactoring our stock-analysis code we carried out our analysis with the help of nested loops - an iterative process within which multiple additional iterative processes are contained. Below is a code snippet of using nested loops :

<img width="500" alt="Screen Shot 2021-08-30 at 2 51 22 PM" src="https://user-images.githubusercontent.com/88418201/131410629-e27cb320-a296-4952-a89a-c7368ac331c5.png">

While nested loops served our purpose to perform the analysis, one of the drawbacks is that they caused our code to take a longer run time.

<img width="244" alt="OriginalScript_ElapsedTimeRun_2017" src="https://user-images.githubusercontent.com/88418201/131411026-c01a1e94-1087-40c1-9e7d-50f476a674c2.png">

<img width="244" alt="OriginalScript_ElapsedTimeRun_2018" src="https://user-images.githubusercontent.com/88418201/131411031-bb3655f8-f0cd-4c9a-bcb0-74af5f064a26.png">

By refactoring our code we can overcome this drawback by using arrays. Arrays helps us to find the data with the help of index. Below is a code snippet of using arrays in our refactored code:

<img width="500" alt="Screen Shot 2021-08-30 at 3 01 34 PM" src="https://user-images.githubusercontent.com/88418201/131411645-c9921f04-b262-4620-a3d2-3b98f6347c75.png">

By using arrays in our code we can increase our code efficacy and hence it will run our code in much lesser time.

<img width="244" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/88418201/131411369-77e670b8-9c32-4ca3-b391-e9e05fa5f956.png">

<img width="244" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/88418201/131411389-b20346dc-0251-41f4-bd15-9b59d6e6d327.png">

## Summary 

### 1. What are the advantages or disadvantages of refactoring code?

Advantages of refactoring the code:
1. The main goal of code refactoring is to make code more maintainable and extensibility. It helps the user to access the code efficiently even if there are any upgrades or future updates into the system.
2. It also makes the code more readable. Any code which is easily understandable makes half of the coder's work easier. The code can also be easily understood by different users.
3. Another potential purpose of code refactoring is performance improvement. So refactoring may enable an application to perform faster or use fewer server capacities.
4. Refactoring of code can also advantage us in cost saving. Refactoring contributes to the occurrence of reusable design elements that may be simply used for new features in the future. Moreover, code refactoring favors bugs' prevention which implies time and costs saving.

Disadvantages of refactoring the code:
1. There can be applications that are too large to not having the proper test cases for the existing codes, which may ultimately pose some risk if we try to refactor our code.
2. It can cost more development time and may not be safe. It is more likely that while refactoring the code we may get caught up into new bugs .
3. It is also risky when the existing code doesn't have proper test cases due to which any different user while working on the refactored code can mess it up.

### 2. How do these pros and cons apply to refactoring the original VBA script?

Refactoring the original VBA script made our code more readable and efficient. In our original VBA script we used nested loops to perform the stock analysis for 2017 and 2018 which made our code a lot complex and it also increased the run time of our code, which in coding world is not a good approach. The reduced run time can be a major drawback while handling large amount of data in organizations.
I personally found refactoring the code much efficient and easily understandable. The use of arrays helped to pick the data from our stock analysis dataset based on the Index value. Fetching data in arrays using index is a very efficient approach and makes our code run much faster.
