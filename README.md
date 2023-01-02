# stock-analysis

## Overview of Project:
Stock analysis is an important task for anyone interested in participating in the stock market. Though many employees do just this as a living, this is by all means accessible to the everyday interested party. There are many factors that can affect the day-to-day ebbs and flows of a stock's performance, but a single snapshot in time is too risky to rely upon for those not engaged in day trading. For instance, when stay at home orders were put in place at the start of the Covid pandemic, stock prices plummeted across the board. Someone who had invested heavily in several key companies at the lowest lows would have increased their net worth many-fold by now. However, the lowest lows didn't last long before a number of companies rebounded. In fact, some increased to their highest highs during the rebound (some have mellowed out more than others in overall performance fluctuation). Another example is the intense impact the subreddit "wallstreetbets" had on GameStop values for a limited amount of time. The risky nature of timing the market led to many people taking losses by investing in GameStop at its height prior to plummeting back down - though, it is still much higher than it had been worth prior to wallstreetbets propping it up. Wars are known to have an impact on the economy, and that has been seen to play out since the War in Ukraine began. There are many more examples that could be given to explain how individual stocks might fluctuate in price as well as the stock market in general, but these are recent examples that the public has seen widely reported in the news.

With the above in mind, it is important to know what the buyer's risk level and general purpose for buying is. Day trading is focused on making a profit off of the day to day stock market fluctuations - it can lead to massive payouts, but those returns come at a high risk level for ending with large losses. This is a great method for some people, particularly those who are still years away from retirement and have the finances to handle some losses. Those who invest for retirement are typically looking for low risks and steady returns while foregoing the potential of large paydays in the present. 

The particular stocks analyzed in this dataset are all for green energy companies. Many investors choose the type of stocks they invest in through some connection to their personal beliefs. The couple that this analysis was conducted for happens to be passionate about green energy. They had originally wanted to invest all of their money into DAQO (DQ) New Energy Corporation, but as can be seen in the results below, DAQO may not be the best stock choice for them to put all their money into. 

For this analysis the data had been received in an Excel file and the data was then analyzed using VBA.

## Results:
### Below are screenshots for the 2017 and 2018 output data tables:
### 2017
![VBA_Challenge_2017_Output](https://user-images.githubusercontent.com/101941048/210194129-cbc61dd3-ab3a-4a77-8716-c2c414a070fc.png)

### 2018
![VBA_Challenge_2018_Output](https://user-images.githubusercontent.com/101941048/210194135-5087bea9-69e2-4fb5-924e-2f1500dbb3e3.png)

### Comparison of Years:
The 2017 analysis shows that across the board stocks performed very well, with the exception of TERP. It is understandable that the couple originally wanted to invest all of their money into the DQ stock considering it had a 199.4% return rate in 2017! However, the 2018 return ended up being -62.6%. And when looking at other stocks, ENPH and RUN were the only stocks with positive return rates for 2018.

### Code Comparison:
The original code created was refactored to speed up the run time. Below can be seen the original run times against refactored times for both 2017 and 2018.

### Original Code: 2017
![Green_Stocks_2017_Timer](https://user-images.githubusercontent.com/101941048/210194473-8ac00373-5028-469d-9905-7186cc17a41a.png)

### Refactored Code: 2017
![VBA_Challenge_2017](https://user-images.githubusercontent.com/101941048/210194498-5a7a9ef0-8624-4250-9dff-a2ae6be92b8f.png)

### Original Code: 2018 
![Green_Stocks_2018_Timer](https://user-images.githubusercontent.com/101941048/210194492-a5a1be97-abea-4b5a-8856-d89a49fcb685.png)

### Refactored Code: 2018
![VBA_Challenge_2018](https://user-images.githubusercontent.com/101941048/210194497-1275a0ea-31c4-46cc-9112-2dbc5ce36642.png)

## Summary:
Not all code is made the same - some code choices can take much longer to run than others. Refactoring code is important to improve the speed (as well as general appearance) of code being used. The dataset analyzed here is relatively small, so the runtime isn't noticeably different for the user. However, when comparing the runtime results, we can see that the refactored code is 10x faster than the original code used for both 2017 and 2018! That is a huge difference and can save a lot of time should this dataset be significantly expanded in the future.

To conclude on how the couple should make their investment choices: it would be bad practice to suggest putting all their money into any one of the stocks with only this dataset to look at. Expanding the dataset to include more years would be the first suggestion. This allows for a much better analysis of each stock's performance over time to be done. Additionally, there is always risk when investing. This couple needs to determine how many years they plan to keep their wealth wrapped up in the stock market, and additionally consider investing in multiple stocks if they lean towards being risk averse. Based upon the data/lack of data, and not knowing what financial goals the couple has, the leaving remark is to suggest that the couple considers meeting first with a financial advisor to determine their financial goals with investing, and then analyze an a larger span of years for the stocks in the dataset so they can make the investment choices best for them.

## Resources:
- https://www.cnbc.com/2021/03/16/one-year-ago-stocks-dropped-12percent-in-a-single-day-what-investors-have-learned-since-then.html
- https://www.investopedia.com/articles/trading/05/011705.asp
- https://www.investopedia.com/managing-wealth/when-should-you-hire-financial-advisor/
- https://www.investopedia.com/terms/s/stock-analysis.asp#:~:text=Stock%20analysis%20is%20a%20method,markets%20by%20making%20informed%20decisions
- https://www.reuters.com/markets/europe/how-ukraine-russia-war-rattled-global-financial-markets-2022-08-24/
- https://en.wikipedia.org/wiki/Code_refactoring
- https://www.wsj.com/articles/reddits-wallstreetbets-was-the-gamestop-kingmaker-but-longtime-users-say-the-thrill-is-gone-11643025602
