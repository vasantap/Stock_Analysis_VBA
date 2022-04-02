
# Stock Analysis using Excel VBA

Analyzing Wall Street Stocks data using Excel VBA


## Overview of the Project
#### This objective of the project was to create a script that loops through all the stocks
#### for one year and outputs the following:
    1.	The Ticker symbols.
    2.	Yearly change from opening price at the beginning of a given year to the closing
        price at the end of that year in $ format. 
    3.	The percent change from opening price at the beginning of a given year to the
        closing price at the end of that year while applying conditional formatting
        that will highlight positive change in green and negative change in red.
    4.  The total stock volume of the stock.
#### The data was given in two Excel sheets. They contained ticker value, the date the stock was issued, the opening price, closing price, highest and lowest price and finally the volume of the stock. Using this data, a script was created to loop through all the worksheets to initially get the opening price of a given ticker at the beginning of the year and then find the corresponding closing price at the end of the year to find out the Yearly change. Percent change was calculated by taking the ratio of the Yearly change and opening price and formatted to a percent. Finally, the Total stock volume for all the stocks of a ticker was calculated. These results were used to analyze and find the ticker and its value which had the greatest percentage increase, greatest percentage decrease and greatest total volume.
#### The results of the analysis concluded that 
    •	For 2018:
            o	Greatest percentage increase was for the ticker THB with a value of 141.42%
            o	Greatest percentage decrease was for the ticker RKS with a value of -90.02%
            o	Greatest total volume for the ticker QKN 1.69E+12 
    •	For 2019:
            o	Greatest percentage increase was for the ticker RYU with a value of 190.03%
            o	Greatest percentage decrease was for the ticker RKS with a value of -91.60% 
            o	Greatest total volume for the ticker ZQD 4.37E+12 
    •	For 2020:
            o	Greatest percentage increase was for the ticker YDI with a value of 188.76%
            o	Greatest percentage decrease was for the ticker VNG with a value of -89.05% 
            o	Greatest total volume for the ticker QKN 3.45E+12
#### To conclude, RKS continued its decrease in value through 2019 and was later replaced by VNG in 2020. The best performers were THB in 2018, RYU in 2019 and YDI in 2020.