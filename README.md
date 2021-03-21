# Stock-Analysis

## The Project
Steve has asked us to assist in helping him look at Stocks for his parents invenstments into DAQO New Energy Corp(DQ) as well as other similar companies in the Stock Market. Taking data from the Stock Market and importing into excel, we will be using VBA scripting to help us to look at DQ as well as other companies numbers from the last few years. Specifically 2017 and 2018. What we want to know is if DQ is worth investing into, or should Steve's parents look into diversifying their portfolio to yield better results. We will be using VBA scripting to loop through the data that will allow us to compare a different variety of companies of their starting and ending price from 2017 and 2018.

### The Data Set
The data set we received from Steve showcased twelve potential companies, including DAQO, and their numbers from Year 2017 and 2018. Here it gives us a breakdown by the Stock ID, date, their opening price, their highest and lowest rate at during the time period, their closing price, and the their adjusted closing price. There is a lot of data that needs to be sorted and calculated for our purpose. What we are looking for is the total daily volume and the percentage return.

### The Method
The process begins with what exactly are we looking for. What we are looking for begins with creating a table that has the 12 companies stock ID, we will call these Tickers, and then assigning two columns, one for Total Daily Volume and the other for the Return in percentage form. Once we've got our basic outline of information that we want, we begin looking at VBA Scripts to help us run though the information from 2017 and 2018.

Starting with the Tickers, we needed to give them a value. There are 12 Tickers in alphabetical order, so we assign each ticker a vale from 0-11. Once we've assigned them their value, we write up a piece of script that will give us total daily Volume.

![Screen Shot 2021-03-21 at 11 58 58](https://user-images.githubusercontent.com/79731109/111913666-2b5bb280-8a3d-11eb-88cf-de209f969c8c.png)

The script we write will take a look at the information in column A and pull information from column F to determine both our starting price and ending price for our tickers. This will help calculate all our total daily volume the reutrn percentage for each year and the respective Tickers.

Once we've gotten our information, we need to find a way to output our information on to our worksheet.

![Screen Shot 2021-03-21 at 12 00 09](https://user-images.githubusercontent.com/79731109/111914183-f3556f00-8a3e-11eb-9b76-73832b7872dc.png)

We then take our information and write up some code to output what we found in our new worksheet.

### The Results

##### 2017
![VBA_Challenge_2017](https://user-images.githubusercontent.com/79731109/111914092-a7a2c580-8a3e-11eb-809c-9a562990c58d.png)

##### 2018
![VBA_Challenge_2018](https://user-images.githubusercontent.com/79731109/111914103-b2f5f100-8a3e-11eb-9ce0-dabecb83cd0d.png)

What we can conclude from our information is that 2017 was a very successful year if you had purchased stock prior to the positive return rate. We see several declines in 2018 for our companies in their return rate. To Steve's Parents, this may look bad, but this gives them an opportunity to purchase stock while the prices are at a low and potentitally sell at a high point if they wish to sell. 

## Summary

The code and script that we were given for our purpose was to help us learn about refactoring, or make more efficient, codes and scripts. When given an inital program, code, or script, it can be pretty messy and hard to understand. Refactoring the code allows us to make it more efficient while not changing it's original function. This will help make the code easier to read and understand what it's doing. The downside to refactoring is having to be careful not to change it's original functions. Which can be pretty easy if you don't input the write values or change up the equation. Refactoring is just another process to "Clean" up the code and set it up in a way where it can run it's program more efficient as well as legible to the next person who has toeither add on to it, or make adjustments to it. 
