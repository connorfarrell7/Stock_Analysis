# VBA_Challenge

## Overview
Steve’s parents are interested in investing in the stock DQ, and he wants to analyze a worksheet of stock data to see if this is a good decision. He is somewhat experienced with excel but he needs some help. I am going to make a macro using Visual Basic for Applications (VBA) to help Steve automate this process. I will produce some recommendations based on the data that I have found. Instead of looking at only the ticker DQ, I will be looking at all the available tickers.

### Analysis
The first step towards analyzing the data was to widen the scope of the analysis from DQ to all available stocks. The macro takes user input to determine which year’s data to analyze. (Figure 1)

![fig 1](https://user-images.githubusercontent.com/86024575/125216725-0f4e7c00-e28d-11eb-9e9c-288972ac3cd7.png)

*Figure 1*

Next, I created arrays in VBA to keep track of three values: the total volume traded daily, the starting price, and the ending price of each individual stock. (Figure 2)

![fig 2](https://user-images.githubusercontent.com/86024575/125216898-779d5d80-e28d-11eb-97d0-24a80e9f9260.png)

*Figure 2*

Then I set each stocks total volume to start at zero. (Figure 3)

![fig 3](https://user-images.githubusercontent.com/86024575/125216991-a7e4fc00-e28d-11eb-8363-04a1930b6076.png)

*Figure 3*

The macro loops through each stock to fill each created array with the appropriate values. (Figure 4)

![fig 4](https://user-images.githubusercontent.com/86024575/125217142-fabeb380-e28d-11eb-91ef-e9dea399c1dc.png)

*Figure 4*

Finally, our script loops through each index in order to print each value into their respective spots in the resulting table. (Figure 5)

![fig 5](https://user-images.githubusercontent.com/86024575/125217245-2f326f80-e28e-11eb-89ef-dd1d72c9f47b.png)

*Figure 5*

### Results
From the data provided we can see that although DQ had a great year in 2017 (+199.45%), while that was not the case in 2018 (-62.6%). Since this shows inconsistency, DQ might not be the best stock to invest money in. Only two stocks continued to appreciate over the course of the data collection period, these being RUN and ENPH (Figures 6 & 7). I would recommend ENPH to Steve and his family due to its consistent growth, and its performance is better than RUN.

![fig 6](https://user-images.githubusercontent.com/86024575/125218751-8c7bf000-e291-11eb-8417-5403170c32f3.png)
![fig 7](https://user-images.githubusercontent.com/86024575/125218581-50488f80-e291-11eb-84c2-1b29d1be2f68.png)

*Figures 6 & 7*

#### Refactoring
Refactoring the script greatly increased the speed of the analysis. My original macro took roughly 20 seconds to run (Figures 6 & 7), whereas the refactored version ran in less than 1 second (Figures 8 & 9). This is possibly due to removing the nested loops that were present in the original VBA script, meaning that the code would not loop through redundant processes each time. Having separate loops for setting the ticker index, gathering the values, and printing the values made it a more streamlined process. A disadvantage to refactoring code is that it is more time consuming, however it is evidently worthwhile.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/86024575/125218447-065fa980-e291-11eb-96e4-f715c5301cb9.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/86024575/125218454-095a9a00-e291-11eb-9735-8b9f864ba5f1.png)

*Figures 8 & 9*
