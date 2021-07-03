# Stock Analysis

## Overview of Project

### Background
Steve would like to help his parents conduct an analysis on investments in 2017 and 2018. Though they currently invest in one stock, Steve would also like to analyze a handful of other green energy stocks to compare Total Daily Volume and Yearly Returns.


### Purpose

The purpose of this project is to analyze the yearly stock data using an automated code that Steve could alter to include other tickers or years in future analyses. To make this code efficient, I also refactored the original code in VBA to perform only one loop, so that the code could be used for much larger data sets. 


## Results
### Stock Performance in 2017 vs. 2018
After analyzing the data, I calculated the total daily volume and yearly return for each ticker over the course of a year. The yearly return was calculated by dividing the ending price by the starting price and subtracting one. To perform this calculation, I used conditions to tell when the ticker changed, so that the prices did not come from multiple stocks. The ticker daily volumes were calculated through a for loop, which added each volume until the ticker register changed. To keep track of the summations, it was important to set up multiple indexes that kept track of ticker counts. 

<img width="524" alt="Screen Shot 2021-07-02 at 10 48 52 PM" src="https://user-images.githubusercontent.com/85946042/124342171-b6456080-db87-11eb-995e-4f2ed9ea07c9.png">



To visualize the data, I printed the yearly return and total daily volume for each ticker. I created a for loop to print the saved values for each ticker before moving onto the next one. I used conditional programming to highlight the cells red if the yearly return fell below zero and green if the returns were positive. 

<img width="520" alt="Screen Shot 2021-07-02 at 10 48 56 PM" src="https://user-images.githubusercontent.com/85946042/124342265-4c798680-db88-11eb-9680-0ceea46fbdff.png"> <img width="389" alt="Screen Shot 2021-07-02 at 10 52 28 PM" src="https://user-images.githubusercontent.com/85946042/124342266-513e3a80-db88-11eb-8bc7-2e29fae035a7.png">




Stock Performance in 2017 and 2018:


<img width="228" alt="Screen Shot 2021-07-02 at 7 48 54 PM" src="https://user-images.githubusercontent.com/85946042/124338613-905f9200-db6e-11eb-8d97-a2d1dbb4cf64.png"> <img width="229" alt="Screen Shot 2021-07-02 at 7 48 40 PM" src="https://user-images.githubusercontent.com/85946042/124338617-93f31900-db6e-11eb-8468-eecfb333652f.png">


In 2017, every stock except ticker "TERP" yielded a positive yearly return. In 2018, every stock yielded negative yearly returns, except tickers "RUN" and "ENPH", which delivered positive yearly returns both in 2017 and 2018. Based on the 2017 yearly returns, it might have seemed wise to invest while there was a positive trend. Still, after seeing the 2018 returns, where majority of returns swung negative, Steve's parents may want to see more data to be able to visualize trends. In the future, it would be useful to not only add more years to our results, but also, to break the data down even further by month. Additionally, these charts do not give any context to the daily volume; if there was a volatility index, that could be helpful as well. 



### Execution Times of Original Script vs. Refactored Script

To analyze the execution times of the scripts, I had to create a Timer that started after the user inputted the analysis year and finished after the chart was created. To keep track of these execution times, I created a message box pop up that displayed the times based on the analysis year. 


<img width="172" alt="Screen Shot 2021-07-02 at 10 55 19 PM" src="https://user-images.githubusercontent.com/85946042/124342316-9bbfb700-db88-11eb-9564-36ca14a22c2f.png"> <img width="595" alt="Screen Shot 2021-07-02 at 10 55 09 PM" src="https://user-images.githubusercontent.com/85946042/124342319-9febd480-db88-11eb-9c57-65be59822ef6.png">




The original script worked through several loops before outputting a result. While the execution time for the original script was fast, 0.289 seconds for the year 2017, and 0.309 seconds for the year 2018, it was not as fast as the refactored script. The refactored script executed in 0.055 seconds for 2017 and 0.055 seconds for 2018. Though these are all quick execution times, they only analyze 12 tickers, each during one year. If Steve wanted to scale up the size of the data for analysis with more tickers or accumulated years, the smaller execution times would be more important. 

Original Code Execution Times: 


<img width="272" alt="Screen Shot 2021-07-02 at 4 42 04 PM" src="https://user-images.githubusercontent.com/85946042/124338395-622d8280-db6d-11eb-8977-a4b054b75e2b.png"> <img width="266" alt="Screen Shot 2021-07-02 at 4 42 16 PM" src="https://user-images.githubusercontent.com/85946042/124338394-5fcb2880-db6d-11eb-9d3c-7a8d9afe5377.png"> 

Refactored Code Execution Times:


<img width="262" alt="Screen Shot 2021-07-02 at 7 41 46 PM" src="https://user-images.githubusercontent.com/85946042/124338433-95701180-db6d-11eb-9f16-188de30741b5.png"> <img width="262" alt="Screen Shot 2021-07-02 at 7 41 39 PM" src="https://user-images.githubusercontent.com/85946042/124338437-99039880-db6d-11eb-943a-333c16a56b24.png">



## Summary

During this analysis, I refactored the VBA code to decrease the execution time. There are advantages to refactoring code, such as: creating quicker execution times and clearing redundant code to produce cleaner logic. A refactored script is also beneficial to the next person who may be working with the code because the algorithm is not loaded with any redundancies. That being said, it is also important to comment effectively if, for instance, there are now multiple commands aggregated in one line. Finally, refactoring can be helpful while practicing a new programming language because you might discover new syntax or looping logic. 

It could be a disadvantage to refactor the code if scripts are deleted prematurely and funcionality is lost. Also, refactoring could require more time editing compared to the time saved during execution. The time spent refactoring might not be efficient if the code will not be applied to bigger data sets in the future. Finally, while performance is enhanced through refactoring, it is also equally important to maintain readability in the code for others to interpret.

When I refactored the original code, I was very hesitant to delete any of the original script. I did not want to lose the resulting charts produced with the original code. Still, it may have been more effective to rewrite the code from scratch, which could produce clearer logic. I came across many challenges when refactoring the existing code directly; I was too attached to the logic. Once I deleted some of the original code (and saved this text to another editor) I was able to find a clearer solution to the problem. As this is my first time programming with VBA in Excel, the time I spent refactoring the code was very helpful in cementing VBA syntax and format. In this stock analysis application, I do think refactoring was beneficial because of the large amount of data that could be added with more tickers or years. Adding more data would help Steve give his parents advice because he would have more information to visualize trends. 




