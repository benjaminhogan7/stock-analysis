#VBA_Challenge

## Overview of Project and Purpose

This projected sought to analyze daily stock price and volume for a variety of green energy companies and to provide trading volume as well as a yearly return. Given that the client wishes to expand his searches beyond just green companies, the goal of the project was to refactor the code to run more quickly and efficiently.

##Challenges

I found this project to be extremely challenging and was ultimately unsuccessful in refactoring the code, despite my best efforts. The primary challenge for me seemed to be in creating output arrays that stored information and then ultimately outputted information to the excel worksheet, as well as creating an index for this use. I struggled to find documentation and code to reuse or repurpose for this. 

Some of the sites I tried to reference included:
https://docs.microsoft.com/en-us/dotnet/visual-basic/programming-guide/language-features/arrays/
https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/can-t-assign-to-an-array
https://docs.microsoft.com/en-us/office/vba/language/concepts/getting-started/type-conversion-functions
https://docs.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-conceptual-topics

I also tried to attend office hours, however, the instructor was busy helping another student and the TA was unable to assist me with the errors that I was getting. Ultimately this was very frustrating for me. I tried to write and rewrite the code and couldn’t get the refactored code to work. 

##Analysis

Green energy companies had a very positive 2017 with many companies in this set seeing tremendous growth in their stock prices. However, in 2018, nearly all companies in this set saw substantial losses in their stock prices. Two companies saw increases in both years, ENPH and RUN. In general, I would say that based on the information here, these companies would not be great investments. My previous knowledge informs the next statement: that this industry’s success can depend on contemporary politics and government subsidies (or lack thereof).

In addition to the volatility of the stock prices, the trading volumes of individual stocks changed wildly during these two years.
<img width="545" alt="Screen Shot 2022-07-14 at 12 22 08 AM" src="https://user-images.githubusercontent.com/108236450/178901499-a55fb6fb-2c41-448d-8bd1-6049774a2ec9.png">
<img width="503" alt="Screen Shot 2022-07-14 at 12 22 28 AM" src="https://user-images.githubusercontent.com/108236450/178901511-d024a17a-bba3-4851-89a3-59115b4f7947.png">

##Summary

The main advantages that I could see with refactoring the code is to 
	•	Provide quicker results through using indexes and arrays and generally be more efficient
	•	Easier to understand and modify
	•	Allow for the analysis of larger data sets
	•	More flexibility 
	•	Using less memory to run
(https://www.coscreen.co/blog/refactoring-your-code-in-agile/)

Refactoring was a new concept for me, however, it makes sense. We want to be able to to make changes quickly and easily, and it seems that there are usually more elegant solutions to the problems we are trying to tackle.

On the other hand, this did not work for me and from what I gather from doing research, I’m not alone. It seems that refactoring introduces the potential for new bugs, requires retesting, and in general is time intensive. Still, I can see the benefits of putting in the work in many cases.

For the data that we were asked to analyze, some 3,000 rows, I would say that the original code did just fine in analyzing the information quickly — in less than a second. Unfortunately, I cannot confirm whether or not the refactored code would offer a significant improvement.
<img width="259" alt="Screen Shot 2022-07-13 at 11 28 32 PM" src="https://user-images.githubusercontent.com/108236450/178901604-76c9455a-ab70-4ad1-a69a-90ed95677466.png">
<img width="265" alt="Screen Shot 2022-07-13 at 11 28 20 PM" src="https://user-images.githubusercontent.com/108236450/178901613-e293fb8e-6859-4fff-a1fb-5f2c2c2229a9.png">

