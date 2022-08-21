# Stock-Analysis with Excel (using Visual Basic for Applications)

## Overview of the Project
Steve has asked for help in analyzing the stock market to see which ones are worth investing in. This challenge was aimed at testing if refactoring the code shortens the time taken for the analysis process.

## Analysis and Challenges

Let us take a look at the data provided to us. We have data on 12 different stocks, over the years 2017 and 2018. 
We can see the _Names, Dates, Stock value information_ and _Volume of stocks traded._ 
We then created macros that would allow us to analyze how well each ticker did for the years 2017 and 2018.

If we take a look at the initial analysis times for [2017](https://github.com/SoumyaAbraham/Stock-Analysis/blob/main/Stock%20Analysis_Misc/Initial_2017.PNG) and [2018](https://github.com/SoumyaAbraham/Stock-Analysis/blob/main/Stock%20Analysis_Misc/Initial_2018.PNG), we can see analysis takes well over a *second*. Is tjere a way we can make this process even more efficient? This is where refactoring the code comes into play.

---

### What _is_ Refactoring?

Simply put, refactoring is the process of restructuring our initial code to make it more efficient without altering the core functionality. 

Let us take a look at a portion of the initial code- 

![Initial Code](https://github.com/SoumyaAbraham/Stock-Analysis/blob/main/Stock%20Analysis_Misc/InitialCode.PNG)

Notice that the presence of two loops, simultaneously. 

Now let us take a look at the same piece of code, rafactored-
![Refactored Code](https://github.com/SoumyaAbraham/Stock-Analysis/blob/main/Stock%20Analysis_Misc/RefactoredCode.PNG)

Here, by introducing the tickerIndex and Output arrays for our Ticker Volume, Starting Price and Ending Price, we make it a smoother process in the background. 

This is clearly visible when we look at how long the same analysis takes when using our Initial Code and Refactored Code for the year 2017 - 
![2017 Initial](https://github.com/SoumyaAbraham/Stock-Analysis/blob/main/Stock%20Analysis_Misc/Initial_2017.PNG) ![2017 Refactored](https://github.com/SoumyaAbraham/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.PNG)

A similar cutdown in time can be seen for analysis for the year 2018-
![2018 Initial](https://github.com/SoumyaAbraham/Stock-Analysis/blob/main/Stock%20Analysis_Misc/Initial_2018.PNG) ![2018 Refactored](https://github.com/SoumyaAbraham/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.PNG)


We did not stop at that thought. Using the [*InputBox()*] feature, we made the program User-interactive. 
We even added buttons so Steve can _literally_ do his analysis with the click of a [button]!
And if he feels like redoing it or moving on to the next one, all he needs to do is click the [Clear] button and he has a blank canvas to start over. 
---
#### Benefits of Refactoring code:
1. Refactoring can improve the design of software
2. Helps in the debugging process
3. Executes the program with greater speed. 

*As an added bonus, it can help change the way you think as a developer.*

So if refactoring is so positive, why is it not the go-to way of coding? Let us find out.

#### Disadvantages to Refactoring code:
1. If the refactoring is not precise, it can cause new bugs and errors in your code
2. It is often more time consuming to refactor a code, especially when a larger team is involved. There is also surprisingly little recognition for the hard work, since from the User perspective, the improvement is often not that noticable. 

---

As we saw above, there was quite an improvement in the time taken to execute the analyses when we used the refactored code. 
In this case, since we only analyzed a little over 3000 points, that improvement may not seem as big a deal. 
However, if you consider other scenarios where you have to go through much heavier datasets, the time improvement may be much more valuable. 
That being said, you can also see the amount of extra coding required for refactoring the initial code. This can be time consuming, confusing and causing a much higher chance for errors. 

In the end, it comes down to weighing the pros and cons. The possible impact refactoring your code could have on your software Vs the amount of time you are willing to spend perfecting. 

Either way, it is always good to know you have the choice of playing arounf with your code and making it more effective! 



