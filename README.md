# Stock-Analysis
## Analysis of stock using VBA
##### **Overview of Project**
Steve needed a quick and easy way to compare stocks.  Using VBA, we are easily able to provide Steve with a tool to easily check on stock performance based on the year he chooses to see the data.
In this module, we created the necessary VBA coding to show Steve what he is looking for to analyze stock performance.  We also learned how to refactor the code to try and get the Macro to run at better performance levels.  This will be helpful if the data set is increased over time and Steve needs the macro to run smoothly each time it is performed.

##### **Results (Before Refactoring)**
When running the macro to do an analysis of 2017 stocks, the macro took 1.363281 seconds to run.  The same macro to analyze the 2018 stocks took 1.203125 seconds

![Before_Refactoring_2017](https://github.com/barrettben/Stock-Analysis/blob/ecbc2f4720061ca08c5e8c94e4d1584a0f2818a0/Resources/Before_Refactoring_2017.png)

![Before_Refactoring_2018](https://github.com/barrettben/Stock-Analysis/blob/ecbc2f4720061ca08c5e8c94e4d1584a0f2818a0/Resources/Before_Refactoring_2018.png)

##### **Results (After Refactoring)**
When running the macro to do an analysis of 2017 stocks after refactoring the VBA code, the macro took 1.289063 seconds to run.  The same macro to analyze the 2018 stocks after refacotring the VBA code took 1.167969 seconds

![VBA_Challenge_2017](https://github.com/barrettben/Stock-Analysis/blob/4c7704c1741d8eb258e2dc8892e10e869e84e6be/Resources/VBA_Challenge_207.png)

![VBA_Challenge_2018](https://github.com/barrettben/Stock-Analysis/blob/4c7704c1741d8eb258e2dc8892e10e869e84e6be/Resources/VBA_Challenge_2018.png)

##### **Summary**
**1.) What are the advantages or disadvantages of refactoring code?**

##### Advantages

1. Improves the design of the software/code
2. Makes the software/code easier to understand
3. Helps the coder find bugs in the original code
4. Helps programming go faster

##### Disadvantages

1. Takes funding resource for developers time
2. Takes extra time

**2.) How do these pros and cons apply to refactoring the original VBA script?**

1. There wasn't much data in the excel document to justify refactoring the code because it didn't take very long to run prior to the changes
2. The code was already easy to understand and likely didn't need edits
3. No bugs were found as the code returned the same information
4. Refactoring took the coder extra time, therefore, would have cost the business extra money and could have delayed the start of another project to be coded
