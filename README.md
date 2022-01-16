# Stock Analysis with VBA

## Overview of Project

### Purpose
Steve is happy with the workbook in its cuurent state. However, he would like to expand the dataset to include the entire stock market. The purpose of this project is to refactor the code so it loops through the data only one time to collect the same information. Then, determine if refactoring improved the run-time of the code.

## Results

### Performance Analysis
The year 2017 saw mostly positive returns on the stocks.  However, the year 2018 was much different with a majority of the stocks showing negative returns.

### Code Analysis
#### Original Code
<img width="512" alt="VBA_Code_All_Stocks_Analysis" src="https://user-images.githubusercontent.com/59906657/149678540-776064b2-72cd-45cd-a4b0-762723dafe23.PNG">

The above, original code uses a nested for loop to loop through the items in the list.  This makes the computer take more steps, and while this may be okay for smaller amounts of data, large amounts of data could put more stress on the system.  As the following two images show, the code ran in .921875 seconds for 2017 and in 1.113281 seconds for 2018.

<img width="800" alt="VBA_Modules_2017" src="https://user-images.githubusercontent.com/59906657/149643288-3498233e-f939-4448-8d65-0a1bc0e4fc7d.png">

<img width="800" alt="VBA_Modules_2018" src="https://user-images.githubusercontent.com/59906657/149643336-16992000-c3d9-4d53-ace3-b922783bcfcb.png">

#### Refactored Code
<img width="671" alt="VBA_Code_Refactored" src="https://user-images.githubusercontent.com/59906657/149679049-be58474d-fbfd-40b3-8cf4-3b1d50903ed1.PNG">
The refactored code utilizes only one for loop and arrays in order to accomplish the same task.  With one for loop, the computer only has to go through the data once, speeding up the process.  The below images show the time improvements.  This time the code ran in .125 seconds for 2017 and in .109375 seconds for 2018.

<img width="800" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/59906657/149643234-95950ade-cd60-4f2a-b02e-2939ac97caa4.png">

<img width="800" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/59906657/149643237-6715a8c6-73b7-4f21-9155-a434a1db8146.png">


## Summary
### General Advantages and Disadvantages
One of the main advantages of refactoring is that it can lead to better quality code. Through better quality code, the program design improves and it can run faster.  It can also become cleaner and more organized which will help with maintenance.  

A disadvantage is that it can be time consuming, so it should only be done if time is available in the project budget.  Also, if the application is large, there is a lot more code to make sure still works together.

(https://stackoverflow.com/questions/43983284/what-are-the-advantages-and-disadvantages-of-refactoring-code-smell-in-software)

### Advantages and Disadvantages of Refactored Project Code
The original code used nested for loops and variables to loop through the stock tickers. The variables were easier to understand and work into the code. The disadvantage was that it looped through the tickers twice to accomplish its task. This makes the code take longer to run and would stress the system if used on very large datasets.

The refactored code used one for loop and arrays to accomplish the task.  The advantage was improved run-time since the tickers were looped over once.  The disadvantage was that it was time consuming to figure out how to use the arrays in place of variables.
