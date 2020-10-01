# Refactoring Result of Stock Analysis Excel Macro

## **Overview of Project:** The purpose of this report is to discuss the impact of code refactoring on an Excel VBA Macro that is used to analyze stock trading volumes and returns. While both analysis files produce the same analysis, the refactored code demonstrates and reduction in run time. The refactoring adjusts the for loop structure to create a more efficient execution of the analysis.

## Results
### Unrefactored Results
![Unrefactored 2017 Analysis](https://github.com/zborglin/stock-analysis/blob/master/Resources/VBA_Challenge_2017_unrefactored.png)
![Unrefactored 2018 Analysis](https://github.com/zborglin/stock-analysis/blob/master/Resources/VBA_Challenge_2018_unrefactored.png)

### Refactored Results
![Refactored 2017 Analysis](https://github.com/zborglin/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)
![Refactored 2018 Analysis](https://github.com/zborglin/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png)
#### **Analysis of Results and Description of Refactoring:** As shown in the refactored analysis above, the runtime decreased by **0.594 (76%)** and **0.635 (78%)** for the 2017 and 2018 analysis respectively. This improvement in performance is attributed to an adjustment of the for loop structure as described in the images below.
![Unrefactored Code](https://github.com/zborglin/stock-analysis/blob/master/Resources/Unrefactored_Code.png)
##### In the original code (unrefactored), a nested for loop is employed to loop through all data for each ticker. In the refactored code shown below, the for loop structure is flattened and the data is extracted for each ticker during a single loop that utilizes an index variable to fill arrays with the relevant information.
![Refactored Code](https://github.com/zborglin/stock-analysis/blob/master/Resources/Refactored_Code.png)

##Summary
### What are the advantages or disadvantages of refactoring code?
In most cases, the first draft of a coding application can be characterized as having inefficiencies and less than clear formatting. Refactoring is an opportunity for developers to review code to find better ways of accomplishing a goal and a more clear way to communicate the functionality. The benefit of this process is to create a better script that can be implemented with fewer resources and can be understood by someone looking at the code for the first time. A disadvantage to refactoring is that it requires time and energy to optimize a code. In some cases, it may be more appropriate to stick with a less efficient code that still provides the correct output if it is not important to scale the code to handle larger amount of data. 
 
### How do these pros and cons apply to refactoring the original VBA script?
In the case of the stock analysis VBA script discussed in this report, refactoring will enable the tool to handle much larger amount of information in a more reasonable computation time. A runtime decrease of ~77% might translate to 4x faster calculations for a larger dataset which will increase the burden on the computer. However, in order to implement this improvement is took time to refactor the code which in some cases might not be the most useful allocation of development time depending on the situation.
