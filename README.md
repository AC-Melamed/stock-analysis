Module 2: Deliverable 2
 
# Stock Analysis with VBA in Excel
 
 
 
## Overview of Project
 
This project consists of a module written in VBA within the Microsoft Excel environment which allows for the automated analysis of a corpus of stock prices data spanning the years 2017-2018 for a selection of companies producing 'green' technologies.  The data that the client commissioning the project was interested in is the "Total Volume" of stocks traded for each company (i.e. the sum of all daily volumes) and each company's ultimate "Return" on investments calculated as a percentage.        
 
 
### Purpose
 
The purpose of this project was ultimately to provide the client with an optimized means of automatically analyzing the full range of stock data from multiple companies included in the provided dataset in order to easily quantify the comparative performances of each over the 2017-2018 time period.  With this analysis accomplished, the client hoped to make an ideally informed investment decision.    
 
 
### Background
 
After an initial analysis targeting a particular stock identified by the client as the most likely candidate for investment revealed subpar performance, this analysis' scope was expanded to the full dataset of stocks provided, for which a VBA macro was programmed.  The original version of this macro functioned properly but reported sub-optimal runtimes, and was subsequently refactored with a successful dramatic decrease in reported runtimes.  This final, refactored version is the file pushed to the current version of this repository.  
 
 
 
## Results
 
The results of this project's final analysis are as follow:  
 
### Comparing Stock Performance Between 2017-2018
The quantitative results of the analysis of stock values for this project can be seen below.  These results show that the companies "ENPH" and "RUN" are the only one whose stocks made positive returns in both 2017 and 2018, with the percentage returns for "ENPH" being comparatively great enough in both years to qualify it as the most likely investment for the client.  
#### Analysis Results 2017
![2017 Results](https://github.com/AC-Melamed/stock-analysis/blob/main/Resources/VBA_Challenge_RESULTS_2017.png "2017 Results")
#### Analysis Results 2018
![2018 Results](https://github.com/AC-Melamed/stock-analysis/blob/main/Resources/VBA_Challenge_RESULTS_2018.png "2018 Results")
 
### Comparing Execution Times Between Original and Refactored Code
 
The original version of the VBA macro code relied on a series of nested 'For' loops whereby the total volume and stock price values were calculated linearly in the order they appeared within the data set and only collated under the label of their respective stock tickers at the end of the process.
#### Pre-Refactoring Code
![Pre-Refact Code](https://github.com/AC-Melamed/stock-analysis/blob/main/Resources/VBA_Challenge_MacroCode_PRE-REFACTORING.png "Pre-Refact Code")

The original, pre-refactoring code reported runtimes of slightly under 1 second for each year's dataset.  
#### Pre-Refactoring Runtime 2017
![Pre-Refact Runtime 2017](https://github.com/AC-Melamed/stock-analysis/blob/main/Resources/VBA_Challenge_2017_PRE-REFACTORING.png "Pre-Refact Runtime 2017") 

#### Pre-Refactoring Runtime 2018
![Pre-Refact Runtime 2018](https://github.com/AC-Melamed/stock-analysis/blob/main/Resources/VBA_Challenge_2018_PRE-REFACTORING.png "Pre-Refact Runtime 2018") 
 
The original code was subsequently refactored such that the total volume and price values were defined as their own arrays with set dimensions for being indexed using the 'ticker' array which stored the 11 distinct stock tickers.  The 'ticker' array was incorporated into the 'For' loops in a way that allowed for the volume and price values for each distinct stock to be calculated and summed before the next, dramatically reducing the runtime required to complete the analysis.  
#### Refactored Code
![Refactored Code](https://github.com/AC-Melamed/stock-analysis/blob/main/Resources/VBA_Challenge_MacroCode_REFACTORED.png "Refactored Code") 

After refactoring, the macro ran much more quickly while producing the same results for each year, completing for both in nearly a tenth of the original times.
#### Refactored Runtime 2017
![Refact Runtime 2017](https://github.com/AC-Melamed/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png "Refact Runtime 2017") 
 
#### Refactored Runtime 2018
![Refact Runtime 2018](https://github.com/AC-Melamed/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png "Refact Runtime 2018") 
 
## Summary
The broad conclusions of this project are hereby summarized:
 
### Advantages and Disadvantages of Refactoring Code (General)
There are many reasons why a program script such as the one created for this project might be subjected to refactoring.  Like when revising a first draft of an essay for editing, a program can be refactored for clarity and concision by consolidating unnecessarily verbose code to make it run faster, take up less storage space, or parse more easily for the user.  In other cases, a script might be refactored to make it more versatile by incorporating more dynamic elements that can be adapted to a broader range of data or use cases besides those it was originally designed for.  However, in each of these cases there are disadvantages involved that must be accounted for.  Paring down existing code too aggressively runs the risk of reducing functionality or even breaking critical components, while more dynamic code can be more unpredictable especially if the refactoring is performed without a clear conception of what the additional use cases being accommodated might actually entail in terms of data structure.            
 
 
### Advantages and Disadvantages of Refactoring Code (Particular)
 
In the particular case of this project, the advantages of refactoring the code certainly outweighed the disadvantages.  The most evident advantage and the main goal of the refactoring was the successful decimation of the typical runtime for the VBA macro utility.  This was accomplished with no compromise to the accuracy of the data analysis itself, a very small hypothetical decrease in versatility due to the hardcoded dimensions of certain arrays, and an ultimately negligible (but technically reductive) change in file size.  New comments were added to prevent any issues with legibility.  In the final accounting, it is clear that the refactored version of this project's code is the more optimized version.  


