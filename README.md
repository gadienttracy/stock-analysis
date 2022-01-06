# Ticker Analysis with Excel

## Overview of Project
The purpose of this study is to refactor the code used to evaluate stocks from two years in order to create a faster query.  The output is the Total Daily Volume and the Percent Return for 11 ticker (stock) values for one year (2017 or 2018).

## Results
2017 was a very good year for the majority of the stocks listed. DQ had the best year with 199.4% increase, however its total daily volume (35,796,200) was significantly less than the average (263,886,5920). TERP was the only stock to see a negative return of 7.2%.  The original script for this script ran in 1.203 seconds.  The refactored script ran significantly faster, 0.195 seconds.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/91269696/148456407-24a1ef10-e3a7-48f0-885b-c9a69330b575.PNG)


There were only two stocks with positive returns for the year 2018, ENPH and RUN.  DQ and JKS had the worst percentage return in 2018, -62.6% and -60.5% respectively.  Making DQ one of the most volatile stocks from the previous year.  The original script for this query ran in 0.953 seconds for the year 2018.  The refactored script returned results in 0.445 seconds.

![VBA_Challenge_2018](https://user-images.githubusercontent.com/91269696/148456424-90f2c6d6-d435-4c06-97a2-ce54d9e95bb7.PNG)

The largest factor in reducing the time to run the script was the addition of the tickerIndex which allowed the script to count each Row and add it to the tickerIndex rather than calculating the value of the tickerIndex at the end using the starting and ending price.

![original code](https://user-images.githubusercontent.com/91269696/148458542-7aab62a5-d39f-4579-a204-fa727fdcf7be.PNG)
* *original code* *

![Refactored](https://user-images.githubusercontent.com/91269696/148458567-0d53fea1-f336-4743-83c9-e46f820f2b41.PNG)
* *refactored code* *

## Summary
**What are the advantages or disadvantages of refactoring code?**
    The advantages to refactoring code are increased speed of the script and easier readability for future users.  Reducing complexity of code can increase accuracy for those that need to refactor the code in the future.  A main disadvantage of refactoring code is the time it takes to do so.  The programmer should be able to weigh the pros and cons of refactoring code before beginning.  Will the refactoring make a substantial user experience improvement?  Will this code be relevant in 1, 2, 10 years when it may need to be refactored again?  These are all questions that should be asked by the programmer before investing the effort to refactor a working code.
**How do these pros and cons apply to refactoring the original VBA script?**
    This VBA script only analyzes two worksheets (2017 and 2018).  The time savings of the script run for the user may be negligible.  If more years of ticker values are added to the sheet it could make a more significant impact and would also be worth refactoring more frequently as to stay current and ease future refactoring.

In summary this project helped to determine the returns of 11 stocks and their daily volume.  The refactoring of the scripts used to evaluate these worksheets saved the user fractions of seconds and increased the readability of the VBA script.
