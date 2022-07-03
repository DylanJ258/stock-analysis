# stock-analysis
## Overview
* Steve wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.
* In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.
## Results
* 2017 Stocks
  * In 2017, most stocks had a positive return. The only one that was negative was TERP, with a -7.2% return. The stock with the highest return was DQ at 199.4%
  * The stock with the highest daily volume was SPWR at 782,187,000
  * The stock with the lowest daily volume was DQ at 35,796,200
  * <img width="232" alt="2017_Chart" src="https://user-images.githubusercontent.com/104036750/177048085-c5668ce4-6767-4e8e-b9d7-69245de23b8e.png">
* 2018 Stocks
  * Stocks in 2018 did't fair as well as 2017. All but ENPH and RUN had negative returns. The stock with the best return was RUN at 84.0% while the stock with the worst return was DQ at -62.6%
  * The stock that had  the highest daily volume was SPWR at 538,024,300. SPWR had the highest daily volume in 2017 as well, but there was a decrease of almost 200,000,000
  * The stock that had the lowest daily volume was AY at 83,079,900
  * <img width="221" alt="2018_Chart" src="https://user-images.githubusercontent.com/104036750/177048297-f44cc0e0-0abd-40c7-be27-c2c69bb5f6df.png">
## Summary
* An advantage of refactoring code is that it optimizes its performance. The code will run faster and more smoothly. It will also be easier to read, which makes the analysis of the project much more simple and quick.
* A disadvantage could be the time spent refactoring. For some projects, it might not be worth the extra time spent on refactoring.
* The original script from the modules ran much slower than the refactored code. For example, the 2017 code in the orignal script ran in 0.7148438 seconds while the refactored code ran in 0.1367188 seconds.
  * This is almost a second difference, which isnt extremely significant, but if this project had a larger data set, the difference would be much greater.
  * Orignal time for 2017 and 2018
  * <img width="254" alt="2017_Original_Time" src="https://user-images.githubusercontent.com/104036750/177048715-968b554c-ab6a-4490-9423-392f76b2b9e8.png">
  * <img width="254" alt="2018_Original_Time" src="https://user-images.githubusercontent.com/104036750/177048754-6dd0737e-ae82-4b34-9187-877e8a75052c.png">
  
  * Refactored time for 2017 and 2018
  * <img width="257" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/104036750/177048766-6889a403-9fc2-4eaa-a472-914517c7c926.png">
  * <img width="251" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/104036750/177048776-6f60619d-0b5f-4612-8c43-ad376a0d436f.png">

