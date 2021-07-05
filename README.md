# VBA-Challenge

  ## Overview of the Project
  Refactor code in VBA to lower the run time of running a macro to loop over a large datasets by adding arrays to create fewer steps.
  
  ### Purpose
The purpose of the VBA Challenge was to practice coding on VBA to set variables, run for loops, nested loops, create arrays, and format cells. By refactoring a previous code       
This VBA Challenge helped to gain experience in not only creating code, but refactoring it as well to make projects run more efficiently. By creating arrays to lower the number   
of steps needed to perform a function we were able to lower the run time of the macro to a fraction of a second. Doing so allows more data to be added to the spreadsheet and the   
results will still appear quickly.

  ## Results

   ### 2017
  Before refactoring the code we can see that the run time for the year 2017 was 0.94 of a second to loop over thousands of rows of data. By refactoring the code by adding 3 
  new arrays which included tickerVolumes(12), tickerStartingPrices(12), and tickerEndingPrices(12), the run time was brought down to roughly 0.19 of a second. By creating
  the 3 arrays and creating a tickerIndex variable makes it possible to avoid having to create a nested loop. This allows the code to run much quicker. As a result we can 
  quickly see that in 2017 the out of the 12 stocks we ran analysis for 11 of them had positive return. Only "TERP" had a negative return of -7.2% while the highest return
  was "DQ" with 199.4% return.
  ![image](https://user-images.githubusercontent.com/85451089/124416106-defe5f00-dd1b-11eb-97be-01dc329b24dc.png) ![image](https://user-images.githubusercontent.com/85451089/124416165-fdfcf100-dd1b-11eb-989c-4b1aba12a132.png)
  The image on the top shows the run time of the code before being refactored while the one on the bottom shows the more efficient refactored method.
  
   ### 2018
  By adding "2018" into the input box shown when we run the analysis without refactoring we can see that it takes a total of about 1.02 seconds to display the data.
  Once refactored the total run time after inputting "2018" into the input box is about 0.19 of a second. By refactoring the code we were able to cut the run time by 
  over 80%. The results shown once running the analysis shows that only 2 of the 12 stocks showed a positive return while the other 10 were negative. In 2018 out of 
  the 12 stocks "RUN" had the highest gain with 84% return while "DQ" went from the highest in 2017 to the lowest in 2018 with a -62.6% return. 
  ![image](https://user-images.githubusercontent.com/85451089/124416821-757f5000-dd1d-11eb-8b2f-5c7d17c04349.png) ![image](https://user-images.githubusercontent.com/85451089/124416836-7b753100-dd1d-11eb-9a35-a1c005e50a56.png)
The first picture above shows the run time of the code before being refactored whie the one below it shows the refactored method.

  ## Summary

   ### What are the advantages or disadvantages of refactoring code?
  The main advantage from refactoring code that is seen in this challenge is cutting down the run time by creating less steps. By using less steps it means that
  there is less coding that is needed to be debugged if needed. There are not many disadvantages to refactoring code so long as there are proper notes to let
  others know what is going on or to remind yourself of the steps taken. While there are less steps when refactoring a code it can mean finding and using
  new codes that you may not be too familiar with.
  
   ### How do these pros and cons apply to refactoring the original VBA script?
  The pros of refactoring this code means we are able to not only analyze the 12 tickers in the data sheet quicker, but it'll be easier to add more tickers to
  analyze for future reference. This means that we can add plenty more data to these datasheets and still receive the information needed in a quick and timely
  manner. 
  As for the cons, it took a good amount of research to find the exact codes that I needed to finish refactoring. Having the previous code was a great
  help as it would take swapping out some variables for the new arrays created to run more smoothly. 
