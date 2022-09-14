# VBA_Challenge
For VBA_Challenge_2018
## Overview 
  The purpose of this analysis was to categorize unique ticker entities while providing some data on each ticker specifically in 2018. The data on each ticker entity is the total volume each ticker was sold and the return on each ticker. This gives us insight on popular tickers while also providing information on the finacial gain from each ticker. 
## Results  
  When we compare the Refactored and Original we notice a huge difference not in the outcome, but in the total run time.  
  ### Original 
  In the Original code we achieved the outcome by using nested loops to loop through the ticker array while also looping through the data sheet. The code looks like this:  
  #### Original Code
  '4) Loop through tickers
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 To RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then
           totalVolume = totalVolume + Cells(j, 8).Value 
           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
           startingPrice = Cells(j, 6).Value
           End If
           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
           endingPrice = Cells(j, 6).Value
           End If 
           Next j 
  ### Original Analysis
This code does get the job done, however it must loop through the sheet multiple times which is even more a burden since the sheet has a lot of data. This makes the code work hard to achieve the same output. 
   ### Refactored   
   In this version of the code instead of using nest for loops in order to aggregate the information we instead use a tickerIndex which will incremently increase as we loop through the data sheet. The code looks like this: 
   #### Refactored Code 
   '2a) Create a for loop to initialize the tickerVolumes to zero.
    
   For i = 0 To 11
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
   
   ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1) = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6)
        End If
        '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1) = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6)
         End If
        '3d Increase the tickerIndex.
             If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
            End If
        Next i
   ### Refactored Analysis 
   We can visually see the difference between the Refactored and Original code, however it may not be obvious on what is happening. In this case we threw out the nested loop approach and instead are using the array that was established earlier. This way we set all initial values to 0 for all tickers so once we loop through the data the values will not stack. Instead of looping through the whole data sheet 12 times for each ticker we are essentially looping through the data sheet once since we are increasing the tickerIndex everytime the current ticker value does not match with the next ticker value in the array. This greatly decreases the time it takes for VBA to run our code since it is essentially running through the data sheet once with a new tickerIndex value each time the previous tickerIdenx value as already been accounted for. This is further supported by the attached run times. 
   ## Summary 
   ### Advantage
  The main advantage to the Refactored code is that it runs faster than the Original code. This mostly has to do with the logic of the code being cleaner, but not having to loop through the data sheet mutliple times. The code would also be easier to manipulate if more data were to be added to the sheet later on by manipulating the array and output array. If there was an addition of more data in an organized manner the Refactored code will be able to run relatively quickly. If more data were to be added with the Original code then it would take even longer since the for loops will just become longer.
  ### Disadvantage 
  I think one disadvantage would trying to run the code if the data sheet was not in order like it is right now. If the data sheet had "CSIQ" preceding "TERP" or anything not in the order of the array then it would not categorize the data correctly. This mostly has to do with the tickerIndex = tickerIndex + 1 line where the tickerIndex will increase when the next cell is not the same as the current one. With the Original code this would not happen since, for example, ticker(0) would be carried throughout the data sheet. However, the Original code will not display any other values properly either, but the bases would be easier to manipulate in order to make it work. In the end the organization of the data sheet is crucial and both codes are not suited to process through a disorganized one.
