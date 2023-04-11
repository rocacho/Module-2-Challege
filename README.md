# Module-2-Challege
Challenge Files for Module 2:
1) Module 2 Challenge code

  1.1 To Aggregate ticker to create summary report forTickers & Total Stock Volume y used a method published by Karen Tateosyan tutorial "2.22 Agreggating data with Excel VBA": https://www.youtube.com/watch?v=nSWoLqLua-s

IMPORTANT NOTE: To be able to use this method I had to creat a new tab in the excel file called "Report" to have the Tickers      without duplicates. When I ran this code within each year tab the code erased all duplicated data from the original data set

  1.2 To retrive informantion for OpenPrice & ClosePrice I used the Index&Macth formula where Index OpenPrice is  & Match is TickerRng. This method was also based on one publication from Karen Tateosyan tutorial "2.19 INDEX and MATCH Functions in Another Sheet with VBA": https://www.youtube.com/watch?v=Rmy5-PC7QKY. With this I was able to calculate the Yearly Change and the Percent Change 
   
   1.3 To retirve the information To calculate Yearly Change and Percent Change I used nested forumlas to retrive the information of the first open price of the year, the last closeing price of that year and simple operators to substract and to calculate percent variation
   
   1.4 In order to generate the Conditional Formating for Yearly Change I used a method punlished by EXCEL DESTINATION in a Tutorial called "Conditional Formatting usingVBA Code": https://www.youtube.com/watch?v=F29G18GdTAQ&t=343s
   
   1.5 The creatrion of the summary for Greates Increase, Decrease and volume was done based on a combination of techniques learned in class and also leveraging the solutions used in point 1.1 
   
   1.6 To make the code run for the other years I just have to chang the sheet in one line of the code: Ln10, Col 5
 Set Aws = ThisWorkbook.Worksheets("2018") where you can replace "2018" for any of the other years and you can get the summarized data

2) Screenshots of the results:
 2.1 File "Module_2_Challemge_2018.PNG" shows the Results for 2018 data 
 2.2 File "Module_2_Challemge_2019.PNG" shows the Results for 2019 data  
 2.3 File "Module_2_Challemge_2020.PNG" shows the Results for 2020 data 
