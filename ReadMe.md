a) How to Run
    1) Download the xlsx file explicitly from the desired link & keep it inside your project folder.
    2) Need a Dotnet environment to run
    3) Command -> dotnet run

b) Things I tried for downloading file
    1) WebClient & HttpClient methods for downloading file
    2) Writing a "Get" request to download the json format of the desired link & then make out the CSV out of it.
    3) Curl command was working for the desired link so I tried converting curl command to C# code but that also didn't work.
    4) Lastly I tried writing a python script for dowmloading  file & then runnning that script from C#, I have also provided python script "temp.py" for your refrence.

c) Things that worked
    1) I was able to download the file explicitly & work on further requirements of creating CSV file of data of last 2 years.
    2) This script would work in a genralized way as I have computed the current year & created the CSV data accordingly.
    3) For xlsx to csv conversion I have used Workbook class of Aspose.Cells.
    4) I have also put final output in "output.csv" file.
    
d) Flow of Application
    1) Excel from Website -> CSV File(Same name a file on website) -> Final CSV file(output.csv)
    2) Reasons to create Intermediate CSV file is whenever we want do anything with the data it is to be done with this file in operation rather than excel file which is the original data or final csv file which is the main output.
    
