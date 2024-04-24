"# python_excel_auto" 

This guide will show you how to use
 - python to read a .xlsx files.
 - python to analyse data.
 - python to write data to new files.
 - "Task Schedule" to run scripts automatically.
 - "micorsoft flow" to send emails anytime a file is created.

Bruke python til å oppdatere excel 
For å følge denne guiden bør du ha kjennskap til å skrive python. For en guide til python for windows se her :
https://learn.microsoft.com/en-us/windows/python/beginners

Først følger litt enkel python kode for å lese et excelark, gjøre en enkel analyse og skrive analysen til et nytt excel ark.
You can find the exemple excel files here 
https://github.com/joarvevle/python_excel_auto

```

# import libraries
import os
import pandas as pd

# set filename for where to import from and to
fileNameImport = "C:\\Users\\"+ os.getlogin()+"\\OneDrive\\myProject\\excelExempleFile_1.xlsx"
fileNameWriteTo = "C:\\Users\\"+ os.getlogin()+"\\OneDrive\\myProject\\excelExempleFile_2.xlsx"

# load data into a dataframe, df
df = pd.concat(pd.read_excel(fileNameImport, sheet_name=None), sort=False)  

"""
make changes to data. Lots of possibilities, and what I do here is just a sample of what is possible
first I create a new column with margin (%). Then I filter out Everything from category "B".
Then I find the item with highest margin in each category. 
dive into the the full docs here, https://pandas.pydata.org/docs/getting_started/index.html or just google it
"""

df['Margin'] = (df['PriceOut'] / df['PriceIn'])-1  # create new cloumn "Margin"
df = df.loc[df['Category1'] !="B"] # filter
df = df.sort_values('Margin', ascending=False).drop_duplicates(['Category1']).sort_values("Category1") 

# there are countless ways to do every task in python. Doing sort_values and drop_duplicates is an efficent way, but 
# group_by() and max() could be more intuitive. Especially if coming from R/tidyverse

# write data to new excelFile, and maybe also to a html file to convenienly write it to emails later
df.to_excel(fileNameWriteTo, index=None)
df.to_html(os.path.splitext(fileNameWriteTo)[0]+'.html', index=None)

```

If you put the exemple files in the folder C:\Users\vevljoa\OneDrive\myProject it will automatically read file_1 arrange and filter and write to a new excel file and a new html file.  

***

# how to start the script with Task Scheduler
Press window-key and type "Task Schedule". 
click the "Create task"  

![image](https://github.com/joarvevle/python_excel_auto/assets/143795683/2dc84c8f-f82f-4d20-b44c-ea5d9ff068b2)

give your task a new name :  

![image](https://github.com/joarvevle/python_excel_auto/assets/143795683/05286c46-4618-4484-a6f9-063794a017ce)


find the "Action" task and add new  

![image](https://github.com/joarvevle/python_excel_auto/assets/143795683/3a1c6aa1-573d-445f-b4b6-4ceefc5712a6)

Best way to have "Task Schedule" run python scripts is to run Python.exe with your script as an argument, and "Start in (optional)" as the folder where you have stored your script. If you dont know the location of your python.exe you could run the following code in a python enviroment :

```
print(sys.executable)
```

Your new action should now looke like this:  

![image](https://github.com/joarvevle/python_excel_auto/assets/143795683/91832ec7-d242-4ff7-ae4e-58080f990fce)

You can now set up triggers to have it run on a schedule. You can also press "run" to run it imidiately















