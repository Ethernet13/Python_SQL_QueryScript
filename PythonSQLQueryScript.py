import pyodbc
import pandas as pd
import os
from datetime import datetime
from plyer import notification

# create sql connection
connection = pyodbc.connect(driver ='{ODBC Driver 17 for SQL Server}', host = 'Add server',
database = "SiteDelivery",trusted_connection = 'yes')


#SQL command to read the data
sqlQuery = "select * FROM ADD TABLE where ADD TABLE IN ('Undecided') and status = 'Active'and FIELD REQUIRED between 'DATE' and 'DATE'"

#Getting the data from sql into pandas dataframe
df = pd.read_sql(sql = sqlQuery, con = connection)

#Export the data on the desktop
df.to_csv(os.environ["userprofile"] + "\\Desktop\\PythonScript\\" + "SQL_Undecided_" + datetime.now().strftime("%d-%b-%Y %H%M%S") + ".csv",index = False)

#Display notification to user
notification.notify(title = "Report Status!!!",
message = f"Undecided datafill has been successfully saved into Excel.\
        \nTotal Rows: {df.shape[0]}\nTotal Columns: {df.shape[1]}",
        timeout = 10)