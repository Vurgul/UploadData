# Project UploadData
### Description
Based on the script, an executable file was created, 
which is launched when you click the "Load" button in the Excel file
with updated data and loads the same data into the database.

#### Implementation principle
1. Created a table in the database to load data
2. Data loading script written
3. An executable file has been created
4. The macro of the download button is configured in the Excel file

### Core technologies
* Python
* PostgreSQL
### Installation
1. Clone repository

2. Install dependencies from a file requirements.txt
```
pip install -r requirements.txt
``` 
3. The upload.exe executable file must be located in the 
   same directory as the EXCEL file.
   
4. As input, upload.exe takes the path to
   the EXCEL file and the password from the database.
   
### Author
Vurgul: https://github.com/Vurgul