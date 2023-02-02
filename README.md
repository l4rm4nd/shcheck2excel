# shcheck2excel 
 Python script that takes shcheck json and brings the results to Excel 
  
 This script will parse json results generated with shcheck and convert them into an Excel file. 
  
 The results can be generated with the following command: 
 ```bash 
 shcheck.py -d -g -j -k --hfile webservices.txt > header_check_new.json 
 ``` 
  
 ## Installation 
 The script requires the python modules mentioned in requirements.txt. These can be installed with the following command: 
 ```bash 
 pip install -r requirements.txt 
 ``` 
  
 ## Usage 
 The script can be started with the following command: 
 ```bash 
 python shcheck2excel -i shcheck_results.json 
 ``` 
  
 The output file will be an Excel file showing the results of the audit. 
 "# shcheck2excel" 
