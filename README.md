# PPT_MISA_CREATOR

A Script created to scrape the website to prepare a presentation for an indonesian church mass
This script uses tkinter for a GUI, BeautifulSoup and Request to scrape the website and also Python-pptx to prepare the presentation.
<br />
Disclaimer: This script is by no means perfect and has a lot of room for improvement. There is no input checking in this script so user are expected to input properly. <br />
Credits: Imankatolik.or.id as source for the text
<br />
Instructions for non-python users(only work for windows devices) : <br />
1. Download the executable created with pyinstaller from:
https://drive.google.com/file/d/1GNi5spebxDOVB5Snx8_OV2fRVeODQBJ4/view?usp=sharing
2. Extract the zip file
3. Run PPT_misa.exe. The GUI will appear as below:
![image](https://user-images.githubusercontent.com/86009873/168170184-5de7f6df-7dfb-46e7-928a-b552b4c35cdb.png)
4. Insert the wanted date(tanggal), month(bulan) and year(tahun) without any leading 0s(e.g. instead of month 07 for july just write 7) and click confirm
5. The pop up request the user to choose a directory, where they want the powerpoint to be saved at
6. Wait for the script to end and check the directory that you choosed. There should be a new powerpoint there which is called PPT_Misa.pptx  

For Python users:  
Simply make sure that you have all the required libraries and python3(Python2 may not work because of syntax differences and compatibility) and run the script using your favourite IDE. 
I used python 3.9.7 in my case. The libraries required are Beautifulsoup, Requests, Python-pptx and also tkinter.
It is also possible to input the data in the script directly without the GUI by commenting out the tkinter part and changing the variable tanggal,bulan and tahun directly.  

For non-windows user that want to have an executable as well:  
Download the script PPT-Misa.py and default pptx and install all the required modules + pyinstaller.
Open a command prompt and go to the directory with the python script and default powerpoint. Run the command:  
__pyinstaller --add-data "Default.pptx;." PPT-Misa.py__  
When done, this will create a dist folder which contains another PPT-Misa folder. This PPT-Misa folder is what you need. <br />
Note: Do not use the --onefile option since this caused the executable to not work in my case.
