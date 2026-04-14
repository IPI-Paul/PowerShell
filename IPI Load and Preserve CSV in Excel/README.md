# IPI Load and Preserve CSV in Excel

I have often had times when I wanted to open CSV files in Excel without the Excel magic of removing leading zeros and converting long digit numbers to scientific notation. With a little help from ChatGPT, I made a good start at building a GUI and module to achieve the end result. As I don't have a paid subscription to ChatGPT, I found that using it without its deep reasoning meant that a lot of suggestions just used up a lot of my time only to be told at the end a reason for the failure. Yesterday, after a 12 hour session with ChatGPT and at the point it confirmed that its suggestions could not work, ChatGPT refreshed the page and all the session content vanished before I could save it.

So, I tried Google AI and asked it why ChatGPT was unable to achieve this task, giving it various info about the session. Google AI said that ChatGPT was using Python behind the scenes and could not achieve the logic I was looking for, then gave me the solution to make it all work.

## Problems Faced

- Writing output to Excel line by line would have been unbearable when deailing with CSVs containing 100K+ lines.
- Using ADODB meant that alterations had to be made to files with spaces in their row headers and creating schema.ini files on the fly. It also meant appending numerous text lines under the headers to ensure that ADODBs CopyFromRecordset did not convert data (The 12hr failure).

## Usage

- Add a shortcut to your Windows shell:sendto or any other folder that uses PowerShell.exe -File PathToPs1Script.
- Run it from VS Code, the commandline of PowerShell ISE.
- Use the browse button to select a CSV file and then choose `Load and Preserve CSV contents in Excel` from the Action dropdown menu.

    ![Image of loaded CSV file and functions menu.](./Images/Loaded_CSV.png)

- Wait until the process has completed before working with the Excel file.

    ![Image of Excel output and preserved columns.](./Images/Formatted_Excel_Columns.png)

- You can also choose a ZIP file when browsing and when running you will be prompted to select a CSV file within if there is more than one.

    ![Image of ZIP file content selection.](./Images/Loading_from_ZIP.png)

- Again, Wait until the process has completed before working with the Excel file.

    ![Image of ZIP file loaded csv log.](./Images/Loaded_from_Zip.png)

- To load csv files from web sites, enter the web site base url to search from and run the `Load and Preserve CSV contents in Excel`. You can search from 0 to 5 levels deep by changing the selection in the dropdown menu to the right of the results dropdown menu. The list of csv and zip files will populate the results dropdown menu to the right of the file name filter.

    ![Image of Web Search.](./Images/Found_in_web_search.png)

- Use +/plus signs in the filter by text box to filter files by file names matching the entries. Once the drop down menu has been populated, select the file you want to open in excel and run the `Load and Preserve CSV contents in Excel` function again. This will download the selected file to your temp folder and load it from there, deleting it once complete.

    ![Image of selected web search results file loaded.](./Images/Loaded_web_search_csv.png)

- You can also choose to download the selected file from the web search or download all files displayed in the results dropdown menu. For this, you will need to enter the path to the folder location you want the files saved in or click the save to folder button next to the text box and select the folder.
