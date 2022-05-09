Instructions:

1. Open up a terminal window
2. Navigate into this folder with the "cd" command.
   Ex: If it is located in your "Downloads" folder -> "cd Downloads/Python_Excel_Diff" and press enter.
3. Run "pip3 install -r requirements.txt"
4. Run "python3 diff.py {old*file_path} {new_file_path} (You should be able to drag the file paths from finder into the terminal window)
   Ex: "python3 diff.py COVID_DDict_4-20_1215a*\(2\).xlsx COVID_DDict_4-22_1215a.xlsx"

Note: The new file will be created in this folder with the default filename DataDictionary*{today's date}.xlsx. If you would like to choose your own filename, you can add that name (without the file extension) as an argument after the two file paths. Ex: "python3 diff.py COVID_DDict_4-20_1215a*\(2\).xlsx COVID_DDict_4-22_1215a.xlsx NewReconciledDataDict"

Note: If your filenames or your desired new file name has spaces in it (not recommended), you will have to surround them with quotes when calling the function.
