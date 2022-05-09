Instructions:

1. Clone this repository and `cd` into it in a terminal
2. Create a virtual environment: `python -m venv env`
3. Activate your virtual environment: `source env/bin/activate`
4. Run `pip install -r requirements.txt`
5. Run `python data_dictionary_diff.py {old_file_path} {new_file_path}` (You should be able to drag the file paths from finder into the terminal window)
   Ex: `python data_dictionary_diff.py COVID_DDict_4-20_1215a*\(2\).xlsx COVID_DDict_4-22_1215a.xlsx`

Note: The new file will be created in this folder with the default filename DataDictionary*{today's date}.xlsx. If you would like to choose your own filename, you can add that name (without the file extension) as an argument after the two file paths. Ex: "python3 diff.py COVID_DDict_4-20_1215a*\(2\).xlsx COVID_DDict_4-22_1215a.xlsx NewReconciledDataDict"

Note: If your filenames or your desired new file name has spaces in it (not recommended), you will have to surround them with quotes when calling the function.
