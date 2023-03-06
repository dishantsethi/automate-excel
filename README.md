# POC Repo for CA Deepak
# Excel automation

## Set up
1. Install Python3.9.7
2. Install pip 22.3.1
3. Install git 2.33.0
4. git clone
5. pip install -r requirements.txt
6. os.makedirs("excel_input_folder", 0o777)
7. os.makedirs("excel_output_folder", 0o777)
8. os.makedirs("pdf_output_folder", 0o777)
9. python main.py

## Assumptions
1. "D" column of "Load Book Movement" will have string "Total" as end of file
2. "B" column of "Collections & Overdues" will have string "Total" as end of file
3. "C2" cell of "Summary" will have datetime.
4. File name format will be "{} {} {bankname} - {Tranche} {Number} {}.xlsx"
