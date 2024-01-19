import utils.Utils as utils
from utils.modules import Sheet
import pandas as pd

sample_file = "samples/loadAllSheetsSample/sample1.xlsx"
output_folder = "samples/loadAllSheetsSample/output/"

sample_copy_file = output_folder + utils.get_file_name_from_file_path(sample_file)
sample_copy_file = utils.replace_extension(sample_copy_file, "xlsx")

result_file = output_folder + "result.xlsx"

# utils.quit_excel()

utils.delete_directory(output_folder)
utils.create_directory(output_folder)

# Create a sample copy
utils.create_file_copy(sample_file, sample_copy_file)

# Prepare the sheet
# utils.create_values_only_excel_file(input_file=sample_copy_file, output_file=result_file)

sheet_list = utils.load_all_sheets(excel_file_path=sample_copy_file)

df1 = pd.DataFrame({'Data': ['a', 'b', 'c', 'd']})
df2 = pd.DataFrame({'Data': [1, 2, 3, 4]})
df3 = pd.DataFrame({'Data': [1.1, 1.2, 1.3, 1.4]})

list = [Sheet("week1", df1), Sheet("week2", df2), Sheet("week3", df3)]

sheet_list = sheet_list + list

utils.write_sheets_to_excel(result_file, sheet_list=sheet_list)
