import utils.Utils as utils

sample_file_1 = "samples/mergeFilesSample/sample1.xlsx"
sample_file_2 = "samples/mergeFilesSample/sample2.xlsx"
sample_file_3 = "samples/mergeFilesSample/sample3.xlsx"
output_folder = "samples/mergeFilesSample/output/"

sample_copy_file_1 = output_folder + utils.get_file_name_from_file_path(sample_file_1)
sample_copy_file_2 = output_folder + utils.get_file_name_from_file_path(sample_file_2)
sample_copy_file_3 = output_folder + utils.get_file_name_from_file_path(sample_file_3)

sample_copy_file_1 = utils.replace_extension(sample_copy_file_1, "xlsx")
sample_copy_file_2 = utils.replace_extension(sample_copy_file_2, "xlsx")
sample_copy_file_3 = utils.replace_extension(sample_copy_file_3, "xlsx")

result_file = output_folder + "result.xlsm"

# utils.quit_excel()

utils.delete_directory(output_folder)
utils.create_directory(output_folder)

# Create a sample copy
utils.create_file_copy(sample_file_1, sample_copy_file_1)
utils.create_file_copy(sample_file_2, sample_copy_file_2)
utils.create_file_copy(sample_file_3, sample_copy_file_3)

file_names = [sample_copy_file_1, sample_copy_file_2, sample_copy_file_3]

utils.merge_multiple_excels_to_one_excel(file_names, result_file)
