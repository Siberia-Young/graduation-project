import os
import shutil

from tools import merge
# import code_1_crawl_basic_product_data
import code_2_filter_by_repeat
import code_3_filter_by_whitelist
# import code_4_crawl_and_save_product_images
# import code_5_extract_image_text
# import code_6_filter_by_image_text
import code_8_classify_and_sort
import code_9_filter_by_sales
# import code_10_crawl_detailed_product_data
# import code_11_cell_style_adjustments

platform_name = "1688"

source_folder = f"data/{platform_name}"
destination_folder = f"data/{platform_name}/merge"
outcome_folder = f"data/{platform_name}/merge/outcome"

# ready_or_not = 'Y'
ready_or_not = input('请确保清空merge文件夹内所有文件以及merge文件夹同级的所有excel文件【Y/N】：')

def my_print(num):
    str1 = num < 10 and '0'+ str(num) or str(num)
    print(f'\n----------------{str1}----------------')

if ready_or_not == 'Y':
    # 将爬取的excel表复制一份到merge文件夹下面
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    for filename in os.listdir(source_folder):
        if filename.endswith(".xlsx"):
            source_file = os.path.join(source_folder, filename)
            destination_file = os.path.join(destination_folder, filename)
            shutil.copy(source_file, destination_file)
    
    # 将merge文件夹下面的所有excel文件统一命名
    index = 0
    for filename in os.listdir(destination_folder):
        if filename.endswith(".xlsx"):
            index += 1
            file_path = os.path.join(destination_folder, filename)
            new_filename = f"{platform_name} ({index}).xlsx"
            new_file_path = os.path.join(destination_folder, new_filename)
            shutil.move(file_path, new_file_path)
    

    merge.merge(os.path.join(destination_folder, "merge.xlsx"), destination_folder)

    my_print(2)
    code_2_filter_by_repeat.filter_by_repeat(os.path.join(destination_folder, "merge.xlsx"))

    my_print(3)
    code_3_filter_by_whitelist.filter_by_whitelist(['适用'], ['huawei','华为'], os.path.join(destination_folder, "merge_2.xlsx"))

    my_print(8)
    code_8_classify_and_sort.classify_and_sort(os.path.join(destination_folder, "merge_2_3.xlsx"))

    my_print(9)
    code_9_filter_by_sales.filter_by_sales(os.path.join(destination_folder, "merge_2_3_8.xlsx"))