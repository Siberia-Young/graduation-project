import os
import shutil
import time
import json
import datetime

from tools import merge
from module import code_1_crawl_basic_product_data
from module import code_2_filter_by_repeat
from module import code_3_filter_by_whitelist
from module import code_4_crawl_and_save_product_images
from module import code_5_extract_image_text
from module import code_6_filter_by_image_text
from module import code_8_classify_and_sort
from module import code_9_filter_by_sales
from module import code_11_cell_style_adjustments

platform_name = "tb"

source_folder = f"data/{platform_name}"
destination_folder = f"data/{platform_name}/merge"
outcome_folder = f"data/{platform_name}/merge/outcome"
final_folder = f"data/z_submit"

# ready_or_not = 'Y'
ready_or_not = input(f'请确保已经完成以下准备：\n1.关闭VPN；\n2.删除data/{platform_name}文件夹内所有的xlsx文件；\n3.删除data/{platform_name}文件夹下merge文件夹及其内部所有文件；\n4.在模拟浏览器上登录淘宝；\n5.保持模拟浏览器最大化并且无遮挡；\n【Y/N】：')

total_time_list = []

def my_print(num):
    str1 = num < 10 and '0' + str(num) or str(num)
    print(f'\n----------------{str1}----------------')

def time_lapse(start_time):
    end_time = time.time()
    total_time_list.append(end_time-start_time)
    print(f"\n耗时：{(end_time-start_time)/60:.2f} min")
    return end_time

if ready_or_not == 'Y':
    # 读取parameter.json的参数
    with open('src/parameter.json', encoding='utf-8') as file:
        parameter = json.load(file)
        keywords = parameter['keywords']
        relatedwords = parameter['relatedwords']
        graphicwords = parameter['graphicwords']
        print(f"\n当前参数：{parameter}")
    
    start_time = time.time()

    my_print(1)
    code_1_crawl_basic_product_data.crawl_basic_product_data(keywords)
    start_time = time_lapse(start_time)
    
    # 将爬取的excel表复制一份到merge文件夹下面
    if not os.path.exists(destination_folder):
        os.makedirs(destination_folder)
    for filename in os.listdir(source_folder):
        if filename.endswith(".xlsx"):
            source_file = os.path.join(source_folder, filename).replace(os.sep, '/')
            destination_file = os.path.join(destination_folder, filename).replace(os.sep, '/')
            shutil.copy(source_file, destination_file)

    # 将merge文件夹下面的所有excel文件统一命名
    index = 0
    for filename in os.listdir(destination_folder):
        if filename.endswith(".xlsx"):
            index += 1
            file_path = os.path.join(destination_folder, filename).replace(os.sep, '/')
            new_filename = f"{platform_name} ({index}).xlsx"
            new_file_path = os.path.join(destination_folder, new_filename).replace(os.sep, '/')
            shutil.move(file_path, new_file_path)

    merge.merge(os.path.join(destination_folder, "merge.xlsx").replace(os.sep, '/'), destination_folder)

    my_print(2)
    code_2_filter_by_repeat.filter_by_repeat(os.path.join(destination_folder, "merge.xlsx").replace(os.sep, '/'))
    start_time = time_lapse(start_time)

    my_print(3)
    code_3_filter_by_whitelist.filter_by_whitelist(relatedwords[0], relatedwords[1], relatedwords[2], os.path.join(destination_folder, "merge_2.xlsx").replace(os.sep, '/'))
    start_time = time_lapse(start_time)

    my_print(8)
    code_8_classify_and_sort.classify_and_sort(os.path.join(destination_folder, "merge_2_3.xlsx").replace(os.sep, '/'))
    start_time = time_lapse(start_time)

    my_print(9)
    code_9_filter_by_sales.filter_by_sales(os.path.join(destination_folder, "merge_2_3_8.xlsx").replace(os.sep, '/'))
    start_time = time_lapse(start_time)

    my_print(4)
    code_4_crawl_and_save_product_images.crawl_and_save_product_images(os.path.join(destination_folder, "merge_2_3_8_9.xlsx").replace(os.sep, '/'))
    start_time = time_lapse(start_time)

    my_print(5)
    code_5_extract_image_text.extract_image_text(os.path.join(destination_folder, "merge_2_3_8_9.xlsx").replace(os.sep, '/'))
    start_time = time_lapse(start_time)

    my_print(6)
    code_6_filter_by_image_text.filter_by_image_text(graphicwords, os.path.join(destination_folder, "merge_2_3_8_9.xlsx").replace(os.sep, '/'))
    start_time = time_lapse(start_time)

    my_print(9)
    code_9_filter_by_sales.filter_by_sales(os.path.join(destination_folder, "merge_2_3_8_9_6.xlsx").replace(os.sep, '/'))
    start_time = time_lapse(start_time)

    # 将目标excel表复制一份到outcome文件夹下面
    if not os.path.exists(outcome_folder):
        os.makedirs(outcome_folder)
    file_list = []
    for filename in os.listdir(destination_folder):
        # 找到以merge开头的excel文件
        if filename.endswith(".xlsx") and filename.startswith("merge"):
            source_file = os.path.join(
                destination_folder, filename).replace(os.sep, '/')
            file_list.append(source_file)

    target_list = ['merge.xlsx', 'merge_2.xlsx', 'merge_2_3.xlsx',
                   'merge_2_3_8_9(副本).xlsx', 'merge_2_3_8_9_6_9.xlsx', 'merge_2_3_8_9(副本).xlsx']
    num = 1
    for filename in target_list:
        for path in file_list:
            if path.split('/')[-1] == filename:
                destination_file = os.path.join(outcome_folder, '文件'+str(num)+'.xlsx').replace(os.sep, '/')
                shutil.copy(path, destination_file)
                num += 1

    my_print(11)
    code_11_cell_style_adjustments.cell_style_adjustments(outcome_folder)
    start_time = time_lapse(start_time)

    for filename in os.listdir(source_folder):
        if filename.startswith("文件与流程关系图"):
            from_path = os.path.join(source_folder, filename).replace(os.sep, '/')
            to_path = os.path.join(outcome_folder, filename).replace(os.sep, '/')
            shutil.copy(from_path, to_path)

    current_time = datetime.datetime.now()
    time_string = current_time.strftime("淘宝_%Y-%m-%d_%H%M%S")
    shutil.copytree(outcome_folder,os.path.join(final_folder, time_string).replace(os.sep, '/'))

    print(f"\n各阶段耗时：{total_time_list}")
    print(f"\n总耗时：{sum(total_time_list)/60:.2f} min")