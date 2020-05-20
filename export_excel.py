import xlsxwriter
import requests
import workbook_format
from datetime import datetime
from pathlib import Path
from io import BytesIO
from PIL import Image


def create_directory(subdir):
    current_directory = Path.cwd()
    folder = str(current_directory) + subdir
    Path(folder).mkdir(parents=True, exist_ok=True)

    return folder


def resize_logo(url):
    with Image.open(BytesIO(requests.get(url).content)) as img:
        width_100 = img.width
        height_100 = img.height

    width_60 = 80
    img = Image.open(BytesIO(requests.get(url).content))
    wpercent = (width_60/float(width_100))
    hsize = int((float(height_100)*float(wpercent)))
    img = img.resize((width_60, hsize), Image.ANTIALIAS)
    image_folder = create_directory('/image/logo')
    image_path = image_folder + '/logo_60.png'
    img.save(image_path)

    return image_path


def export_excel(report_data):
    # Create file path
    report_folder = create_directory('/report/excel')
    file_path = report_folder + '/report' + \
        datetime.now().strftime("%d%m%y%H%M%S%f") + '.xlsx'

    # Create workbook
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet()

    # Add logo
    if report_data['logo_url']:
        image_path = resize_logo(report_data['logo_url'])
        worksheet.insert_image('A1', image_path)

    # Report title
    report_title_format = workbook.add_format(workbook_format.report_title)
    worksheet.merge_range(
        0, 0, 0, 2, report_data['title'].upper(), report_title_format)

    # Created at
    report_summary_format = workbook.add_format(workbook_format.report_summary)
    created_at = datetime.now().strftime("%d/%m/%y %H:%M")
    worksheet.write_string(2, 2, 'Thời điểm tạo: {}'.format(
        created_at), report_summary_format)

    # Range time
    if report_data['range_time']:
        worksheet.write_string(
            3, 2, 'Khoảng thời gian: {}'.format(report_data['range_time']), report_summary_format)

    # Total
    if report_data['total']:
        worksheet.write_string(4, 2, 'Tổng số: {}'.format(
            report_data['total']), report_summary_format)

    # Table
    if report_data['tables']:
        start_row = 8
        start_col = 0
        max_col = 0

        table_title_format = workbook.add_format(workbook_format.table_title)
        table_summary_format = workbook.add_format(
            workbook_format.table_summary)
        table_label_format = workbook.add_format(workbook_format.table_label)
        table_item_format = workbook.add_format(workbook_format.table_item)

        for index, sub_table in enumerate(report_data['tables']):
            sub_table_col_num = len(sub_table['labels'])
            sub_table_start_row = start_row + index * \
                (len(report_data['tables'][index - 1]['datas']) + 2)

            # Set max col
            if sub_table_col_num > max_col:
                max_col = sub_table_col_num

            # Table title
            worksheet.merge_range(sub_table_start_row, start_col,
                                  sub_table_start_row, start_col + sub_table_col_num - 2, sub_table['title'].upper(), table_title_format)
            # Table summary
            worksheet.write_string(
                sub_table_start_row, start_col + sub_table_col_num - 1, 'Tổng số: {}'.format(sub_table['total']), table_summary_format)
            # Table label
            for index, label in enumerate(sub_table['labels']):
                worksheet.write_string(
                    sub_table_start_row + 1, start_col + index, label, table_label_format)
            # Table content
            for data_index, data in enumerate(sub_table['datas']):
                for item_index, item in enumerate(data):
                    worksheet.write_string(
                        sub_table_start_row + 2 + data_index, start_col + item_index, item, table_item_format)

            worksheet.set_column(0, max_col - 1, 40)

    workbook.close()

    return file_path
