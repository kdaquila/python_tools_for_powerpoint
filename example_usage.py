import util_tools
import pptx_tools
import textbox_tools
import xlsx_tools
import image_tools
import email_tools
import os
import datetime

# Define paths
pptx_template_path = 'templates/sample_template.pptx'
data_directory = "data"
xls_file_path = 'data/sample_data.xlsx'

# Load the template from the template directory
prs = pptx_tools.load_pptx_file(pptx_template_path)

# Insert current date text box to slide 1
slide_num = 0
text_string = util_tools.get_curr_date_str()
table_top = 4.4
font_size_pt = 24
font_name = 'Calabri'
textbox_tools.add_textbox_horiz_align_center(prs, slide_num, text_string, table_top, font_size_pt, font_name)

# Insert excel range image to slide 5
sheet = 'Summary'
cell_range = 'A3:H15'
col_start = 'A'
col_stop = 'H'
row_start = 3
row_stop = 14
slide_num = 4
table_top = 1
table_left = 0
table_width = 13.333
table_height = 5
font_size_factor = 0.7
xlsx_tools.insert_excel_range(prs, xls_file_path, sheet, col_start, col_stop, row_start, row_stop,
                              slide_num, table_top, table_left, table_width, table_height, font_size_factor)

# Insert notes to slide 5
slide_num = 4
today = datetime.datetime.now()
slide_notes_text = '''
Today is {0}
'''.format(today.strftime('%m/%d'))
pptx_tools.insert_slide_notes(prs, slide_num, slide_notes_text)

# Append multiple image slides
img_extensions = ['.png', '.jpg', '.tif']
image_tools.add_multiple_image_slides(prs, data_directory, img_extensions)

# Build the powerpoint's final path
today_date = util_tools.get_curr_date_str()
pptx_name = 'sample_presentation_{}.pptx'.format(today_date)
pptx_file_path_final = os.path.join(data_directory, pptx_name)

# Save the powerpoint to the data directory
prs.save(pptx_file_path_final)
print("\nSaved file to: '{0}'".format(pptx_file_path_final))

# Email all files in the data directory
from_addr = "host@example.com"
to_addr = "host@example.com"
subject = "subject"
body = "body"
smtp_addr = "hostname"
email_tools.email_all_files(data_directory, from_addr, to_addr, subject, body, smtp_addr)
