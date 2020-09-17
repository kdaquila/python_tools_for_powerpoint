import util_tools
import pptx_tools
import textbox_tools
import xlsx_tools
import image_tools
import email_tools
import os


# Define paths
pptx_template_path = 'templates/sample_template.pptx'
data_directory = "data"
xls_file_path = 'data/sample_data.xlsx'


# Load the template from the template directory
prs = pptx_tools.load_pptx_file(pptx_template_path)


# Insert current date text box to slide 1
slide_num = 0
text_string = util_tools.get_curr_date_str()
top_inch = 4.4
font_size_pt = 24
font_name = 'Calabri'
textbox_tools.add_textbox_horiz_align_center(prs, slide_num, text_string, top_inch, font_size_pt, font_name)


# Insert excel range image to slide 5
sheet = 'Summary'
cell_range = 'A3:H15'
slide_num = 4
top_inch = 1
left_inch = 0
xlsx_tools.insert_excel_as_img(prs, xls_file_path, sheet, cell_range, slide_num, top_inch, left_inch)


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

