import datetime
import os
import re
from python_tools_for_powerpoint import pptx_tools, textbox_tools, xlsx_tools, image_tools


def generate_image_slide_notes(filename):
    if re.search("Image1", filename):
        today = datetime.datetime.now()
        slide_notes = [
            [{'text': filename}],
            [{'text': ""}],
            [{'text': "A bold heading", 'bold': True}],
            [{'text': "Sample Text"}],
            [{'text': "Range of the last 10 days ({1} – {0})".format(
                (today - datetime.timedelta(days=1)).strftime('%m/%d'),
                (today - datetime.timedelta(days=10)).strftime('%m/%d'))}],
            [{'text': ""}],
            [{'text': "Another bold heading", 'bold': True}],
            [{'text': "Sample Text"}],
            [{'text': "Range of the last 10 days ({1} – {0})".format(
                (today - datetime.timedelta(days=1)).strftime('%m/%d'),
                (today - datetime.timedelta(days=10)).strftime('%m/%d'))}],
        ]
        return slide_notes

    elif re.search("Image[2-3]+", filename):
        slide_notes_text = filename + "\n\n"
        slide_notes_text += "More Sample Text\n"
        today = datetime.datetime.now()
        slide_notes_text += "Range of the last 10 days ({1} – {0})\n".format(
            (today - datetime.timedelta(days=1)).strftime('%m/%d'),
            (today - datetime.timedelta(days=10)).strftime('%m/%d'))
        return slide_notes_text
    else:
        return None


def cond_format_rule1(value_str):
    value_num = float(value_str)
    standard = 12756
    if (value_num - standard) / standard * 100 > 300:
        return 'FF0000'
    elif (value_num - standard) / standard * 100 > 200:
        return '0070C0'
    else:
        return None


def cond_format_rule2(value_str):
    value_num = float(value_str)
    standard = 9.8
    if (value_num - standard) / standard * 100 > 100:
        return 'FF0000'
    elif (value_num - standard) / standard * 100 > 10:
        return '0070C0'
    else:
        return None


def cond_format_rule3(value_str):
    value_num = float(value_str)
    standard = 24
    if (value_num - standard) / standard * 100 > 90:
        return 'FF0000'
    elif (value_num - standard) / standard * 100 > 50:
        return '0070C0'
    else:
        return None


def cond_format_rule4(value_str):
    value_num = float(value_str)
    standard = 365
    if (value_num - standard) / standard * 100 > 90:
        return 'FF0000'
    elif (value_num - standard) / standard * 100 > 50:
        return '0070C0'
    else:
        return None


def build_from_template(img_directory, data_directory, xls_file_name, pptx_template_path):
    # Define paths
    xls_file_path = os.path.join(data_directory, xls_file_name)

    # Load the template from the template directory
    prs = pptx_tools.load_pptx_file(pptx_template_path)

    # Insert current date text box to slide 1
    curr_date_str = datetime.datetime.now().strftime('%m/%d/%y')
    textbox_tools.set_textbox_text(curr_date_str, prs, slide_num=0, shape_id=2)

    # Set title text box on slide 3
    textbox_tools.set_textbox_text("Astronomy Summary – " + curr_date_str, prs, slide_num=2, shape_id=3)

    # Open the Excel File
    wb = xlsx_tools.open_excel_file(xls_file_path)

    # Insert excel range image to slide 3
    xlsx_tools.copy_range_to_table(prs, wb, sheet='Summary', range_str="B3:F11", slide_num=2, shape_id=1)

    # Define Conditional Formatting
    cond_format_rules = [
        {"column": 1, "func": cond_format_rule1},
        {"column": 2, "func": cond_format_rule2},
        {"column": 3, "func": cond_format_rule3},
        {"column": 4, "func": cond_format_rule4},
    ]

    # Apply Conditional Formatting
    pptx_tools.apply_conditional_formatting_to_columns(prs, slide_num=2, shape_id=1, col_rules=cond_format_rules)

    # Append multiple image slides
    # Images slides are ordered by image creation date
    # Image slide notes depend on filename-based rule
    img_extensions = ['.png', '.jpg', '.tif']
    image_paths = image_tools.find_images(img_directory, img_extensions)
    image_paths = image_tools.sort_files_by_creation_date(image_paths)
    image_tools.add_multiple_image_slide_from_list_with_notes(prs, image_paths, generate_image_slide_notes)

    # Build the powerpoint's final path
    pptx_name = 'Astronomy_SlideDeck_{}_Draft.pptx'.format(
        datetime.datetime.now().strftime('%m.%d.%y'))
    pptx_file_path_final = os.path.join(img_directory, pptx_name)

    # Save the powerpoint to the data directory
    prs.save(pptx_file_path_final)
    print("\nSaved file to: '{0}'".format(pptx_file_path_final))



