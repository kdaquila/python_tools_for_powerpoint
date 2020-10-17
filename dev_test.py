import example_application
import os

img_directory = "img"
data_directory = "data"
xls_file_name = "Astronomy_Data.xlsx"

pptx_template_path_dph = os.path.join('template', 'template.pptx')
example_application.build_from_template(img_directory,
                                        data_directory,
                                        xls_file_name,
                                        pptx_template_path_dph)
