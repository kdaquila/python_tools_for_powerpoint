# python_tools_for_powerpoint
Tools for editing PowerPoint slides with python

# Usage
Import python_tools_for_powerpoint and then use it in your own script. 
The basic workflow involves opening an existing .pptx file, adding items to the slides, then saving the edited file.
Check the example_application.py file for sample usage. Run the script dev_test.py for a working demo.

# Installation
Note, this library has not been packaged yet for Pip or Conda, so you will need to manually download and 
place the code within your project. The library depends on python-pptx. You will also need pyopenxl for 
the Excel features.

Here are some sample instructions for getting started using Conda:
1. Install an anaconda distribution (e.g. [Miniconda](https://docs.conda.io/en/latest/miniconda.html)) 
2. Create and activate a new conda virtual environment
    ```
    conda create -n powerpoint_env python=3
    activate powerpoint_env
    ```
3. Install the dependencies
    ```
    conda install python-pptx openpyxl
    ```
4. Get the source code
    ```
    git clone https://github.com/kdaquila/python_tools_for_powerpoint.git
    cd python_tools_for_powerpoint
    ```
5. Run the demo
    ```
    python dev_test
    ```
6. The generated PointPoint file will be located in the 'img' folder

