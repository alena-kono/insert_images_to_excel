import pathlib
import tkinter
import tkinter.filedialog

import gui

tkinter.Tk().withdraw()

PATH_TO_IMAGES_DIRECTORY = pathlib.Path(
    tkinter.filedialog.askdirectory(
        title='Choose directory with images',
        initialdir=pathlib.os.getcwd()))

PATH_TO_TARGET_FILE = pathlib.Path(
    tkinter.filedialog.askopenfilename(
        title='Choose target file',
        initialdir=pathlib.os.getcwd()))

PATH_TO_UPDATED_FILE = pathlib.Path(
    PATH_TO_TARGET_FILE.with_name(
        PATH_TO_TARGET_FILE.stem + '_UPD.xlsx'))

CELL_WIDTH = 15
CELL_HEIGHT = 85

TARGET_COLUMN = 'A'

LOOKUP_COLUMN = gui.create_simple_dialog(
    title='Setup',
    prompt='Enter column letter for lookup: ').upper()

START_ROW = int(gui.create_simple_dialog(
    title='Setup',
    prompt='Enter start row: '))

# # for testing only

# PATH_TO_IMAGES_DIRECTORY = pathlib.Path(
# '/Users/alena/Dev/insert_images_to_excel/test_data/images')
# PATH_TO_TARGET_FILE = pathlib.Path(
# '/Users/alena/Dev/insert_images_to_excel/test_data/test_target_file.xlsx')
# PATH_TO_UPDATED_FILE = pathlib.Path(
# '/Users/alena/Dev/insert_images_to_excel/test_data/test_target_file_UPD.xlsx')
# TARGET_COLUMN = 'A'
# LOOKUP_COLUMN = 'B'
