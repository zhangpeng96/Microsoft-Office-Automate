"""
    Change all background color for PowerPoint

    references:
      https://learn.microsoft.com/zh-cn/office/vba/api/powerpoint.slide.background
      https://answers.microsoft.com/en-us/msoffice/forum/all/powerpoint-vba-to-change-slide-color/8a3f82ea-b142-4a92-9119-26adefa24dbe
"""

from win32com import client as wc
from os import path
from os import walk as walk


w = wc.Dispatch('PowerPoint.Application')

def change_background_color(path):
    doc = w.Presentations.Open(path)
    for slide in list(doc.slides):
        slide.FollowMasterBackground = False
        slide.Background.Fill.ForeColor.RGB = 16777215 # Solid White
        change_textframe_color(slide)
    doc.Save()
    print('Convert finished {}'.format(path))
    doc.Close()

def change_textframe_color(slide):
    # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textframe.hastext
    # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.textrange.font
    for shape in slide.Shapes:
        if shape.TextFrame.HasText and shape.TextFrame.TextRange.Font.Color.RGB == 16777215:
            shape.TextFrame.TextRange.Font.Color.RGB = 0


files_list = []
path_input = input('Enter dir: ')
if path.isdir(path_input):
    for _, _, files in walk(path_input):
        for file in files:
            if path.splitext(file)[-1] in ('.ppt', '.pptx'):
                files_list.append( path.join(path_input, file) )

    # print(files_list)
    for file in files_list:
        change_background_color(file)

else:
    print('error')
