from win32com import client as wc

w = wc.Dispatch('PowerPoint.Application')
# w.Visible = 0
doc = w.Presentations.Open("D:\\math.ppt")
# doc.SaveAs("D:\\math3.pptx")
# f = doc.Slides(1).BackgroundStyle
# f = doc.Slides(5).Background.Fill
# doc.Slides(4).FollowMasterBackground = False
# doc.Slides(4).Background.Fill.ForeColor.RGB = 16777215

for slide in list(doc.slides):
	slide.FollowMasterBackground = False
	slide.Background.Fill.ForeColor.RGB = 16777215
	# print(slide.Background.Fill.ForeColor.RGB)
doc.Save()
# print(
# 	doc.Slides(2).Background.Fill.ForeColor.RGB, 
# 	doc.Slides(1).Background.Fill.ForeColor.RGB, 
# 	doc.Slides(5).Background.Fill.ForeColor.RGB, 
# )
 # = "255,255,255"
# print(f)
# doc.Close()
# doc.Quit()


"""
https://learn.microsoft.com/zh-cn/office/vba/api/powerpoint.slide.background
https://answers.microsoft.com/en-us/msoffice/forum/all/powerpoint-vba-to-change-slide-color/8a3f82ea-b142-4a92-9119-26adefa24dbe
Sub test()
    With ActivePresentation.Slides(2)
    .FollowMasterBackground = False
    .Background.Fill.Solid
        .Background.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
End Sub
"""