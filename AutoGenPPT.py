'''
from pptx import Presentation
import os
 
prs = Presentation()
slide = prs.slides.add_slide(prs.slide_layouts[8])
placeholder = slide.placeholders[1]
picture = placeholder.insert_picture('GTA5_0.jpg')
prs.save("ESEMPIO.pptx")
os.startfile("ESEMPIO.pptx")
'''

##
##  Copy & Paste Tool for images to PowerPoint(.pptx)
##
import pptx
import pptx.util
import glob
import imageio

OUTPUT_TAG = "BBB"

# new
prs = pptx.Presentation()

path = "D:\BBB\Grade8Autumn\HighLevelCourses\Module1_HowToLearnEnglish_Volcabulary"
# open
# prs_exists = pptx.Presentation("some_presentation.pptx")

# default slide width
#prs.slide_width = 9144000
# slide height @ 4:3
#prs.slide_height = 6858000
# slide height @ 16:9
#prs.slide_height = 5143500

# title slide
slide = prs.slides.add_slide(prs.slide_layouts[0])
# blank slide
#slide = prs.slides.add_slide(prs.slide_layouts[6])

# set title
title = slide.shapes.title
title.text = OUTPUT_TAG

#pic_left  = int(prs.slide_width * 0.15)
#pic_top   = int(prs.slide_height * 0.1)
pic_left  = int(prs.slide_width * 0.0)
pic_top   = int(prs.slide_height * 0.1)
pic_width = int(prs.slide_width * 1.0)


bmp_file = path + "\*.bmp"
for g in glob.glob(bmp_file):
    pic_left  = int(prs.slide_width * 0.0)
    pic_top   = int(prs.slide_height * 0.1)
    pic_width = int(prs.slide_width * 1.0)
    print(g)
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    #tb = slide.shapes.add_textbox(0, 0, prs.slide_width, pic_top / 2)
    #p = tb.text_frame.add_paragraph()
    #p.text = g
    #p.font.size = pptx.util.Pt(14)

    img = imageio.imread(g)
    if(img.shape[0] < img.shape[1]):
        pic_height = int(pic_width * img.shape[0] / img.shape[1])
    else:
        pic_height = int(prs.slide_height * 0.8)
        pic_width = int(pic_height * img.shape[1] / img.shape[0])
    #pic   = slide.shapes.add_picture(g, pic_left, pic_top)
    pic   = slide.shapes.add_picture(g, pic_left, pic_top, pic_width, pic_height)

prs.save("%s.pptx" % OUTPUT_TAG)