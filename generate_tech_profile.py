import comtypes.client
import os
import io
from pptx import Presentation
import pptx.shapes
from pptx.util import Inches
import component_data

def load_image_data(slide, imageNames):
    # Function that saves the size and position data of image in a dictionary, given a slide from a PowerPoint and a list of names to identify specific images to save
    imageData = {}

    for shape in slide.shapes:
        if shape.name in imageNames:
            imageData[shape.name] = {
                'left': shape.left,
                'top': shape.top,
                'width': shape.width,
                'height': shape.height
            }

    return imageData

def generate_tech_profile():

    # Load template presentation
    tp = Presentation("template_v2.pptx")

    numSites = 1
    singleImageNames = ["Switch1", "Switch2", "ManagementSwitch", "Server1", "Server2", "Server3", "Server4", "Server5", "Server6", "Server7", "Server8", "Server9", "Server10", "Server11", "Server12", "Storage", "Backup", "Virtualization"]
    doubleImageNames = ["Switch1_1", "Switch2_1", "Switch1_2", "Switch2_2", "ManagementSwitch1", "ManagementSwitch2", "Server1_1", "Server2_1", "Server3_1", "Server4_1", "Server5_1", "Server6_1", "Server1_2", "Server2_2", "Server3_2", "Server4_2", "Server5_2", "Server6_2", "Storage1", "Storage2", "Backup1", "Backup2", "Virtualization"]

    # Load images' left, top, width, and height data from the slide based on if this is a single or double site setup
    slide = tp.slides[numSites - 1]

    if numSites == 1:
        imageData = load_image_data(slide, singleImageNames)

        # Print the results
        for name, dims in imageData.items():
            print(f"Shape: {name}")
            print(f"  Left: {dims['left']}")
            print(f"  Top: {dims['top']}")
            print(f"  Width: {dims['width']}")
            print(f"  Height: {dims['height']}")

    else:
        imageData = load_image_data(slide, doubleImageNames)

        # Print the results
        for name, dims in imageData.items():
            print(f"Shape: {name}")
            print(f"  Left: {dims['left']}")
            print(f"  Top: {dims['top']}")
            print(f"  Width: {dims['width']}")
            print(f"  Height: {dims['height']}")



    # Edit text fields

    # for 

    # tp.save("test.pptx")

    # return

if __name__ == "__main__":
    generate_tech_profile()
    print("Done!")