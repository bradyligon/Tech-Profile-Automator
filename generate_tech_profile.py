import comtypes.client
import os
import io
from pptx import Presentation
import pptx.shapes
from pptx.util import Inches

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
    singleImageNames = ["Switch1", "Switch2", "ManagementSwitch"]
    doubleImageNames = []

    # Load images' left, top, width, and height data from the slide based on if this is a single or double site setup
    singleImageData = {}
    doubleImageData = {}

    if numSites == 1:
        singleImageData = load_image_data(tp.slides[numSites - 1], singleImageNames)

        # Print the results
        for name, dims in singleImageData.items():
            print(f"Shape: {name}")
            print(f"  Left: {dims['left']}")
            print(f"  Top: {dims['top']}")
            print(f"  Width: {dims['width']}")
            print(f"  Height: {dims['height']}")
    else:
        doubleImageData = load_image_data(tp.slides[numSites - 1], doubleImageNames)

        # Print the results
        for name, dims in doubleImageData.items():
            print(f"Shape: {name}")
            print(f"  Left: {dims['left']}")
            print(f"  Top: {dims['top']}")
            print(f"  Width: {dims['width']}")
            print(f"  Height: {dims['height']}")

    return

if __name__ == "__main__":
    generate_tech_profile()
    print("Done!")