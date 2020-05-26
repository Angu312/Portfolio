# This script is meant to be run with the Trainable WEKA Segmentation tool within FIJI to apply a trained classifier model and apply them to a new set of images.

import os
from trainableSegmentation import WekaSegmentation, Weka_Segmentation
from ij import IJ, ImagePlus
from ij.plugin import LutLoader
 
def run():
	srcDir = r'C:\Users\Angu312\Pictures\Test\1. Pre Processed'
	dstDir = r'C:\Users\Angu312\Pictures\Test\2. Post Processed'
	for root, directories, filenames in os.walk(srcDir):
		filenames.sort();
		for filename in filenames:
      		# Check for file extension
			if not filename.endswith('.jpg'):
        			continue
			process(srcDir, dstDir, root, filename)
 
def process(srcDir, dstDir, currentDir, fileName):
	print "Processing:"

	# Opening the image
	print "Open image file", fileName
	image = IJ.openImage(os.path.join(srcDir, fileName))

	weka = WekaSegmentation()
	weka.setTrainingImage(image)

	# Manually loads trained classifier
	weka.loadClassifier(r'C:\Users\Angu312\Scripts\VectorClassifier.model')
	# Apply classifier and get results
	segmented_image = weka.applyClassifier(image, 0, False)
	# assign same LUT as in GUI. Within WEKA GUI, right-click on classified image and use Command Finder to save the "LUT" within Fiji.app\luts
	lut = LutLoader.openLut(r'C:\Users\Angu312\Documents\Fiji.app\luts\Golden ARC Lut.lut')
	segmented_image.getProcessor().setLut(lut)

	# Saving the image as a .tif file
	saveDir = dstDir
	if not os.path.exists(saveDir):
		os.makedirs(saveDir)
	print "Saving to", saveDir
	IJ.saveAs(segmented_image, "Tif", os.path.join(saveDir, fileName))
	image.close()

run()