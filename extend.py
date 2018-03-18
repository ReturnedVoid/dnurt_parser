import sys
import os
import re
import shutil

from colorama import init
from colorama import Fore, Back, Style

#init colorama
init()

#------------------------constants-------------------------
MODELS_FILES_PATH = 'upload/presses/3/monographs/10/submission/proof/Formulas/3d/' #path to models
IMAGES_FILES_PATH = 'upload/presses/3/monographs/10/submission/proof/Formulas/'    #path to images
MODELS_DIRECTORY_NAME = '3d/'

SOURCE_FILE_NAME = 'source.htm'
OUTPUT_FILE_NAME = 'index.html'

VIDEO_HOST = 'https://www.youtube.com/' #youtube by default
YOUTUBE_FRAME_WIDTH = 45 #width in percent
YOUTUBE_FRAME_HEIGHT = 45 #height in percent

MODEL_FRAME_WIDTH = 45 #width in percent
MODEL_FRAME_HEIGHT = 45 #height in percent
#----------------------------------------------------------

#for logging
youtube_links_count = 0 
models_count = 0
source_name = SOURCE_FILE_NAME.split('.')[0]
project_dir = os.getcwd().replace('\\', '/') #project root path

def printInfo():
	print ('All parsed links:\nIf some links are incorrect, contact script programmer\n')

def parseLine(string):
	rez = re.findall('>.*?<', string) #find words between >< tags
	if rez: #if we found some words
		temp  = '>'.join(rez) #create string, splitted by '>'

		words = [] #here will be all symbols between >< tags
		for i in temp:
			if i == '<' or i == '>': #we don`t need these symbols
				continue
			else:
				words.append(i) #append symbol to list

		return ''.join(words) #return words in string format
	return '' #return empty string if we didn`t find words

#Close given files
def closeFiles(*files):
	for file in files:
		file.close()

#returns frame by given Youtube link
def returnVideoFrame(srcLine):
	index = srcLine.find('v=') #video id starts after v=
	code = srcLine[index+2::] #skip v=
	src = '{0}embed/{1}'.format(VIDEO_HOST, code)
	print('youtube src = {}'.format(src))
	return ''' 
	<div align="center">
	<iframe align="middle" width="{0}%" height="{1}%"
	src="{2}" frameborder="0" allow="autoplay; encrypted-media" allowfullscreen></iframe>
	</div></br>
	'''.format(YOUTUBE_FRAME_WIDTH, YOUTUBE_FRAME_HEIGHT, src)

#returns frame by given 3d model`s link
def return3DFrame(filename):
	return ''' 
	<p><div align="center">
	<iframe align="middle" width="{0}%" height="{1}%"
	src="{2}" frameborder="0" allow="autoplay; encrypted-media" allowfullscreen></iframe>
	</div></p></br>
	'''.format(MODEL_FRAME_WIDTH,MODEL_FRAME_HEIGHT,filename)

def setImagesPath():
	file_r = open(OUTPUT_FILE_NAME, 'r') #open output file
	file_text = file_r.read()
	file_r.close()

	#replace generated path with new
	changed_images_path_text = re.sub('{}.files/'.format(source_name), IMAGES_FILES_PATH, file_text)
	changed_models_path_text = re.sub('3d/', MODELS_FILES_PATH, changed_images_path_text)

	file_w = open(OUTPUT_FILE_NAME, 'w')
	file_w.write( changed_models_path_text )
	closeFiles(file_w, file_r)

def removeDirTree(path):
	try:
		shutil.rmtree(path) #delete dir if already exists
	except:
		pass

def createDirTree(path):
	try:
		os.makedirs(path)
	except:
		pass

def copyFiles(source, out):
	try:
		os.chdir(source)
		for i in os.listdir():
			shutil.move(i, '{}/'.format(project_dir) + out)
		os.chdir('../')
	except FileNotFoundError:
		pass
	except shutil.Error as sh_e: #if error with copy or file`s already exists
		os.chdir('../')

def setDirs():
	#try to remove models dir tree if exists
	removeDirTree(MODELS_FILES_PATH.split('/')[0])

	#try to remove images dir tree if exists
	removeDirTree(IMAGES_FILES_PATH.split('/')[0])

	#create dir tree for models
	createDirTree(MODELS_FILES_PATH)

	#create dir tree for images
	createDirTree(IMAGES_FILES_PATH)

	#open 3d folder and copy all files to MODELS_FILES_PATH	
	copyFiles('3d', MODELS_FILES_PATH)

	#open .files folder and copy all files to OUTPUT_FILE_NAME
	copyFiles('{}.files'.format(source_name), IMAGES_FILES_PATH)

	#delete 3d folder
	removeDirTree('3d')

	#delete .files folder
	removeDirTree('{}.files'.format(source_name))

#deletes source file after parsing
def deleteSourceFile():
	try:
		os.remove('{}.htm'.format(source_name))
	except FileNotFoundError:
		print( '{}.htm not found'.format(source_name))		

def openFile(file_path, mode):
	try:
		return open(file_path, mode)
	except:
		sys.exit(Fore.RED + 'Critical error: file {} does not exist'.format(file_path))

def convertLinksToFrames():
	global youtube_links_count
	global models_count
	lines_buffer = []

	for line in file_r:
		if line.find('</p>') == -1: #if it`s not end of paragraph
			lines_buffer.append(line) #add this line to buffer
			continue
		if line.find('</p>') != -1: #if it`s end of paragraph
			lines_buffer.append(line)
			for string in lines_buffer:
				src = parseLine(string)

				#if it`s 3d model
				if src.startswith(MODELS_DIRECTORY_NAME):
					models_count += 1
					print('model path = {}'.format(src))
					file_w.write(return3DFrame(src))
					lines_buffer.clear()

				#if it`s youtube link
				elif src.startswith(VIDEO_HOST):
					youtube_links_count += 1
					file_w.write(returnVideoFrame(src)) #write frame to file
					lines_buffer.clear() #clear buffer

		file_w.write(''.join(lines_buffer)) #write from buffer to file
		lines_buffer.clear()

#create dirs for images and models
setDirs()

#header info
printInfo()

#open 2 files
file_w = openFile(OUTPUT_FILE_NAME, 'w') #create new file
file_r = openFile(SOURCE_FILE_NAME, 'r') #open generated by MS Word file

#converts links to frames
convertLinksToFrames()

#print links count
print ('\nmodels count = {}'.format(models_count))
print ('youtube links count = {}'.format(youtube_links_count))

#free resources
closeFiles(file_w, file_r)

deleteSourceFile()

#change old src
setImagesPath()