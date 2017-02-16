import sys
import os
import re
import win32com.client
import subprocess
import time
import datetime

######## Events and stuff for CANoe ########

class measurements:
	def __init__(self):
		print "measurement events up!"
	def OnInit(self):
		print "measurement started"

def load():
	global mCANoeApp
	mCANoeApp = win32com.client.Dispatch('CANoe.Application')
	cfg = os.path.abspath(r'..\..\Others\MACANPA_Tools\CANoe\SimulationMMCanoe5\kombi_mm_macanpa_for_CANoe8_1.cfg')
	mCANoeApp.Open(cfg)

######## functions to handle commands from NS ########

def counter(EnvVar, val):
	mVar = win32com.client.Dispatch(mCANoeApp.Environment.GetVariable(EnvVar))
	while(val):
		mVar.Value = 0 if mVar.Value == 1 else 1
		time.sleep(0.8)
		val -= 1
	mVar.Value = 0

def iterate(EnvVar, val):
	k = 0
	mVar = win32com.client.Dispatch(mCANoeApp.Environment.GetVariable(EnvVar))
	while(k != val):
		mVar.Value = k
		time.sleep(1)
		k += 1
		
def setter(EnvVar, val):
	mVar = win32com.client.Dispatch(mCANoeApp.Environment.GetVariable(EnvVar))
	mVar.Value = val

######## Start meaesurements and send the first KL15 ########
	
def start():
	mKl15 = win32com.client.Dispatch(mCANoeApp.Environment.GetVariable("EnvZAS_Kl_15_"))
	mMFL = win32com.client.Dispatch(mCANoeApp.Environment.GetVariable("Env_MFL_Gen2"))
	mMeasurement = win32com.client.DispatchWithEvents(mCANoeApp.Measurement, measurements) # \m/!
	mMeasurement.Start()
	mKl15.Value = 1
	mMFL.Value = 1
	
######## Read a NS script and parse commands ########

def readNS(file1):
	with open('Others\NinjaScript\\' + file1) as f:
		for line in f:
			command = line.split('\t')
			if command[0] == 'DUMP':
				try:
					address = command[1]
					size = command[2]
					dumpFile = dump(address, size.rstrip())
				except:
					try:
						section = command[1]
						dumpFile = dump(section.rstrip())
					except:
						print "Some problem in DUMP -"
			elif command[0] == 'DANDC':
				try:
					section = command[1]
					dumpFile = dump(section.rstrip())
					time.sleep(2) # to ensure that the dump is complete good and proper
					print "converting...\nfile: " + dumpFile + "\nlength: " + command[2] + "\nwidth: " + command[3] + "\nFormat: " + command[4]
					convert(str(dumpFile), command[2], command[3], command[4].rstrip())
				except:
					print "Some problem in DUMP AND CONVERT -"
			elif command[0] == 'COUNT':
				try:
					counter(command[1], int(command[2]))
				except:
					print "Some problem in COUNT -"
					break
			elif command[0] == 'LOOP':
				try:
					iterate(command[1], int(command[2]))
				except:
					print "Some problem in LOOP -"
			elif command[0] == 'SET':
				try:
					setter(command[1], int(command[2]))
				except:
					print "Some problem in SET -"
			elif command[0] == 'PAUSE':
				print "Paused"
				time.sleep(int(command[1]))
			else:
				print "Wrong Command!"

######## do stuff with captured framebuffers ########

def diff(pixelInfo1, pixelInfo2, delta, l = 0, w = 0, x1 = 0, y1 = 0, x2 = 0, y2 = 0): # generator function! \m/
	diff = {}
	k = 0
	if l != 0 and w != 0:
		for row in range(int(l)):
			for col in range(int(w)):
				channels1 = (pixelInfo1[row][col]).split(',')
				r1 = int(channels1[0])
				g1 = int(channels1[1])
				b1 = int(channels1[2])
				channels2 = (pixelInfo2[row + delta][col + delta]).split(',')
				r2 = int(channels2[0])
				g2 = int(channels2[1])
				b2 = int(channels2[2])
			
				if abs(r1 - r2) >= 3 or abs(g1 - g2) >= 3 or abs(b1 - b2) >= 3:
					key = str(row) + ', ' + str(col)
					diff[key] = str(abs(r1 - r2)) + ', ' + str(abs(g1 - g2)) + ' ' + str(abs(b1 - b2))
					k += 1
				else:
					k -= 1
	else: # cropping!
		for row in range(int(x1), int(x2)):
			for col in range(int(y1), int(y2)):
				channels1 = (pixelInfo1[row][col]).split(',')
				r1 = int(channels1[0])
				g1 = int(channels1[1])
				b1 = int(channels1[2])
				channels2 = (pixelInfo2[row + delta][col + delta]).split(',')
				r2 = int(channels2[0])
				g2 = int(channels2[1])
				b2 = int(channels2[2])

				if abs(r1 - r2) >= 3 or abs(g1 - g2) >= 3 or abs(b1 - b2) >= 3:
					key = str(row) + ', ' + str(col)
					diff[key] = str(abs(r1 - r2)) + ', ' + str(abs(g1 - g2)) + ' ' + str(abs(b1 - b2))
					k += 1
				else:
					k -= 1
	diff['k'] = k
	yield diff
				
def compare(file1, file2, x1 = 0, y1 = 0, x2 = 0, y2 = 0):
	
	# startTime = datetime.now()
	record = open('record.txt', 'w')

	buffer1 = open(file1, 'rb')
	buffer2 = open(file2, 'rb')

	header1 = buffer1.read(15) # read header
	header2 = buffer2.read(15) # TODO: Change this to 15 when data is captured from debugger. DONE
	
	result = re.search(r'^(\d+)\s(\d+)', header1, re.MULTILINE)
	l1 = int(result.group(1))
	w1 = int(result.group(2))
	result = re.search(r'^(\d+)\s(\d+)', header2, re.MULTILINE)
	l2 = int(result.group(1))
	w2 = int(result.group(2))

	data1 = buffer1.read()
	data2 = buffer2.read()

	data1 = [ord(c) for c in data1]
	data2 = [ord(c) for c in data2]
	
	########### first, let's make sure that buffer2 has data from the larger image ###########
	k = 0
	l = l2 if l2 > l1 else l1
	w = w2 if w2 > w1 else w1
	data = data2 if len(data2) > len(data1) else data1
	pixelInfo2 = [[0 for i in range(l)] for j in range(w)]
	for row in range(l):
		for col in range(w):
			pixelInfo2[row][col] = (str(data[k]) + ", " + str(data[k + 1]) + ", " + str(data[k + 2])) # got the pixel data in a 2D array
			k += 3

	k = 0
	l = l1 if l1 < l2 else l2
	w = w1 if w1 < w2 else w2
	data = data2 if len(data2) < len(data1) else data1
	pixelInfo1 = [[0 for i in range(l)] for j in range(w)]
	for row in range(l):
		for col in range(w):
			pixelInfo1[row][col] = (str(data[k]) + ", " + str(data[k + 1]) + ", " + str(data[k + 2])) # got the pixel data in a 2D array
			k += 3

	k = 0
	delta = abs(l2 - l1) / 2
		
	if x1 == 0 and y1 == 0 and x2 == 0 and y2 == 0: # pixel-to-pixel match both images as a whole
		
		l = l1 if l1 < l2 else l2
		w = w1 if w1 < w2 else w2
					
		for d in diff(pixelInfo1, pixelInfo2, delta, l, w):
			k = d['k']
			for key in d:
				record.write('(' + str(key) + ') ' + str(d[key]) + '\n') 
	else: # crop a certain area from (x1, y1) to (x2, y2). The logic remains the same but the area to be matched is smaller
		for d in diff(pixelInfo1, pixelInfo2, delta, 0, 0, x1, y1, x2, y2):
			k = d['k']
			for key in d:
				record.write('(' + str(key) + ') ' + str(d[key]) + '\n') 

	print "k = " + str(k)
	record.close()
	# print datetime.now() - startTime

def convert(aFile, length, width, format):
	iF = open(aFile, 'rb')
	bin = iF.read()
	filename = aFile.split('.')[0]
	oF = open(filename + '.ppm', 'wb')
	oF.write('P6\n')
	oF.write(length + ' ' + width + '\n')
	
	bin = [ord(c) for c in bin] # change each character read to its corresponding value in UNICODE
	i = 0
	
	if format == 'RGB565':
		oF.write('255\n')
		while i < len(bin): # do some bit shift magic.
			n = bin[i] * 0x1 + bin[i + 1] * 0x100
			b = (n >> 0) & 0x1f
			g = (n >> 5) & 0x3f
			r = (n >> 11) & 0x1f
			# 3, 2, 3 because 5, 6, 5 ;)
			r = r << 3 
			g = g << 2
			b = b << 3
			#print r, g, b
			oF.write(chr(r)+chr(g)+chr(b)) # then change each value back to a single character representing RGB value.
			i += 2
	elif format == 'RGB6666':
		oF.write('63\n')
		while i < len(bin):
			n = bin[i] * 0x1 + bin[i + 1] * 0x100 + bin[i + 2] * 0x10000
			a = (n >> 0)  & 0x3f
			b = (n >> 6)  & 0x3f
			g = (n >> 12) & 0x3f
			r = (n >> 18) & 0x3f
			#print r ,g ,b
			oF.write(chr(r)+chr(g)+chr(b)) # then change each value back to a single character representing RGB value.
			i += 3
	elif format == 'RGB8888':
		oF.write('255\n')
		while i < len(bin):
			n = bin[i] * 0x1 + bin[i + 1] * 0x100 + bin[i + 2] * 0x10000 + bin[i + 3] * 0x1000000
			a = (n >> 0)  & 0xff
			b = (n >> 8)  & 0xff
			g = (n >> 16) & 0xff
			r = (n >> 24) & 0xff
			#print r ,g ,b
			oF.write(chr(r)+chr(g)+chr(b)) # then change each value back to a single character representing RGB value.
			i += 4
	else:
		print "Wrong format!"

	oF.close() 
	iF.close()
	
######## Dumping from memory and stuff ########

def returnSizeAndAddress(fromFile1, section): # another generator function!
	sizeAndAddress = []
	flag = 0
	for someLine in fromFile1:
		result = re.search(r'(?<=([0-9A-Fa-f]{8})\+([0-9A-Fa-f]{6}))\s' + section, someLine) # look-behind assertion
		if result != None:
			sizeAndAddress = [result.group(1)]
			sizeAndAddress.append(result.group(2))
			break
	yield sizeAndAddress

def dump(address, size=None):
	debugFile = open("IDE\GHS\debugger.txt", 'w')
	path = 'IDE\GHS\DebuggerData\\'
	debugFile.write('halt\n')
	if size != None: # address and the size was passed
		ts = str(time.time()).split('.')[0]
		filename = ts + '.bin'
		print "dumping " + ts + ".bin"
		debugFile.write('memdump raw ' + filename + ' ' + address + ' ' + size + '\n')
	else: # address is the section name
		print "dumping " + address + ".bin"
		ts = str(time.time()).split('.')[0]
		filename = path + address + '_' + ts + '.bin'
		with open('IDE\GHS\hello.map') as f:
			for sizeAndAddress in returnSizeAndAddress(f, address):
				debugFile.write('memdump raw ' + filename + ' ' + '0x' + sizeAndAddress[0] + ' ' + '0x' + sizeAndAddress[1] + '\n')

	debugFile.write('C\n')
	debugFile.write('wait -time 2000\n')
	debugFile.write('quit')
	debugFile.close()

	p = subprocess.Popen('IDE\\GHS\\tool.cmd', creationflags=subprocess.CREATE_NEW_CONSOLE).wait() # important because otherwise, it starts a subprocess inside the same console and severes the COM
	print "done"
	return filename

def helpText():
	print "Strobe implements the following commands. You're not getting any more.\n"
	print "load: Let the games begin!\n"
	print "start: Start measurements and send the first KL15.\n"
	print "run <file.ns>: read a NinjaScript file and run the action sequence. You must have started the measurements beforehand. Like obviously.\n"
	print "dump 0x<address> 0x<size> | <section name>: dump <size> memory from a given address or just pass the section name, I'll calculate the size myself.\nMust have a hardware debugger connected for this to work. I like GHS Debugger very much.\n"
	print "convert <buffer.bin> <length> <width> <format>: convert a framebuffer to an image of given length and width.\nI do not have intelligence, you must tell me the length and width. Format can be one of RGB565, RGB6666 or RGB8888. I can't do more than that for you.\n"
	print "compare <image1.ppm> <image2.ppm>: compare two raw PPM images. PPM only. Because it is easier to read and I'm lazy.\nComparison prints a negative value (k). A more negative k signifies more similarity in images\n"
	print "compare <image1.ppm> <image2.ppm> x1 y1 x2 y2: crop area from (x1, y1) to (x2, y2) from the smaller (dimensionally) image\nand compare with the same area in the larger (dimensionally) image. Comparison prints a negative value (k). A more negative k signifies more similarity in images\n"
	print "help: print this crap."
	print "exit: well, you know what this means obviously.\n"

global loadFlag
global startFlag
global mapFile
loadFlag = 0
startFlag = 0

# startup hook: make the directory where Strobe dumps data from the debugger and read the test scripts.
if not os.path.exists('IDE\GHS\DebuggerData'):
    os.makedirs('IDE\GHS\DebuggerData')
if not os.path.exists('Others\NinjaScript'):
    os.makedirs('Others\NinjaScript')

while(1): # one infinite loop
	command = raw_input('strobe> ')
	subs = command.split(' ') # subs[0] == the command. subs[1..n] == options
	if subs[0] == 'exit':
		break # exit strobe
	elif subs[0] == 'help':
		helpText()
	elif subs[0] == 'run':
		try:
			if subs[1].split('.')[1] != 'ns':
				print "invalid NinjaScript. You can't fool me, you know."
				continue
		except:
			print "Run what?!"
			continue
		readNS(subs[1])
	elif subs[0] == 'compare':
		try:
			if subs[1].split('.')[1] == 'ppm' and subs[2].split('.')[1] == 'ppm':
				try:
					compare(subs[1], subs[2], subs[3], subs[4], subs[5], subs[6])
				except:
					compare(subs[1], subs[2])
			else:
				print "not a PPM file. I told you, you can't fool me."
				continue
		except:
			print "Comparison is between two things. I hope you know that."
			continue
	elif subs[0] == 'dump':
		try:
			address = subs[1]
			size = subs[2]
			dumpFile = dump(address, size)
			print "file: " + dumpFile + "\n"
		except:
			try:
				section = subs[1]
				dumpFile = dump(section)
				print "file: " + dumpFile + "\n"
			except:
				print "I'd rather dump you."
				continue
	elif subs[0] == 'convert':
		try:
			print "file: " + subs[1]
			print "length: " + subs[2]
			print "width: " + subs[3]
			print "format: " + subs[4]
			convert(subs[1], subs[2], subs[3], subs[4])
		except:
			print "I need more information than that to convert your binary crap to a beautiful picture."
			continue
	elif subs[0] == 'load':
		try:
			print "opening config..."
			load()
			loadFlag = 1
		except:
			print "I don't load crap."
	elif subs[0] == 'start':
		if loadFlag == 1:
			print "Starting measurement."
			start()
			startFlag = 1
		else:
			print "what do you want to start? An affair? I'm up!"
			continue
	else:
		print subs[0] + " not included."
		helpText()
print "Ciao!\n... Ninja."
time.sleep(3)

# NINJA
