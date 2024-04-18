# This app is designed to aid the user in distinguishing a single PillSense Capsule from a group # 

import tkinter as tk
import serial
import time
import tkinter as tk
from tkinter import*
from tkinter import messagebox
from os.path import join
import serial.tools.list_ports
# from tkinter import filedialog
# from reportlab.pdfgen import canvas
# from reportlab.lib import colors
from datetime import datetime
# from fpdf import FPDF 
import csv
import os.path
# from barcode import Code39
from barcode import Code128 # Sepha can only accept Code 128
from barcode.writer import ImageWriter
import threading
import win32api
import win32print
# import pathlib
from PIL import Image, ImageFont, ImageDraw
from tkinter import filedialog
import re
from tkinter import ttk
import tkinter.messagebox


# Create the main window
window = tk.Tk()

# Create the toplevel window
top = Toplevel(window)
top.title("Instructions")

w = 660 # width for the Tk root
h = 535 # height for the Tk root

width= window.winfo_screenwidth()
height= window.winfo_screenheight()

# calculate x and y coordinates for the Tk root window
x = (width/1.8) - (w/1.8)
y = (height/8) - (h/8)

# Set window title (optional)
window.title('Select Capsule')

# Set window size (optional)
window.geometry('590x220+%d+%d' % (x, y))


# global definitions #

Cap1 = Radiobutton()
Cap2 = Radiobutton()
Cap3 = Radiobutton()
Cap4 = Radiobutton()
Cap5 = Radiobutton()

Date_Time = datetime.now().strftime("%Y %m %d,  %H %M")
colours = ["red","green","blue","black"]	
# alreadyPrinted = []
# password = '2476'
# pwAttempt = StringVar()
CapsuleSelected = IntVar()
# selected_location = ''
# save_location = ''
Port_Found = False
i=0
com = ""



# Functions #

def CapOneSelected():
	global Cap2
	global Cap3
	global Cap4
	global Cap5

	Cap2.deselect()
	Cap3.deselect()
	Cap4.deselect()
	Cap5.deselect()

	# SubmitButton.configure(state='normal')


def CapTwoSelected():
	global Cap1
	global Cap3
	global Cap4
	global Cap5

	Cap1.deselect()
	Cap3.deselect()
	Cap4.deselect()
	Cap5.deselect()

	# SubmitButton.configure(state='normal')


def CapThreeSelected():
	global Cap1
	global Cap2
	global Cap4
	global Cap5

	Cap1.deselect()
	Cap2.deselect()
	Cap4.deselect()
	Cap5.deselect()

	# SubmitButton.configure(state='normal')


def CapFourSelected():
	global Cap1
	global Cap2
	global Cap3
	global Cap5

	Cap1.deselect()
	Cap2.deselect()
	Cap3.deselect()
	Cap5.deselect()

	# SubmitButton.configure(state='normal')


def CapFiveSelected():
	global Cap1
	global Cap2
	global Cap3
	global Cap4
	
	Cap1.deselect()
	Cap2.deselect()
	Cap3.deselect()
	Cap4.deselect()

	# SubmitButton.configure(state='normal')

def Set_COMport():

	t = list(serial.tools.list_ports.comports())
	global Port_Found
	global ser
	Port_Found = False

	for p in t:
		if p.device == com:
			Port_Found = True
			# break
	
	if Port_Found == True:

		# COM_Port_frame.pack_forget()

		ser = serial.Serial(com, baudrate=9600, timeout=3)
		ser.reset_input_buffer()
		ser.reset_output_buffer()
		ser.write(b'\xAB\xCD\xEF\xFF\xFF\xFF')
		ser.flush()
		time.sleep(0.100)

		# data = ser.readline(8).decode().strip()
		# print(data)
		
		# Label2 = Label(COM_Port_frame2, text= 'Successfully connected to ' + COMport.get(), fg='blue')
		# Label2.grid(row=1, column=1)
		
		# Ava_COM_Ports_t.config(state='disabled')
		# Ava_COM_Ports.config(state='disabled')

		# for child in COM_Port_frame2.winfo_children():
		# 	child.configure(state='disabled')

		# for child in Search_frame.winfo_children():
		# 	child.configure(state='normal')
		
	# else :
		# Label2 = Label(COM_Port_frame2, text= '               Port does not exist!           ', fg='red')
		# Label2.grid(row=1, column=1)
		# for child in Search_frame.winfo_children():
		# 	child.configure(state='disabled')
			
def Refresh():
	"""
	Disables refresh button, starts progress bar, calls CheckForCapsules().

	Args:
		None.

	 Returns:
		None.
	"""
	# window.withdraw()
	RefreshButton.config(state=DISABLED)
	progress_bar.start(10)
	progress_bar.grid()
	
	Cap1LEDvalue.config(fg = 'grey')
	Cap2LEDvalue.config(fg = 'grey')
	Cap3LEDvalue.config(fg = 'grey')
	Cap4LEDvalue.config(fg = 'grey')
	Cap5LEDvalue.config(fg = 'grey')

	t = threading.Thread(target=CheckForCapsules)
	t.start()	
	# CheckForCapsules()


def CheckForCapsules():
	"""
	reads address' of capsules in range, reads values for each capsule in range
	
	Args:
		None.

	 Returns:
		None.
	"""

	global Capsules
	global ser


	# SearchButton.config(state=DISABLED)
	# SearchButton.update()

	# Unique_ID_t.config(state=NORMAL)
	# Unique_ID_t.config(foreground="grey")
	# Unique_ID_t.delete(1.0,END)
	# Unique_ID_t.insert(INSERT, 'Checking for Capsules... ')
	# Unique_ID_t.update()
	# Unique_ID_t.config(state=DISABLED)

	try:
		# unpair / flush memory
		ser.reset_input_buffer()
		ser.reset_output_buffer()
		ser.write(b'\xAB\xCD\xEF\xFF\xFF\xFF')
		ser.flush()

		# Read address'
		time.sleep(2)
		ser.reset_input_buffer()
		ser.reset_output_buffer()
		ser.write(b'\xAA\xBB\xCC\xFF\xFF\xFF')
		ser.flush()
		time.sleep(0.100)

		Capsules = ser.readline(44).decode().strip()# read Capsule Addresses
		Capsules = Capsules.split(',') # Separate by ','
		print(Capsules)

		# If launhpad adds 'F' to the begining of any Capsule ID, replace with '0'
		temp=[]
		for i in Capsules:
			i = list(i)
			if i[0] == 'F': #swap
				# print(i[0])
				i[0] = '0' #swap
				# print(i[0])
				temp.append(''.join(map(str, i)))
			else:
				temp.append(''.join(map(str, i)))

		Capsules = temp # Update the list of Capsule IDs
		print("Capsules:", Capsules)


		# for child in window.winfo_children(): # enable all buttons (needed for multiple Tests in a row)
		# 	child.configure(state='normal')


		# SubmitButton.configure(state='disabled')
		Cap1.config(text=Capsules[0])
		Cap2.config(text=Capsules[1])
		Cap3.config(text=Capsules[2])
		Cap4.config(text=Capsules[3])
		Cap5.config(text=Capsules[4])
		

		# If no Capsule is present (I.e, '00000000'), disable that RadioButton
		if Capsules[0] == '00000000':
			# for child in window.winfo_children():
			Cap1LEDvalue.config(text = "N/A", fg = 'grey')
			Cap1.configure(state='disabled')
		else:
			# pair address A
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xAA\xAA\xAA\xFF\xFF\xFF')
			ser.flush()
			time.sleep(2)

			# Sensor value
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xDD\xEE\xBB\xFF\xFF\xFF')
			ser.flush()
			time.sleep(0.100)

			# read values
			Cap1value = ser.readline(24).decode().strip()
			# Cap1value = Cap1value.split(",")
			print(Cap1value)
			
			# function to check if received data is in the correct format
			if is_valid_format(Cap1value):
				Cap1LEDvalue.config(text = Cap1value, fg = 'black')
				Cap1.configure(state='normal')

				if is_target_capsule(Cap1value):
					Cap1LEDvalue.config(text = Cap1value, fg = 'Green')

				if not is_LedValue_good(Cap1value):
					Cap1LEDvalue.config(text = Cap1value, fg = 'Red')
			else:
				Cap1LEDvalue.config(text = "N/A", fg = 'grey')


			# unpair
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xAB\xCD\xEF\xFF\xFF\xFF')
			ser.flush()
			time.sleep(0.100)
			

		if Capsules[1] == '00000000':
			# for child in window.winfo_children():
			Cap2LEDvalue.config(text = "N/A", fg = 'grey')
			Cap2.configure(state='disabled')
		else:

			# pair address B
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xBB\xBB\xBB\xFF\xFF\xFF')
			ser.flush()
			time.sleep(2)

			# Sensor value
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xDD\xEE\xBB\xFF\xFF\xFF')
			ser.flush()
			time.sleep(0.100)

			# read values
			Cap2value = ser.readline(24).decode().strip()
			print(Cap2value)
			
			# function to check if received data is in the correct format
			if is_valid_format(Cap2value):
				Cap2LEDvalue.config(text = Cap2value, fg = 'black')
				Cap2.configure(state='normal')

				if is_target_capsule(Cap2value):
					Cap2LEDvalue.config(text = Cap2value, fg = 'Green')

				if not is_LedValue_good(Cap2value):
					Cap2LEDvalue.config(text = Cap2value, fg = 'Red')
			else:
				Cap2LEDvalue.config(text = "N/A", fg = 'grey')


			# unpair
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xAB\xCD\xEF\xFF\xFF\xFF')
			ser.flush()
			time.sleep(0.100)


		if Capsules[2] == '00000000':
			# for child in window.winfo_children():
			Cap3LEDvalue.config(text = "N/A", fg = 'grey')
			Cap3.configure(state='disabled')
		else:
			# pair address C
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xCC\xCC\xCC\xFF\xFF\xFF')
			ser.flush()
			time.sleep(2)

			# Sensor value
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xDD\xEE\xBB\xFF\xFF\xFF')
			ser.flush()
			time.sleep(0.100)

			# read values
			Cap3value = ser.readline(24).decode().strip()
			print(Cap3value)
			
			# function to check if received data is in the correct format
			if is_valid_format(Cap3value):
				Cap3LEDvalue.config(text = Cap3value, fg = 'black')
				Cap3.configure(state='normal')

				if is_target_capsule(Cap3value):
					Cap3LEDvalue.config(text = Cap3value, fg = 'Green')

				if not is_LedValue_good(Cap3value):
					Cap3LEDvalue.config(text = Cap3value, fg = 'Red')
			else:
				Cap3LEDvalue.config(text = "N/A", fg = 'grey')

			

			# unpair
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xAB\xCD\xEF\xFF\xFF\xFF')
			ser.flush()
			time.sleep(0.100)


		if Capsules[3] == '00000000':
			# for child in window.winfo_children():
			Cap4LEDvalue.config(text = "N/A", fg = 'grey')
			Cap4.configure(state='disabled')	
		else:
			# pair address D
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xDD\xDD\xDD\xFF\xFF\xFF')
			ser.flush()
			time.sleep(2)

			# Sensor value
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xDD\xEE\xBB\xFF\xFF\xFF')
			ser.flush()
			time.sleep(0.100)

			# read values
			Cap4value = ser.readline(24).decode().strip()
			print(Cap4value)
			
			# function to check if received data is in the correct format
			if is_valid_format(Cap4value):
				Cap4LEDvalue.config(text = Cap4value, fg = 'black')
				Cap4.configure(state='normal')

				if is_target_capsule(Cap4value):
					Cap4LEDvalue.config(text = Cap4value, fg = 'Green')

				if not is_LedValue_good(Cap4value):
					Cap4LEDvalue.config(text = Cap4value, fg = 'Red')
			else:
				Cap4LEDvalue.config(text = "N/A", fg = 'grey')
			
			
			# unpair
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xAB\xCD\xEF\xFF\xFF\xFF')
			ser.flush()
			time.sleep(0.100)


		if Capsules[4] == '00000000':
			# for child in window.winfo_children():
			Cap5LEDvalue.config(text = "N/A", fg = 'grey')
			Cap5.configure(state='disabled')
		else:
			# pair address E
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xEE\xEE\xEE\xFF\xFF\xFF')
			ser.flush()
			time.sleep(2)

			# Sensor value
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xDD\xEE\xBB\xFF\xFF\xFF')
			ser.flush()
			time.sleep(0.100)

			# read values
			Cap5value = ser.readline(24).decode().strip()
			print(Cap5value)
			
			# function to check if received data is in the correct format
			if is_valid_format(Cap5value):
				Cap5LEDvalue.config(text = Cap5value, fg = 'black')
				Cap5.configure(state='normal')
				
				if is_target_capsule(Cap5value):
					Cap5LEDvalue.config(text = Cap5value, fg = 'Green')

				if not is_LedValue_good(Cap5value):
					Cap5LEDvalue.config(text = Cap5value, fg = 'Red')
			else:
				Cap5LEDvalue.config(text = "N/A", fg = 'grey')

			
			# unpair
			ser.reset_input_buffer()
			ser.reset_output_buffer()
			ser.write(b'\xAB\xCD\xEF\xFF\xFF\xFF')
			ser.flush()
			time.sleep(0.100)

		# display the window
		# window.deiconify()
		progress_bar.grid_remove()
		RefreshButton.config(state=NORMAL)

		# Unique_ID_t.config(state=NORMAL)
		# Unique_ID_t.delete(1.0,END)
		# Unique_ID_t.update()
		# Unique_ID_t.config(state=DISABLED)
		
		ser.reset_input_buffer()
		ser.reset_output_buffer()
		ser.write(b'\xCC\xEE\xBB\xFF\xFF\xFF')
		ser.flush()

	except IOError: # Connection to COM port lost
		ser = serial.Serial(com, baudrate=9600, timeout=3)
		# SearchButton.config(state=NORMAL)
		# Unique_ID_t.config(state=NORMAL)
		# Unique_ID_t.delete(1.0,END)
		# Unique_ID_t.config(state=DISABLED)
		
		messagebox.showerror("Warning!", "Error has occured connecting to COM port!" + 
									"\n" +  "\n" + "                  Please try again")


def set_COM():

	global com
	
	myports=([comport.description for comport in serial.tools.list_ports.comports()]) # Get list of COM ports
	COMstr= ''.join(myports)	#convert list to string

	findUART= COMstr.find('UART')

	if (findUART!=-1):
		#UART (COMXX) --> from U to C is 6 in addition to OMXX
		setUART= (COMstr[findUART+6:findUART+11])
		setUART=setUART.strip()		# remove the space if any
		setUART=setUART.strip(')')	# remove the ')' if any

		com = setUART

	print("Setting COM port")
	t = list(serial.tools.list_ports.comports())
	global Port_Found
	global ser
	Port_Found = False
	for p in t:
		if p.device == com:
			Port_Found = True
			# break
	
	if Port_Found == True:

		# COM_Port_frame.pack_forget()

		ser = serial.Serial(com, baudrate=9600, timeout=3)
		ser.reset_input_buffer()
		ser.reset_output_buffer()
		ser.write(b'\xAB\xCD\xEF\xFF\xFF\xFF')
		ser.flush()
		time.sleep(0.100)

		# for child in Search_frame.winfo_children():
		# 	child.configure(state='normal')

		CheckForCapsules()

def No_Communication():
	messagebox.showerror('UART Timeout', 'No data received from Launchpad!' 
							+ '\n' + 'Check the UART Port and make sure Radio Communication  is stable')


def is_valid_format(text):
	"""
	Checks if a string is in the format "xxxx,xxxx,xxxx,xxxx,xxxx" where "x" is any digit 0-9.

	Args:
		text: The string to be checked.

	 Returns:
		True if the string is in the valid format, False otherwise.
	"""

	regex = r"^\s*(\d+\s*)*,\s*[0-9\s]{4}\s*,\s*[0-9\s]{4}\s*,\s*[0-9\s]{4}\s*,\s*[0-9\s]{4}\s*$"
	return bool(re.match(regex, text))


def is_LedValue_good(RGBdata):
	"""
	Checks that each value is greater than 1000 (values below this threshold indicate a faulty LED / PD).
	
	Args:
		RGBdata. the string to be checked.

	 Returns:
		True if the string is in the valid format, False otherwise.
	"""

	try:
		RGBdata = RGBdata.split(',')
		RGBdata = [ int(x) for x in RGBdata ]
		
		return all(value >= 1700 for value in RGBdata[:4])

	except:
		print("An exception occurred in 'def is_LedValue_good'")
		return False


def is_target_capsule(RGBdata):
	"""
	Checks if a capsule is the target I.e., it's being covered (OFF value is low).
	
	Args:
		RGBdata. the string to be checked.

	 Returns:
		True if the OFF value is in the less than 100, False otherwise.
	"""
	
	try:
		RGBdata = RGBdata.split(',')
		OFFvalue = int(RGBdata[4])

		if OFFvalue < 100:
			return True
		else:
			return False
		
	except:
		print("An exception occurred in 'def is_target_capsule'")
		return False


def askQuit(): # If the user attempts to exit the app
	"""
	askyesno messagebox: "Do you want to quit?".
	
	Args:
		None.

	 Returns:
		None.
	"""
	if messagebox.askyesno("Quit", "Do you want to quit?"):
		top.destroy()
		window.destroy()


def disable_Exit():
	"""
	This function does nothing. It simple invokes "pass"
	"""
	pass


# Function to open the toplevel window
def open_instructions():
  
	# Define the text instructions
	instructions = """
	1.	Place the Capsule you wish to pair in the fixture (P00378-01A).

	2.	Cover the capsule to sheild it from light (this lowers the 'OFF' value).
	
	3.	Capsules with an 'OFF' value less than 100 are highlighted in Green.
		
		Notes:
		 - If a value is N/A this means the LED data was received in the incorrect format. press refresh.
		 - The LED values are in the format R,G,B,FR,OFF.
	"""

	# Create a label to display the instructions
	label = Label(top, text=instructions, justify="left")
	label.pack(padx=10, pady=10)

	# Add a close button (optional)
	# close_button = Button(top, text="Close", command=top.destroy)
	# close_button.pack()

	# Focus the toplevel window
	#   top.grab_set()

# Create a button to open the instructions window (optional)
# button = Button(root, text="Open Instructions", command=open_instructions)
# button.pack()

# Open the instructions window directly (uncomment the following line 
# and remove the button creation code to open automatically on launch)



# Capsule Selection Window radio buttons and labels (OFF values for each capsule)

Cap1 = Label(window, width=15, font = 'arial 13')
Cap1.grid(row=1, column=1,padx=10, pady=0)
# Cap1.deselect()

Cap1LEDvalue = Label(window, text= "N/A", font = 'arial 10', fg="grey")
Cap1LEDvalue.grid(row=2, column=1, padx=20, pady=0)

Cap2 = Label(window, width=15, font = 'arial 13')
Cap2.grid(row=1, column=2,padx=10, pady=0)
# Cap2.deselect()

Cap2LEDvalue = Label(window, text= "N/A", font = 'arial 10', fg="grey")
Cap2LEDvalue.grid(row=2, column=2, padx=5, pady=0)

Cap3 = Label(window, width=15, font = 'arial 13')
Cap3.grid(row=1, column=3,padx=10, pady=0)
# Cap3.deselect()

Cap3LEDvalue = Label(window, text= "N/A", font = 'arial 10', fg="grey")
Cap3LEDvalue.grid(row=2, column=3, padx=20, pady=0)

Cap4 = Label(window, width=15, font = 'arial 13')
Cap4.grid(row=3, column=1,padx=10, pady=(20, 0), columnspan = 2)
# Cap4.deselect()

Cap4LEDvalue = Label(window, text= "N/A", font = 'arial 10', fg="grey")
Cap4LEDvalue.grid(row=4, column=1, padx=5, pady=0, columnspan = 2)

Cap5 = Label(window, width=15, font = 'arial 13')
Cap5.grid(row=3, column=2,padx=10, pady=(20, 0), columnspan = 2)
# Cap5.deselect()

Cap5LEDvalue = Label(window, text= "N/A", font = 'arial 10', fg="grey")
Cap5LEDvalue.grid(row=4, column=2, padx=5, pady=0, columnspan = 2)

# SubmitButton = Button(window, width = 15, text = 'Submit', command=SubmitCapsule)
# SubmitButton.grid(row=4, column=2,padx=5, pady=10, columnspan=2)

RefreshButton = Button(window, width = 15, text = 'Refresh', command=Refresh)
RefreshButton.grid(row=5, column=1,padx=5, pady=10, columnspan=4)

progress_bar = ttk.Progressbar(window, orient = HORIZONTAL, length = 200, mode = "indeterminate")
progress_bar.grid(row=6, column=1,padx=5, pady=10, columnspan=4)
progress_bar.grid_remove()


set_COM()
open_instructions()

# Start the main event loop
window.protocol("WM_DELETE_WINDOW", askQuit)
top.protocol("WM_DELETE_WINDOW", disable_Exit)
window.mainloop()
