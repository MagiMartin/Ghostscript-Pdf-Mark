import fitz  # PyMuPDF
import re
import ghostscript
import io
import subprocess
from datetime import date
import os
import glob
import time
import shutil
import math
import win32print
import win32api


def main():
	while True:
		print("Running")
		# Usage
		pdf_dir = os.getcwd() +  "\\input"
		mark_navn = []
		mark_koordinater_x = []
		mark_koordinater_y = []
		direction = []
		output_pdf = os.getcwd() + "\\output"
		
		gs_code = """
				/mark {% x y string direct
				/direct exch def
				/string exch def
				/y exch def
				/x exch def
				
				gsave
				direct (up) eq {
				%x y string stringwidth pop 2 div 25 add sub translate
				x 15 add y translate
				/showpage {{}} bind def
				90 rotate
				(mark.eps) run
				} if
				
				direct (right) eq {
				%x string stringwidth pop 2 div 25 add sub y translate
				x y 15 sub translate
				/showpage {{}} bind def
				(mark.eps) run
				} if
				
			    direct (down) eq {
				%x y string stringwidth pop 2 div 25 add add translate
				x 15 sub y translate
				/showpage {{}} bind def
				-90 rotate
				(mark.eps) run
				} if
				
				direct (left) eq {
				%x string stringwidth pop 2 div 25 add add y translate
				x y 15 add translate
				/showpage {{}} bind def
				-180 rotate
				(mark.eps) run
				} if
				
				grestore
				} def
				
				/center { % x y string direct
				/direct exch def
				/string exch def
				/y exch def
				/x exch def
				
				gsave
				direct (up) eq {
				%x y string stringwidth pop 2 div sub
				x 15 add y string stringwidth pop pop 25 add moveto
				%moveto
				90 rotate
				string
				show } if
				
				direct (right) eq {
				%x string stringwidth pop 2 div sub
				%y moveto
				x 25 add y 15 sub moveto
				string
				show } if
				
				direct (down) eq {
				%x y string stringwidth pop 2 div add
				%moveto
				x 15 sub y string stringwidth pop pop 25 sub
				moveto
				-90 rotate
				string show
				} if
				
				direct (left) eq {
				%x string stringwidth pop 2 div add
				%y moveto
				x string stringwidth pop pop 25 sub y 15 add moveto
				-180 rotate
				string
				show } if
				
				grestore
				} def
		
			<<
				/EndPage
				{
				2 eq { pop false }
				{
				gsave
    			0 0 0 1 (TRAFFIC BLACK)findcmykcustomcolor 1 setcustomcolor
				/Helvetica-Narrow        
				20 selectfont
				"""
				
		gs_command = [
				"gs",
				"-dBATCH",
				"-dNOPAUSE",
				"-dNOSAFER",
				"-dQUIET",
				"-sDEVICE=pdfwrite", 
				"-dDEVICEWIDTHPOINTS=8505",
				"-dDEVICEHEIGHTPOINTS=3500",
				"-sFONTPATH#C:\PROGRA~1\gs\gs10.02.1\Resource\Font",
				]

		pdf_path, pdf_name = find_files(pdf_dir)
		
		
		if len(pdf_path) > 0:
			try:
				vector_data = extract_vector_coordinates_and_mediabox(pdf_path, mark_navn)
				for t in range(len(pdf_path)):
					ok = run_postscript(vector_data[t], mark_navn, mark_koordinater_x, mark_koordinater_y, gs_code, gs_command, output_pdf, pdf_path[t], pdf_name[t], direction)
			except Exception as e:
				print("Error Ocurred, Check logs for details!")
				tids_stamp = date.today()
				f = open("error.log", "a")
				f.write(tids_stamp.strftime('%H:%M:%S - %d-%m-%y')+" - "+str(e)+"\n")
				f.close()
				break
			print("Processed")
		time.sleep(10)
		
def find_files(pdf_dir):
	names = []
	pdf_files = glob.glob(os.path.join(pdf_dir, "*.pdf"))
	pdf_names = [file for dirs in os.walk(pdf_dir, topdown=True)
                     for file in dirs[2] if file.endswith(".pdf")]
	for f in pdf_names:
		names.append(os.path.splitext(f)[0])

	return pdf_files, names

def extract_vector_coordinates_and_mediabox(pdf_path, mark_navn):
	
	vector_data = []
	
	for f in pdf_path:
		doc = fitz.open(f)


		for page_num in range(len(doc)):
			page = doc.load_page(page_num)
			mediabox = page.mediabox
			path_list = page.get_drawings()
			
			page_data = {
				"page_num": page_num + 1,
				"mediabox": mediabox[3],
				"mediabox2": mediabox[0],
				"mediabox1": mediabox[1],
				"paths": []
			}
			
			for path in path_list:
				path_data = {"path": path, "items": []}
				for item in path["items"]:
					if "Mark" in path['layer']:
						mark_navn.append(path['layer'])
						if item[0] == "l":  # line
							path_data["items"].append({
								"type": "line",
								"from": item[1],  # Starting coordinate
								"to": item[2]     # Ending coordinate
							})
					# Add more conditions for other types of vector graphics
				page_data["paths"].append(path_data)
			
			vector_data.append(page_data)

	return vector_data
	
def run_postscript(vector_data, mark_navn, mark_koordinater_x, mark_koordinater_y, gs_code, gs_command, output_pdf, pdf_path, pdf_name, direction):
	today = date.today()
	y = 0
	return_var = True
	#for page in vector_data:
	gs_code_temp = gs_code
	gs_command_temp = gs_command.copy()
	for path in vector_data["paths"]:
		for item in path["items"]:
			matchX0 = re.match("^.*\((.*)\,.*$",str(item['from']))
			matchY0 = re.match("^.*\,(.*)\).*$",str(item['to']))
			matchX1 = re.match("^.*\((.*)\,.*$",str(item['to']))
			matchY1 = re.match("^.*\,(.*)\).*$",str(item['from']))
			
			mark_koordinater_x.append(vector_data['mediabox2']+float(matchX0.group(1)))
			mark_koordinater_y.append(vector_data['mediabox']-float(matchY1.group(1)))
			
			#Calculate vector and angles between points and given vector (0,1)
			vectorx = float(matchX1.group(1)) - float(matchX0.group(1))
			vectory = float(matchY1.group(1)) - float(matchY0.group(1))
			
			vector_dot = (1 * vectorx) + (0 * vectory)
			
			vector_determinant = (0 * vectorx) - (1 * vectory)
			result_radian = math.atan2(vector_determinant, vector_dot)
			
			if result_radian < 0:
				result_degree = 360 + (result_radian * (180/math.pi))
			else:
				result_degree = result_radian * (180/math.pi)
						
			if 135 > result_degree > 45:
				direction.append("down")
			elif 315 > result_degree > 225:
				direction.append("up")
			elif 225 > result_degree > 135:
				direction.append("left")
			else:
				direction.append("right")				
				
	for i in range(len(mark_koordinater_x)):
		
		gs_code_temp +=	f"""
						getX{i} getY{i} Get_Text{i} Direction{i} center
						getX{i} getY{i} Get_Text{i} Direction{i} mark
						"""
		
		
		gs_command_temp.append(f"-dgetX{i}={mark_koordinater_x[i]}")
		gs_command_temp.append(f"-dgetY{i}={mark_koordinater_y[i]}")
		gs_command_temp.append(f"-sDirection{i}={direction[i]}")
		
		
		if mark_navn[i] == "Mark1":
			gs_command_temp.append(f"-sGet_Text{i}=In Red")
			y += 1
		elif mark_navn[i] == "Mark2":
			gs_command_temp.append(f"-sGet_Text{i}=In Yellow")
			y += 1
		elif mark_navn[i] == "Mark3":
			gs_command_temp.append(f"-sGet_Text{i}=In Blue")
			y += 1
		elif mark_navn[i] == "Mark4":
			gs_command_temp.append(f"-sGet_Text{i}=In Green")
			y += 1
	
	
	print("Marks inserted: " + str(y))
	gs_command_temp.extend([f"-sOutputFile={output_pdf}\example_output.pdf"])
		
	gs_command_temp.extend(["temp.ps", f"{pdf_path}"])
	gs_code_temp += """
						grestore
						true
						} ifelse
						} bind
					>> setpagedevice
						"""

	with open("temp.ps", "w") as ps_file:
		ps_file.write(gs_code_temp)
	
		
	subprocess.run(gs_command_temp, check=True)
	mark_koordinater_x.clear()
	mark_koordinater_y.clear()
	direction.clear()
	return return_var			
	

if __name__ == '__main__':
    main()
