import pandas as pd
import math
import sys
import os
from pandas import ExcelWriter
from pandas import ExcelFile

def program(arg1, arg2):
	#Output File
	f = open(arg1+"_"+arg2, "w")
	
	print("Name of Excelsheet = ",arg1)
	print("Name of the sheet  = ",arg2)
	print("Name of output file= ",arg1+"_"+arg2)
	#Input File
	df = pd.read_excel(arg1, sheet_name=arg2)
	
	print("Column headings:")
	print(df.columns)
	
	area_data=df['Normalized Area']
	latency_data=df['Normalized Latency']
	security_data=df['Security Level']
	power_data=df['Normalized Power']
	
	#Baseline value to indicate in graph and normalize the data
	base_sec_l=security_data[area_data[area_data == 1].index[0]]
	base_area=area_data[area_data[area_data == 1].index[0]]
	base_latency=latency_data[area_data[area_data == 1].index[0]]
	base_power=power_data[area_data[area_data == 1].index[0]]
	
	#If two baseline values, exit
	if(not(base_area == base_latency and base_area == 1.0)):
		print("Baseline values are not correct")
		exit(1)
	print("Baseline Security = ", base_sec_l)
	print("Baseline area= ", base_area)
	print("Baseline Latency = ", base_latency)
	print("Baseline Power = ", base_power)
	
	#Min and Max value for legends and graph axis min, max
	area_min = area_data.iloc[area_data.nonzero()].min()
	area_max = area_data.max()
	latency_min = latency_data.iloc[latency_data.nonzero()].min()
	latency_max = latency_data.max()
	power_min = power_data.iloc[power_data.nonzero()].min()
	power_max = power_data.max()
	
	print(area_min)
	print(area_max)
	print(latency_min)
	print(latency_max)
	print(power_min)
	print(power_max)
	#power grid to indicate power intensity by redness
	power_grid=[]
	for i in range(0,56):
		grid=(power_max-power_min)/56
		power_grid.append(power_min+i*grid)
	print(power_grid)
	
	#File write start
	f.write("%This code is generated from --> "+arg1+" sheet name - "+arg2+"\n")
	f.write("\documentclass{standalone}\n")
	f.write("\\usepackage{tikz}\n")
	f.write("\\usepackage{comment}\n")
	f.write("\\usepackage{pgfplots, pgfplotstable}\n")
	f.write("\\usetikzlibrary{math, decorations.pathreplacing,angles,quotes,bending, arrows.meta}\n")
	f.write("\pgfplotsset{compat=1.15,every tick label/.append style={font=\\tiny}}\n")
	f.write("\n")
	f.write("\pgfplotstableread{\n")
	f.write("X Y Z \n")
	f.write(repr(base_area)+"   "+repr(base_latency)+"   "+repr(base_sec_l)+"\n")
	f.write("}\\baseline\n")
	f.write("\n")
	counter1=65
	counter2=65
	tables=['datatablezero','datatable','datatabletwo','datatablethree','datatablefour', 'datatablefive']
	color=['gray','blue','yellow','red','brown','green']
	shape=['oplus','','diamond','triangle','square','pentagon']
	size=['3','3','3','3','3','3']
	myColor=['red!0','red!1','red!2','red!3','red!4','red!5','red!6','red!7','red!8','red!9','red!10','red!11','red!12','red!13','red!14','red!15','red!16','red!17','red!18','red!19','red!20','red!21','red!22','red!23','red!24','red!25','red!26','red!27','red!28','red!29','red!30','red!31','red!32','red!33','red!34','red!35','red!36','red!37','red!38','red!39','red!40','red!44','red!48','red!52','red!56','red!60','red!64','red!68','red!72','red!76','red!80','red!84','red!88','red!92','red!96','red!100']
	#For loop draws separate plot for each point(or row in excel
	for (i,j,k,l) in zip(area_data, latency_data, security_data,power_data):
		temp=[i,j,l]
		#reject the values from excel with 0 or NaN
		if(not all(temp) or i!=i or j!=j or k!=k):
			continue
		f.write("\pgfplotstableread{\n")
		f.write("X Y Z \n")
		f.write(repr(i)+"   "+repr(j)+"   "+repr(k)+"\n")
		f.write("}\\"+tables[int(k)]+chr(counter2)+chr(counter1)+"\n")
		if(counter1 == 90):
			counter1 = 65
			counter2 =counter2 + 1
		else:
			counter1=counter1+1
	f.write("\n")
	f.write("\makeatletter\n")
	f.write("        \pgfdeclareplotmark{dot}\n")
	f.write("        {%\n")
	f.write("            \\fill circle [x radius=0.02, y radius=0.08];\n")
	f.write("        }%\n")
	f.write("\makeatother\n")
	f.write("\n")
	f.write("\n")
	f.write("\pgfplotsset{\n")
	f.write("/pgfplots/colormap={autumn}{rgb255=(255,255,255) rgb255=(255,0,0) }\n")
	f.write("}\n")
	f.write("\n")
	f.write("\\begin{document}\n")
	f.write("\\begin{tikzpicture}[scale=1.5]\n")
	f.write("\\begin{axis}\n")
	f.write("    [   \n")
	f.write("    view={120}{40},\n")
	f.write("        width=220pt,\n")
	f.write("        height=220pt,\n")
	f.write("        grid=major,\n")
	f.write("        colorbar,\n")
	f.write("        z buffer=sort,\n")
	f.write("        xmin="+repr(area_min)+",xmax="+repr(area_max)+",\n")
	f.write("        ymin="+repr(latency_min)+",ymax="+repr(latency_max)+",\n")
	f.write("        zmin=0,zmax=5,\n")
	f.write("        enlargelimits=upper,\n")
	f.write("        xlabel style={sloped},\n")
	f.write("        xlabel={Normalized Area},\n")
	f.write("        legend style={at={(0.65,0.33)}, cells={align=left}, anchor=north,legend columns=1, font=\\tiny, fill opacity=0.7, text opacity=1, draw opacity=1},\n")
	f.write("                ylabel style={sloped},\n")
	f.write("        ylabel={Normalized Latency},\n")
	f.write("        zlabel={Security level},\n")
	f.write("        %point meta={x+y},\n")
	f.write("        point meta max = "+repr(power_max)+",\n")
	f.write("        point meta min = "+repr(power_min)+",\n")
	f.write("        colorbar style={\n")
	f.write("        title= \\tiny{Normalized Power},\n")
	f.write("        title style={\n")
	f.write("            text width=3em,       % Abstand yticks zu colorbar\n")
	f.write("        },\n")
	f.write("        at={(1.075,0.1)}, % Coordinate system relative to the main axis. (1,1) is upper right corner of main axis.\n")
	f.write("        anchor=south west,\n")
	f.write("        %/pgf/number format/precision=3,\n")
	f.write("        ytick={"+repr(power_min)+",")
	#Finding intermediate numbers for power graph label
	for i in range(math.ceil(power_min+0.15),math.ceil(power_max-0.15)):
		f.write(repr(i)+",")
	f.write(repr(power_max)+"},\n")
	f.write("        height=2/3*\pgfkeysvalueof{/pgfplots/parent axis height}, % Scale the colorbar relative to the main axis\n")
	f.write("        /pgf/number format/.cd, % Change the key directory to /pgf/number format\n")
	f.write("        %fixed, fixed zerofill, precision=1,\n")
	f.write("        /tikz/.cd  % Change back to the normal key directory\n")
	f.write("        }\n")
	f.write("    ]\n")
	f.write("\n")
	f.write("\n")
	
	counter1=65
	counter2=65
	
	#To draw the grpah using different colors
	for (i,j,k,l) in zip(area_data, latency_data, security_data,power_data):
		temp=[i,j,l]
		m = 0
		if(not all(temp) or i!=i or j!=j or k!=k):
			continue
		for m in range(0,56):
			if(l>=power_grid[m]):
				m_color=m
				grid_value = power_grid[m]
				power_color=myColor[m]
				#print("m = ",m," power_color = ",power_color)
		#print("security = ",k," power = ",l," color = ",power_color)
		# for debug - f.write(repr(l)+"  index = "+repr(m_color)+" grid value = "+repr(grid_value)+"\\addplot3[only marks,fill="+power_color+",mark="+shape[int(k)]+"*,mark size="+size[int(k)]+"] table {\\"+tables[int(k)]+chr(counter2)+chr(counter1)+"};\n");
		f.write("\\addplot3[only marks,fill="+power_color+",mark="+shape[int(k)]+"*,mark size="+size[int(k)]+"] table {\\"+tables[int(k)]+chr(counter2)+chr(counter1)+"};\n");
		if(counter1 == 90):
			counter1 = 65
			counter2 =counter2 + 1
		else:
			counter1=counter1+1
	#Baseline color and writing into file
	for m in range(0,56):
		if(base_power>=power_grid[m]):
			power_color=myColor[m]
	f.write("\\addplot3[only marks,fill="+power_color+",mark=*,mark size=5] table {\\baseline};\n")
	#Writing legend for only different security level
	f.write("\legend{")
	k_prev=50
	for (i,j,k,l) in zip(area_data, latency_data, security_data,power_data):
		temp=[i,j,l]
		if(not all(temp) or i!=i or j!=j or k!=k):
			continue
		if(k_prev != k):
			f.write("Security Level-"+repr(int(k))+",")
		else:
			f.write(",")
		k_prev=k
	f.write("Ref. Baseline}\n")
	security_data=security_data.dropna()
	security_data=security_data.drop_duplicates('first')
	f.write("\n")
	f.write("\end{axis}\n")
	f.write("\end{tikzpicture}\n")
	f.write("\n")
	f.write("\end{document}\n")
	
	f.close()

if __name__ == "__main__":
    try:
        arg1 = sys.argv[1]
        arg2 = sys.argv[2]
    except IndexError:
        print ("Usage: ",os.path.basename(__file__), " arguments are missing")
        sys.exit(1)

    # start the program
    program(arg1,arg2)
