import telnetlib
import time
import xlsxwriter
from datetime import datetime
import argparse
import sys

parser = argparse.ArgumentParser(description="Script to calculate the span losses of links between NEs with given IPs")
parser.add_argument('--filename', type=str, help='Please Enter the name of text file that contains a list of nodes loopback IPs each in a line.')
parser.add_argument('--version', action='version', version='SpanLoss Calculator 1.0')
args = parser.parse_args()
run_time = datetime.now().strftime('%Y-%m-%d_%H%M%S')
cmd_file = "inputTelnet" #intial telnet commands file name to extract the inventory of NE's and discover amplifiers. 
filename = 'network_ip'
clean_ip_list=[]
with open (f"{filename}.txt",'r') as y: #Reading the ip's in network IP file and storing them in a list
	ip_list=y.readlines()
	for j,ip in enumerate(ip_list):
		ip_list[j]=ip.strip() #strip the line to remove \n
for ip in ip_list:
	if ip =='': continue
	else: clean_ip_list.append(ip)
print (f"The IP list is {clean_ip_list}")

'''
Telnet Function to connect to nodes.
'''
def open_telnet_conn(ip, cmd_file):
	x=False
	username="admin"
	password="Dil@pss8"
	cli="cli"
	port = 23
	connection_timeout = 5
	reading_timeout = 5
	connection = telnetlib.Telnet(ip,port,connection_timeout)
	node_output = connection.read_until(b'Username:', connection_timeout) #username and passwords should be entered bit-encoded.
	#print(node_output)
	connection.write(cli.encode('ascii') + b'\n')
	node_output = connection.read_until(b'Username: ', reading_timeout)
	connection.write(username.encode('ascii') + b'\n')
	#print(node_output)
	node_output = connection.read_until(b"Password:", reading_timeout)
	#print(node_output)
	connection.write(password.encode('ascii') + b"\n")
	node_output = connection.read_until(b'(Y/N)?', reading_timeout)
	connection.write(b'y\n')
	time.sleep(2)
	with open (f"{cmd_file}.txt",'r') as i: #read commands from cmd file
		cmd = i.readlines()
		for line in cmd:
			connection.write(line.encode('ascii'))
			time.sleep(2)
	node_output = connection.read_very_eager()
	with open (f"output_{ip}.txt",'wb') as f:
		f.write(node_output)
	print(type(node_output))
	print (node_output.decode('ascii'))
	connection.close()
	x=True
	return x

def write_intial_commands():
	with open (f'{cmd_file}.txt','w') as c:
		c.write(f'show general name'+'\n')
		c.write(f'paging status disabled'+'\n')
		c.write(f'show interface topology *'+'\n')
		c.write('y'+'\n')
		c.write('y'+'\n')
		c.write('y'+'\n')
		c.write('y'+'\n')
	with open (f'{cmd_file}2.txt','w') as c:
		c.write(f'paging status disabled'+'\n')
		c.write(f'show card inventory *'+'\n')
		
amplifiers = [] #A list to store the slots of Amplifiers
raman_type=["8DG60567AA", "8DG60567AB", "8DG64137AA"] #"8DG60567AAAA03"
dict_connectivity={}
dict_output_power={}
dict_input_power={}
dict_input_power_osc={}
dict_output_power_osc={}
spanLoss=[]
spanLoss_osc=[]
raman_amp=[]
dict_raman={}
dict_connectivity_keys=[]
networkmap={}
dict_design_values={}
key_error_list=[]
osc_calc=[]
fiber_cut=[]
raman_osc_dict={}

def get_raman_amplifiers(ip):
	raman_amp=[]
	with open (f"output_{ip}.txt",'r') as f: # open the output file from telnet function for site inventory and topology
		siteA = f.readlines()
		for line in siteA:
			for raman in raman_type: #search for raman amplifiers in inventory
				if raman in line:
					raman_amp.append((line.lstrip()).split()[0])
					#print (f'the raman is {raman_amp}')
	return raman_amp
	
def get_amplifiers(ip):
	amplifiers=[]
	with open (f"output_{ip}.txt",'r') as f: # open the output file from telnet function for site inventory and topology
		siteA = f.readlines()
		for line in siteA:
			if 'Ext' in line: #search for topolgies with external connection.
				amplifiers.append((line.lstrip()).split()[0])  #strip the spaces in the beginning of the line, take the first element in line (amplifier slot)
				#print (amplifiers)
			if 'System Name' in line:
				node_name = (line.split(':')[-1]).lstrip()
				networkmap[ip] = node_name
	return amplifiers, networkmap
	
def write_commands(ip,amplifier,role):
	amplifier_slot = amplifier.split('/')[0:2]
	amplifier_slot = '/'.join(amplifier_slot)
	if role == 'conn':
		with open (f'cmd_{ip}.txt','w') as c:
			c.write(f'show interface topology {amplifier}'+'\n')
	if role == 'rx':
		with open (f'cmd_{ip}.txt','w') as c:
			c.write(f'show interface {amplifier} detail'+'\n')
			c.write(f'config interface {amplifier_slot}/osc pm opr baseline newsystem'+'\n')
			c.write(f'show interface {amplifier_slot}/osc pm opr baseline'+'\n')
			#c.write(f'config powermgmt egress {amplifier_slot} spanlossdefault'+'\n')
	if role == 'tx':		
		with open (f'cmd_{ip}.txt','w') as c:
				c.write(f'show interface {amplifier} detail'+'\n')
				c.write(f'config interface {amplifier_slot}/osc pm opt baseline newsystem'+'\n')
				c.write(f'show interface {amplifier_slot}/osc pm opt baseline'+'\n')

def discover_connectivity(ip,amplifier):
	with open (f"output_{ip}.txt",'r') as f:
		siteA = f.readlines()
		for line in siteA:
			if "To Destination" in line:
				connectivity=line.split(' : ')[-1].strip()
				connectivity = connectivity.split('/')
				if connectivity[-1]=='1': connectivity[-1]='LINEIN'
				if connectivity[-1]=='4': connectivity[-1]='LINE'
				connectivity='/'.join(connectivity)
				dict_connectivity[f"{ip} {amplifier}"] = connectivity
				#print(dict_connectivity)

def get_design_values(ip,amplifier):
	with open (f"output_{ip}.txt",'r') as f:
		siteA = f.readlines()
		for line in siteA:
			if "Powermgmt SpanLossOut" in line:
				design_value=float(line.split(' ')[2].strip())
				dict_design_values[f"{ip} {amplifier}"] = design_value
				#print(f"the design values are: {dict_design_values}")
			if "No egress IRoadmf or IRoadmv or IRoadm9m or IR9 or IRoadm9r or IRoadm20 or IRoadm32 or Iroadm32l amplifier card." in line:
				dict_design_values[f"{ip} {amplifier}"] = 'Design Value Not Supported'
			else: 
				dict_design_values[f"{ip} {amplifier}"] = 'Design Value Not Supported'
			#print(dict_design_values)
	
def analyze_receive_osc_ports(ip,amplifier):
	with open (f"output_{ip}.txt",'r') as f:
		siteA = f.readlines()
		for line in siteA:
			#if 'LINEIN' in amplifier:	
			if ("Supvy In Power" in line):
				#print (line)
				rx_siteA_osc=(line.split(':')[-1]).strip()
				if ((rx_siteA_osc == '-') | (rx_siteA_osc == 'nil')): 
					rx_siteA_osc = '-99'
				rx_siteA_osc_num = float(rx_siteA_osc.split(' ')[0])
				dict_input_power_osc[f"{ip} {amplifier}"]= rx_siteA_osc_num
				if ((rx_siteA_osc == 'nil') | (rx_siteA_osc == '-99') | (rx_siteA_osc == '')): 
					rx_siteA_osc = '-99'
					fiber_cut.append(f"{ip} {amplifier}")
			else:
				if "Baseline Value" in line:
					#print (line)
					rx_siteA_osc = (line.split(':')[-1]).strip()
					if ((rx_siteA_osc == '-') | (rx_siteA_osc == 'nil')): 
						rx_siteA_osc = '-99'
					rx_siteA_osc_num = float(rx_siteA_osc.split(' ')[0])
					dict_input_power_osc[f"{ip} {amplifier}"]= rx_siteA_osc_num
					if ((rx_siteA_osc == 'nil') | (rx_siteA_osc == '-99')): 
							rx_siteA_osc = '-99'
							fiber_cut.append(f"{ip} {amplifier}")
							print(f"fiber cut is: {fiber_cut}")
		print("the input osc power is: " + str(dict_input_power_osc))
	
def analyze_receive_line_ports(ip,amplifier):
	with open (f"output_{ip}.txt",'r') as f:
		siteA = f.readlines()
		for line in siteA:
			if 'LINEIN' in amplifier:
				#if "Ingress OA Total Input Power" in line: #search for total input power.
				if ("Total Power In" in line) | ("Ingress OA Total Input Power" in line):
					#print(line)
					rx_siteA=(line.split(':')[-1]).strip() #split the line with : and get the last element in the list.
					
					if ((rx_siteA == 'nil') | (rx_siteA == 'Off') | (rx_siteA == '-')): 
						rx_siteA = '-99'
					rx_siteA_num = float(rx_siteA.split(' ')[0]) #split the value to remove dBm and get the first element and convert it to float.
					dict_input_power[f"{ip} {amplifier}"]= rx_siteA_num
				#print("the input power is: " + str(dict_input_power))

def analyze_transmit_osc_ports(ip,amplifier):
	with open (f"output_{ip}.txt",'r') as f:
		siteA = f.readlines()
		for line in siteA:
			#if 'LINEOUT' in amplifier:	
			if ("Supvy Out Power" in line) | ("OSC Out Power" in line):
				#print (line)
				tx_siteA_osc=(line.split(':')[-1]).strip()
				if ((tx_siteA_osc == '-') | (tx_siteA_osc == 'nil')): tx_siteA_osc = '99'
				tx_siteA_osc_num = float(tx_siteA_osc.split(' ')[0])
				dict_output_power_osc[f"{ip} {amplifier}"]= tx_siteA_osc_num
				if ((tx_siteA_osc == 'nil') | (tx_siteA_osc == '99')): 
						tx_siteA_osc = '99'
						fiber_cut.append(f"{ip} {amplifier}")
			else:
				if "Baseline Value" in line:
					#print (line)
					tx_siteA_osc = (line.split(':')[-1]).strip()
					if (tx_siteA_osc == '-'): tx_siteA_osc = '99'
					tx_siteA_osc_num = float(tx_siteA_osc.split(' ')[0])
					dict_output_power_osc[f"{ip} {amplifier}"]= tx_siteA_osc_num
					if ((tx_siteA_osc == 'nil') | (tx_siteA_osc == '99')): 
						tx_siteA_osc = '99'
		print(f"Output OSC dict is:{dict_output_power_osc}")
			
def analyze_transmit_line_ports(ip,amplifier):			
	with open (f"output_{ip}.txt",'r') as f:
		siteA = f.readlines()
		for line in siteA:
			if 'LINEOUT' in amplifier:
				if ("Total Power Out" in line) | ("Egress OA Total Output Power" in line): #search for total output power
					#print(line)
					tx_siteA=(line.split(':')[-1]).strip()
					if ((tx_siteA == 'nil') | (tx_siteA == 'Off') | (tx_siteA == '-')):
						tx_siteA = '99'
						osc_calc.append(f"{ip} {amplifier}")
							
					tx_siteA_num = float(tx_siteA.split(' ')[0])
					dict_output_power[f"{ip} {amplifier}"] = tx_siteA_num
				print("the output power is: " + str(dict_output_power))
				
def raman_process(ip,raman_amp):
	for raman in range(len(raman_amp)):
		with open (f'cmd_{ip}.txt','a') as c:
			c.write(f'show interface {raman_amp[raman]}/LINEIN detail'+'\n')
			c.write(f'show interface topology {raman_amp[raman]}/LINEOUT'+'\n')

	finished = open_telnet_conn(ip,f"cmd_{ip}")
	if finished:
		with open (f"output_{ip}.txt",'r') as f:
			siteA = f.readlines()
			for line in siteA:
				if 'To Destination' in line:
					connectivity = line.split(':')[-1].split('-')[0].strip()
					raman_osc_dict[f'{ip} {raman_amp[raman]}/LINEIN'] = f'{ip} {connectivity}'
					#print(raman_osc_dict)
					
				if 'Operating Gain' in line:
					#print(line)
					raman_gain= (line.split(':')[-1]).strip() #split the line with : and get the last element in the list.
					#if (raman_gain == 'Off') : raman_gain = '0'
					raman_gain_num = float(raman_gain.split(' ')[0])
					dict_raman[f"{ip} {raman_amp[raman]}/LINEIN"]=raman_gain_num
					#print (f"The raman dictionary is {dict_raman}")
		raman_amp=[]

def calculate_line_spanloss(key,value):
	spanLoss_value = dict_output_power[key] - dict_input_power[value]
	if (value in raman_osc_dict.keys()): 
		spanLoss_value = spanLoss_value + dict_raman[value]
	return spanLoss_value

def calculate_osc_spanloss(key,value):
	if (value in raman_osc_dict.keys()):
		spanLoss_value_osc = dict_output_power_osc[key] - dict_input_power_osc[raman_osc_dict[value]]
		spanLoss_value_osc = spanLoss_value_osc + dict_raman[value]
	else:
		spanLoss_value_osc = dict_output_power_osc[key] - dict_input_power_osc[value]
	return spanLoss_value_osc

def create_excel_sheet():
	workbook = xlsxwriter.Workbook(f'result_{run_time}.xlsx')
	worksheet = workbook.add_worksheet("Span Losses")
	worksheet.write(0,0 , 'Connection')
	worksheet.write(0,1 , 'Span Loss - Line(dB)')
	#worksheet.write(0,2 , 'Design Value')
	worksheet.write(0,2 , 'Span Loss - OSC(dB)')
	worksheet.write(0,3 , 'Comments')
	
	for i,(key,value) in enumerate(dict_connectivity.items()):
		far_end_ip = value.split(' ')[0]
		far_end_port = value.split(' ')[-1]
		near_end_ip = key.split(' ')[0]
		near_end_port = key.split(' ')[-1]
		
		try:
			if far_end_ip in clean_ip_list:
				#if (key in dict_output_power.keys()) | (dict_input_power[value] in dict_input_power.keys()):
				spanLoss_value = calculate_line_spanloss(key,value)
				spanLoss_value_osc = calculate_osc_spanloss(key,value)
				#print(f"The span Loss is {spanLoss_value}")
				#if (dict_design_values[key] == 'Design Value Not Supported'): dict_design_values[dict_connectivity[key]]
				worksheet.write(i+1 ,0 , networkmap[near_end_ip] + " " + near_end_port + "<>" + networkmap[far_end_ip] + " " + far_end_port)
				worksheet.write(i+1 ,1 , spanLoss_value)
				#worksheet.write(i+1 ,2 , dict_design_values[key])
				worksheet.write(i+1 ,2 , spanLoss_value_osc)
				if key in osc_calc: worksheet.write(i+1 ,3 , "No Channels Passing through the link")
				if dict_connectivity[key] in fiber_cut: worksheet.write(i+1 ,3 , "Fiber Cut")
				#if (dict_design_values[key] == 'Design Value Not Supported'): dict_design_values[dict_connectivity[key]]
				# if (value in dict_output_power.keys()) | (key in dict_input_power.keys()):
					# spanLoss_value = calculate_line_spanloss(value,key)
					# spanLoss_value_osc = calculate_osc_spanloss(value,key)
				
				worksheet.write(i+1 ,0 , networkmap[near_end_ip] + " " + near_end_port + "<>" + networkmap[far_end_ip] + " " +far_end_port)
				worksheet.write(i+1 ,1 , spanLoss_value)
				#worksheet.write(i+1 ,2 , dict_design_values[key])
				worksheet.write(i+1 ,2 , spanLoss_value_osc)
				if key in fiber_cut: worksheet.write(i+1 ,3 , "Fiber Cut")
			else: 
				spanLoss_value = "Link is outside provided nodes list"
				worksheet.write(i+1 ,0 , networkmap[near_end_ip] + " " + near_end_port + "<>" + far_end_ip + " " +far_end_port)
				worksheet.write(i+1 ,1 , spanLoss_value)
		except KeyError:
			#print(key,value)
			key_error_list.append(key)
			key_error_list.append(value)
			continue
	workbook.close()


##############################################################################################################################	
def main():
	while True:
		write_intial_commands()
		for ip in clean_ip_list: #one main for loop for each ip
			ip.strip()
			finished = open_telnet_conn (ip,cmd_file) #call telnet function with initial commands to extract inventory for the first time only with cmd_file.
			if finished:
				amplifiers, networkmap = get_amplifiers(ip)
			finished = open_telnet_conn (ip,(cmd_file+'2'))
			if finished:
				raman_amp = get_raman_amplifiers(ip)
				print(raman_amp)
				
		# 	print(f"Start analyzing node {networkmap[ip]}....")
		# 	for amp in range(len(amplifiers)): #for loop to write the commands to grab the tx power and rx power for each port.
				
		# 		write_commands(ip,amplifiers[amp],role='conn')
		# 		finished = open_telnet_conn(ip,f"cmd_{ip}")
		# 		if finished:
		# 			discover_connectivity(ip,amplifiers[amp])
				
		# 		write_commands(ip,amplifiers[amp],role='rx')
		# 		finished = open_telnet_conn(ip,f"cmd_{ip}")
		# 		if finished:
		# 			analyze_receive_osc_ports(ip,amplifiers[amp])
		# 			analyze_receive_line_ports(ip,amplifiers[amp])
					
		# 		write_commands(ip,amplifiers[amp],role='tx')
		# 		finished = open_telnet_conn(ip,f"cmd_{ip}")
		# 		if finished:
		# 			analyze_transmit_osc_ports(ip,amplifiers[amp])
		# 			analyze_transmit_line_ports(ip,amplifiers[amp])
				
		# 	amplifiers = []
			
		# 	if raman_amp:
		# 		raman_process(ip,raman_amp)
		# 	raman_amp = []
			
		# #print(f'the connectivity is {dict_connectivity}')
		# print ("Creating Span Loss Report ...")
		# create_excel_sheet()
		# while True:
		# 	if key_error_list:
		# 		print(f"The following links need to be done again {key_error_list}")
		# 	user_end = input("Please enter any key to exit...")
		# 	if user_end == '': user_end='y'
		# 	if user_end:
		# 		sys.exit()

if __name__ == "__main__":
	main()