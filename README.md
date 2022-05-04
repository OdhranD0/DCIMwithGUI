# DCIMCreator

# Creates a .02m file from an excel sheet for use with Marcoms OPC to Modbus Converter with PSO.

## Installation

Uses python 3.7

The following libraries are required:
1. Pandas for excel
2. numpy
3. Tkinter


## Running Software
To run the software open the DCIM_Converter folder and run DCIM_Gui.exe

The software takes in an excel sheet in the correct format (see import template) and converts it to an o2m file to be used with marcom software.
Operates by writing to multiple text files then combining them.

Certatin files will not change and are located in layout files folder.

Latest Release for Marcom Software is 2.57.

## Template

The template is located in "Template" Folder. See below for an explanation of headings.

NOTE: Back slash replaced with word SLASH because of python error, code then changes it.

Tag: Tag from scada with 'cluster name'. prefixing it e.g. 'c1.'
Address: Modbus Address from Tag
Bit: Bit of address for real, int or long this will be 0.
DataTypeNo: See Data Type Nos below
DataType: Data Type name in caps
Client Handle: Increment starting from 1.
Sname: This alos increments from 100.
Node: Modbus Node (1 or 2)
SetupNames/Setup: Names of 2 nodes and their register size.


Data Type No ( Each data type has an associated data type number. See below)
Boolean: 11
Int: 2
Real 4
Long 3

#Odhran Doherty 2019
