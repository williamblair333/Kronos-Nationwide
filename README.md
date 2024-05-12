# Kronos-Nationwide

This VBA script will do the following:
- save the kronos 401a xls as an xlsm
- import vba code into new xlsm file and create a module
- delete the first 7 rows
- merge rows that have the same SS# (Should only be two in this environment)
	- add data (money) contained in either Column G or H (Only one of them) \
	  with data (money) found in either Column I, J, or K (Only one of them).
	  
- Column G & H totals needs to have this formula applied. =MIN(Column*100%,125)
- clear contents of remaining rows after last entry
- merge two remaining columns together
- save the kronos export xlsm file into a csv

TODO:
Refine the VBA so we can manipulate variables and cells without using .Select or .Activate 
https://stackoverflow.com/questions/10714251/how-to-avoid-using-select-in-excel-vba/10717999#10717999
