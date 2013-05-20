Summary:
	Create APIs to automate web applications launched by IE.


Dev. Environment:
	VB6.0

File/Folder Structure:
	Test - Contains main test scripts that provide script entry
	Component - Several reusable components defined in each vbs files that called in main test.
  	FunctionLibrary - Common functions reused in main test and component.
	Report - Custom output htm report files.
	help - help document to introduce the APIs.
	Excite.l022.cls - Souce code of OCX component "Excite.dll"
	ObjectRepository.xml - Describe and save the controls under test.
	RunManager.xls - Running template of the bulk of test script.Support multiple terminals as script execution machines,support automatic tracking running status.