Set A=CreateObject("QuickTest.Application")
A.Launch
A.Visible=True

Dim testarray(3)
testarray(1)="C:\Program Files\HP\QuickTest Professional\Tests\Jacksonville\Regular Expression"
testarray(2)="C:\Program Files\HP\QuickTest Professional\Tests\Jacksonville\Output Value"
testarray(3)="C:\Program Files\HP\QuickTest Professional\Tests\Jacksonville\String Function"

For i=1 to Ubound(testarray)
	A.Open testarray(i)

	Set qttest=A.Test
	qttest.Run,True
	qttest.Close
	Set qttest=Nothing

Next
A.Quit