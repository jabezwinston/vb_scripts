' Author : Jabez Winston
'  
' Date : 3-12-2017

'Declare variables
DIM A,B,C
'Get values
A=INPUTBOX("Enter 1st Number")
B=INPUTBOX("Enter 2nd Number")
'Add them
C=INT(A)+INT(B)
'Print them
MSGBOX ("Sum of " + A + " and " + B + " is " + CSTR(C))