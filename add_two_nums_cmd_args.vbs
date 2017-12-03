' Author : Jabez Winston
'  
' Date : 3-12-2017

'Run as add_two_nums_cmd_args.vbs <num1> <num2>
Set args = WScript.Arguments 
'Declare variables
DIM A,B,C
'Get values
A=args(0)
B=args(1)
'Add them
C=INT(A)+INT(B)
'Print them
MSGBOX ("Sum of " + A + " and " + B + " is " + CSTR(C))