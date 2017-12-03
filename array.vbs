' Author : Jabez Winston
'  
' Date : 3-12-2017

DIM test_array(2)
test_array(0) = "Winst"
test_array(1) = "on"
wscript.echo test_array(0) + test_array(1)

DIM test_array2(2,3) 
test_array2(0,1) = "Hello"
test_array2(1,2) = "World"
wscript.echo test_array2(0,1) + " " + test_array2(1,2)