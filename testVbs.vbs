'The Purpose of this code is to take the sum of two numbers and add them'
'Please note to run this code you can not use a browswer unless isinside a html script'


DIM number1,number2,number3,number4,sum1,sum2,result 
'Variables are declared using the dim keyword. Multiple can be comma seperated'

call sum_of_two()
'Call the Function'

Function sum_of_two()
	number1 = CInt(InputBox("Please enter your 1st number:", "1st Number"))
	Msgbox number1
	number2 = CInt(InputBox("Please enter your 2nd number:", "2nd Number"))
	Msgbox number2
	number3 = CInt(InputBox("Please enter your 3rd number:", "3rd Number"))
	Msgbox number3
	number4 = CInt(InputBox("Please enter your 4th number:", "4th Number"))
	Msgbox number4
	sum1= number1+number2
	Msgbox sum1
	sum2= number3+number4
	Msgbox sum2
	result = sum1+sum2
	Msgbox "The Result is" + Cstr(result)
End Function
'Defines the function'
'set value of variable equal to InputBox for user input.'