<%
Function gen_key(digits)
dim char_array(50)
char_array(0) = "0"
char_array(1) = "1"
char_array(2) = "2"
char_array(3) = "3"
char_array(4) = "4"
char_array(5) = "5"
char_array(6) = "6"
char_array(7) = "7"
char_array(8) = "8"
char_array(9) = "9"
randomize
do while len(output) < digits
cnum = char_array(Int((10 - 0 + 1) * Rnd + 0))
output = output + cnum
loop
gen_key = output
End Function
%>