<%
if session("ok")="yes" then
response.write("function killErrors() {return true;}")
response.write("window.onerror = killErrors;")
response.write("function show(){pop=window.open('login.asp','','width=270,height=150,top=200,left=300');")
response.write("if(pop){if((pop.document.body.clientWidth>280)||(pop.document.body.clientHeight>160)){")
response.write("pop.close();")
response.write("window.showModalDialog(""about:<""+""script language=javascript>window.open('login.asp','','width=270,height=150,top=200,left=300');window.close();""+""</""+""script>"","""",""dialogWidth:0px;dialogHeight:0px"");")
response.write("}}}show();")
else
response.Redirect("../index.asp")
end if
%>

