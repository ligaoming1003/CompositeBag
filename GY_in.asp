<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>�˼�������ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>

<SCRIPT language=javascript  src="js/selectcity.js"></script>
<script language="JavaScript" src="js/validate.js" type="text/JavaScript"></script>

<SCRIPT language=javascript src="css/init.js"></SCRIPT>

<style type="text/css">
<!--
td {  font-family: "����"; font-size: 9pt}
body {  font-family: "����"; font-size: 9pt}
select {  font-family: "����"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "����"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "����"; font-size: 9pt} 
-->
</style>




<SCRIPT LANGUAGE=javascript>
<!--
function Juge(myform)
{

	
}


function SelectAll() {
	for (var i=0;i<document.selform.selBigClass.length;i++) {
		var e=document.selform.selBigClass[i];
		e.checked=!e.checked;
	}
}
//-->
</script>


<!--media=print ������Կ����ڴ�ӡʱ��Ч--> 
<style media=print> 
.Noprint{display:none;} 
.PageNext{page-break-after: always;} 
</style>
</HEAD>


<BODY topMargin=0 rightMargin=0 leftMargin=0>   
<% dim sql,rs

  select case request("action")
  case "add"
       call SaveAdd()
  case "modify"
       call SaveModify()
  case "del"
       call delCate()
  case "edit"
       isEdit=True
       call myform(isEdit)
  case "view"
       isEdit=True
       call myform(isview)
  case else
       isEdit=False
       call myform(isEdit) 
  end select

  
 
  
 

  %> <% sub myform(isEdit) %> 	 

 

	
	<%
	    set rs=server.createobject("adodb.recordset")
	   if isedit then
		   rs.open "select * from cailiao_in_store where id=" & request("id"),conn,1,1
            id=rs("id")
                if rs("fukuang")="0" then
                num=rs("jinge")
                else
                num=rs("fukuang")
                end if
	   end if %>    

        <table width="700" border="0" cellpadding="0" style="border-collapse: collapse" >

 <tr>
 <td align=center valign=middle height=107><b><font size="5">�ͻ�����ƾ֤</font></b><hr width="200" ></td>
</tr>

<tr>
<td>
<table width="700" border="0" cellpadding="0" style="border-collapse: collapse">



<tr>
<td height="30" align=center width="62%" colspan="4"><p align="left">
<b>
<font size="4">&nbsp;&nbsp;&nbsp;&nbsp; </font><font size="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ���ռ�����װ���޹�˾</font></b><font size="3"><b>��ͻ���<%=rs("gonghuodanwei")%></b></font><p align="left">
<font size="3">&nbsp;&nbsp;&nbsp; </font><b>
<font size="4">&nbsp;&nbsp;&nbsp;&nbsp; ��˫����ʵ����ֹ<%=rs("uptime")%>��</b>��ǰ��������<b><%=rs("pinming")%></b></font></td>
</tr>

<tr>
<td height="30" align=left width="100%" colspan="4"  >
<p><b><font size="2">&nbsp; </font></b>
<font size="3"></font><p><font size="3">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font>
<b><font size="4">�����</font></b><font size="3">(��д)
<b>
<%     
  dim   num   'Ҫת���ɴ�д�Ľ��     
  dim   atoc   'ת��֮���ֵ     
  Dim   String1   '���¶���     
  Dim   String2   '���¶���     
  Dim   String3   '��ԭAֵ��ȡ����ֵ     
  Dim   I   'ѭ������     
  Dim   J   'A��ֵ����100���ַ�������     
  Dim   Ch1   '���ֵĺ������     
  Dim   Ch2   '����λ�ĺ��ֶ���     
  Dim   nZero   '����������������ֵ�Ǽ���     

  String1   =   "��Ҽ��������½��ƾ�"     
  String2   =   "��Ǫ��ʰ��Ǫ��ʰ��Ǫ��ʰԪ�Ƿ�"     
  nZero   =   0     
    
  If   InStr(1,   CStr(num   *   100),   ".")   <>   0   Then     
  err.Raise   5000,   ,   "�˺���(   AtoC()   )ֻ��ת��С���������λ���ڵ�����"     
  End   If     
    
  J   =   Len(CStr(num   *   100))     
  String2   =   Right(String2,   J)   'ȡ����Ӧλ����STRING2��ֵ     
    
  For   I   =   1   To   J     
  String3   =   Mid(num   *   100,   I,   1)   'ȡ����ת����ĳһλ��ֵ     
    
  If   I   <>   (J   -   3)   +   1   And   I   <>   (J   -   7)   +   1   And   I   <>   (J   -   11)   +   1   And   I   <>(J   -   15)   +   1   Then     
  If   String3   =   0   Then     
  Ch1   =   ""     
  Ch2   =   ""     
  nZero   =   nZero   +   1     
  ElseIf   String3   <>   0   And   nZero   <>   0   Then     
  Ch1   =   "��"   &   Mid(String1,   clng(String3)   +   1,   1)     
  Ch2   =   Mid(String2,   I,   1)     
  nZero   =   0     
  Else     
  Ch1   =   Mid(String1,   clng(String3)   +   1,   1)     
  Ch2   =   Mid(String2,   I,   1)     
  nZero   =   0     
  End   If     
  Else   '��λ�����ڣ��ڣ���Ԫλ�ȹؼ�λ     
  If   String3   <>   0   And   nZero   <>   0   Then     
  Ch1   =   "��"   &   Mid(String1,   clng(String3)   +   1,   1)     
  Ch2   =   Mid(String2,   I,   1)     
  nZero   =   0     
  ElseIf   String3   <>   0   And   nZero   =   0   Then     
  Ch1   =   Mid(String1,   clng(String3)   +   1,   1)     
  Ch2   =   Mid(String2,   I,   1)     
  nZero   =   0     
  ElseIf   String3   =   0   And   nZero   >=   3   Then     
  Ch1   =   ""     
  Ch2   =   ""     
  nZero   =   nZero   +   1     
  Else     
  Ch1   =   ""     
  Ch2   =   Mid(String2,   I,   1)     
  nZero   =   nZero   +   1     
  End   If     
    
  If   I   =   (J   -   11)   +   1   Or   I   =   (J   -   3)   +   1   Then   '�����λ����λ��Ԫλ�������д��     
  Ch2   =   Mid(String2,   I,   1)     
  End   If     
    
  End   If     
  AtoC   =   AtoC   &   Ch1   &   Ch2     
    
  If   I   =   J   And   String3   =   0   Then   '���һλ���֣�Ϊ0ʱ�����ϡ�����     
  AtoC   =   AtoC   &   "��"     
  End   If     
    
  Next     
  if   num=0   then     
  atoc="��Ԫ��"     
  end   if    
response.write atoc 
  %>&nbsp;&nbsp;&nbsp;&nbsp;��:<u><%=Formatnumber((num),2,-1,,0)%>Ԫ</u></b></font><p>��<p align="center"><font size="2"><b>�� �� �ˣ�<%=name%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �ͻ���ǩ�֣���_______________</b></font>&nbsp;&nbsp;&nbsp;</td>




</font>

</td>
</tr>




</table>
<%end sub%>


</td>
</tr>


	</table> 			
	<table><tr><td height="30" colspan="8" align="center"><OBJECT id=WebBrowser classid=CLSID:8856F961-340A-11D0-A96B-00C04FD705A2 height=0 width=0 VIEWASTEXT> 
</OBJECT> 
<input type=button value=��ӡ onClick="document.all.WebBrowser.ExecWB(6,1)" class="NOPRINT"> 
<input type=button value=ֱ�Ӵ�ӡ onClick="document.all.WebBrowser.ExecWB(6,6)" class="NOPRINT"> 
<input type=button value=ҳ������ onClick="document.all.WebBrowser.ExecWB(8,1)" class="NOPRINT"> 
<input type=button value=��ӡԤ�� onClick="document.all.WebBrowser.ExecWB(7,1)" class="NOPRINT">
</td></tr>

				
				</table>    
		
</BODY></HTML>