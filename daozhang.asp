<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>�˼�������ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>
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
function Juge(powersearch)
{

		if (powersearch.Loan_num.value == "")
	{
		alert("��������Ϊ�գ�");
		powersearch.Loan_num.focus();
		return (false);
	}
	    


}


function SelectAll() {
	for (var i=0;i<document.selform.selBigClass.length;i++) {
		var e=document.selform.selBigClass[i];
		e.checked=!e.checked;
	}
}
//-->
</script>


</HEAD>
<BODY topMargin=0 rightMargin=0 leftMargin=0>       
<!--#include file="top.asp"-->
<%
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
  case else
       isEdit=False
       call myform(isEdit) 
  end select
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
  
sub SaveAdd
Loan_num=request.Form("Loan_num")
kehu=trim(request.Form("kehu"))
fanshi=trim(request.Form("fanshi"))
yinhang=trim(request.Form("yinhang"))
leixin=trim(request.Form("leixin"))
comm=gen_key(8)
set rs=server.createobject("adodb.recordset") 
sql="select * from daozhang"
rs.open sql,conn,1,3
rs.addnew
rs("kehu") = kehu
rs("uptime") = request.Form("uptime")
rs("fanshi") = fanshi
rs("Loan_num") = Loan_num
rs("yinhang") = yinhang
rs("leixin")=leixin
rs("comm") =comm
rs.update
rs.close
set rs=nothing

set rs1=server.createobject("adodb.recordset") 
sql1="select * from JS_out_store"
rs1.open sql1,conn,1,3
rs1.addnew
rs1("use_dep") = kehu
rs1("uptime") = request.Form("uptime")
rs1("hkdz") = Loan_num
rs1("comm") =comm
rs1("pinming")="�����"
rs1.update
rs1.close
set rs1=nothing


set rs2=server.createobject("adodb.recordset") 
sql2="select * from JS_daozhang"
rs2.open sql2,conn,1,3
rs2.addnew
rs2("kehu") = kehu
rs2("uptime") = request.Form("uptime")
rs2("fanshi") = fanshi
rs2("Loan_num") = Loan_num
rs2("yinhang") = yinhang
rs2("leixin")=leixin
rs2("comm") =comm
rs2.update
         response.Write "<script language=javascript>alert('�������ӳɹ���');</script>"
         response.write "<meta http-equiv=""refresh"" content=""0;url=daozhang.asp"">"
         response.end
rs2.close
set rs2=nothing

end sub

 
sub SaveModify  
Loan_num=request.Form("Loan_num")
old_num=request.Form("old_num")
comm=request.Form("old_comm")
old_kehu=request.Form("old_kehu")
kehu=trim(request.Form("kehu"))
fanshi=trim(request.Form("fanshi"))
yinhang=trim(request.Form("yinhang"))
leixin=trim(request.Form("leixin"))

num= old_num-Loan_num

        if old_kehu<>kehu then
         response.Write "<script language=javascript>alert('�Բ��𣡿ͻ��������޸�.n\n\ֻ���޸�����Ϊ����������룡');</script>"
         response.write "<meta http-equiv=""refresh"" content=""0;url=daozhang.asp"">"
         response.end
        end if
set rs=server.createobject("adodb.recordset") 
sql="select * from daozhang where id="&request.Form("id")
rs.open sql,conn,1,3
rs("kehu") = kehu
rs("uptime") = request.Form("uptime")
rs("fanshi") = fanshi
rs("Loan_num") = Loan_num
rs("yinhang") = yinhang
rs("leixin")=leixin
rs.update
rs.close
set rs=nothing

set rs1=server.createobject("adodb.recordset") 
sql1="select * from JS_out_store where use_dep='"&kehu&"'and  comm="&comm&""
rs1.open sql1,conn,1,3
rs1("use_dep") = kehu
rs1("uptime") = request.Form("uptime")
rs1("hkdz") = Loan_num
rs1.update
rs1.close
set rs1=nothing


set rs2=server.createobject("adodb.recordset") 
sql2="select * from JS_daozhang where kehu='"&kehu&"'and  comm="&comm&""
rs2.open sql2,conn,1,3
rs2("kehu") = kehu
rs2("uptime") = request.Form("uptime")
rs2("fanshi") = fanshi
rs2("Loan_num") = Loan_num
rs2("yinhang") = yinhang
rs2("leixin")=leixin
rs2.update
         response.Write "<script language=javascript>alert('���˻����޸ĳɹ���');</script>"
         response.write "<meta http-equiv=""refresh"" content=""0;url=daozhang.asp"">"
         response.end
rs2.close
set rs2=nothing


        

end sub   
   sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('��û��ѡ����¼��');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=daozhang.asp"">"
  response.end
else
        conn.execute("delete from daozhang where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('ɾ���ɹ���');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=daozhang.asp"">"
        response.end
end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="admin" or oskey="xiaoshou" then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from daozhang where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�༭�����</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">��ӻ����</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��������</td>
            <td width="30%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			�ͻ�����</td>
			
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��������</td>
			
			<td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">����</td>
			
			<td width="15%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">���ʷ�ʽ</td>
			
			<td width="15%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��������</td>
			
          </tr>
		  <form name="powersearch" method="post" action="daozhang.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"> 
			
<select name='kehu' size="1" style='width:202; height:21'>
			<% if isedit then%>
				<option selected="selected"><%=rs("kehu")%></option>

			<%else%>
		        <option selected value='0000'>��ѡ��ͻ�</option>
			<%end if%>
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from kehu  order by kehu",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		   kehu=rs2("kehu")
		   %>
		   <option value="<%=rs2("kehu")%>"><%=rs2("kehu")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%>			</select></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='leixin'  size="1" style='width:59; height:17'>
			<% if isedit then%>
			
			<option selected value='����'>����</option>
            <option selected value='����'>����</option>
            <option selected value='�ֿ�'>�ֿ�</option>
            <option selected value='����'>����</option>
            <option selected="selected"><%=rs("leixin")%></option>
            <input type="hidden" value="<%=rs("Loan_num")%>" name="old_num">
            <input type="hidden" value="<%=rs("comm")%>" name="old_comm">
            <input type="hidden" value="<%=rs("kehu")%>" name="old_kehu">
			<%else%>
			<option selected value='����'>����</option>
            <option selected value='����'>����</option>
            <option selected value='�ֿ�'>�ֿ�</option>
            <option selected value='����'>����</option>
            <option selected value=''>��ѡ��</option>
			<%end if%>

			
			
			</select></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="Loan_num" type="text"  size="19" value='<% if isedit then
		                                                         response.write rs("Loan_num")
																  end if %>'></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">

			<select name='fanshi'  size="1" style='width:79; height:14'>
			<% if isedit then%>
			
			<option selected value='�ֽ�'>�ֽ�</option>
            <option selected value='ת��'>ת��</option>
            <option selected value='�ж�'>�ж�</option>
            <option selected="selected"><%=rs("fanshi")%></option>
			<%else%>
			<option selected value='�ֽ�'>�ֽ�</option>
            <option selected value='ת��'>ת��</option>
            <option selected value='�ж�'>�ж�</option>
            <option selected value=''>��ѡ��</option>
			<%end if%>

			
			
			</select></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">

			<select name='yinhang' size="1" style='width:100; height:21'>
			<% if isedit then%>
				<option selected="selected"><%=rs("yinhang")%></option>

			<%else%>
		        <option selected value='0000'>��ѡ����Դ</option>
			<%end if%>
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from yinhang order by yinhang",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		   yinhang=rs2("yinhang")
		   %>
		   <option value="<%=rs2("yinhang")%>"><%=rs2("yinhang")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%>			</select></td>
          </tr>
			<tr>
			<td colspan="6" align="center"><input type="submit" value="<% if isedit then
		                                                         response.write "�༭"
																 else
																 response.write "���"
																  end if %>"> </td>
          </tr></form>
		  
           </table>
        <%end if%>


</td>
</tr>
</table>
<br><br>
<TABLE width=820 align=center border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6" height="143">
<tr>
<td >



			<form action="daozhang_out.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp;�� �� �� �� �� ��</font></b>

</td>
</tr>
</table>
	  
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			��������</td>
			<td width="21%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			��&nbsp;&nbsp; ��</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			��&nbsp; ��</td>
			<td width="9%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			���˷�ʽ</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			��������</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			����</td>
            <td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�� ��</td>
          </tr>
		 <%
PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 20          'ÿҳ��ʾ������
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
set rs=server.CreateObject("ADODB.RecordSet") 
search=request("search")
if search<>"" then 
keyword=request("keyword")

rs.Source="select* from daozhang where kehu like '%"&keyword&"%' order by -id"
else
rs.Source="select* from daozhang  order by -id"
end if
rs.Open rs.Source,conn,1,1
if not rs.EOF then
rs.PageSize     = MyPageSize
MaxPages         = rs.PageCount
rs.absolutepage = MyPage
total            = rs.RecordCount
for i=1 to rs.PageSize
if not rs.EOF then
%>
<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">
<td height="18" align="center"><%=rs("uptime")%></td>
<td height="18" align="center">&nbsp;<a href="daozhang.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("kehu")%></a></td>
<td height="18" align="center"><%=Formatnumber(rs("Loan_num"),2,-1,0,0)%><%if rs("Loan_num")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("fanshi")%><%if rs("fanshi")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("yinhang")%><%if rs("yinhang")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("leixin")%><%if rs("leixin")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="center" colspan="7" height="22">
<p align="right">�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
			if search<>"" then
			url="daozhang.asp?keyword="&keyword&"&search=search&"
			else
            url="daozhang.asp?"
			end if				
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "page=" & PagePrev & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a> "
Response.write "<a class=black href='" & Url & "page=1' title='��1ҳ'>ҳ��</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "page=" & Prev_Page & "' title='��" & Prev_Page & "ҳ'>��һҳ</a> "
end if

if Maxpages>=PageNextSize*PageShowSize then
PageSizeShow = PageShowSize
Else
PageSizeShow = Maxpages-PageShowSize*(PageNextSize-1)
End if
If PageSizeShow < 1 Then PageSizeShow = 1
for PageCounterSize=1 to PageSizeShow
PageLink = (PageCounterSize+PageNextSize*PageShowSize)-PageShowSize
if PageLink <> MyPage Then
Response.write "<a class=black href='" & Url & "page=" & PageLink & "'>[" & PageLink & "]</a> "
else
Response.Write "<B>["& PageLink &"]</B> "
end if
If PageLink = MaxPages Then Exit for
Next

if Mypage+1 <=Pagetpage  then
Next_Page = MyPage + 1
Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='��" & Next_Page & "ҳ'>��һҳ</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='��"& Pagetpage &"ҳ'>ҳβ</A>"
Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='��" & PageShowSize & "ҳ'>��һ��ҳ</a>"
End if
rs.Close
set rs=nothing
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="admin" or oskey="xiaoshou" then%>
        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> ѡ��/��ѡ
              <input type=submit value=��ӡ name=action> 
<%end if%></td></tr>
        </table>

	  
</td>
    </tr>

  </table>


</form>

</td>
</tr>
</table>
<br><br>
<TABLE width=820 align=center border="1" cellspacing="20" cellpadding="0" bordercolor="#0055E6" height="92">
<tr>
<td>
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">��������</font></b>

</td>
</tr>
</table>

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="daozhang.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">���ͻ����ң� 
				<select name="keyword" size="1">
			<option value="">��ѡ��</option>
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(kehu) from daozhang order by kehu",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		   kehu=rs2("kehu")
		   %>
		   <option value="<%=rs2("kehu")%>"><%=rs2("kehu")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td></td>
                      
			<td align="center" width="13%" ><input type="submit"  value="�� ��"> </td>
			

          </tr></form>
			</TD>
              </TR>
          </table>
		         
      <table width="100%" border="0" cellspacing="0" cellpadding="4">
      
      </table>
  </TBODY>
</TABLE>
<%end sub%>
<P>
<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center border=0>
  <TBODY>
  <TR vAlign=center>
    <TD align=middle width="100%"><!--#include file="footer.htm"--></TD>
  </TR></TBODY></TABLE></P></BODY></HTML>