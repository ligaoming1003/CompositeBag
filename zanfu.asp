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

	if (powersearch.Loan_price.value == "")
	{
		alert("���۲���Ϊ�գ�");
		powersearch.Loan_price.focus();
		return (false);
	}
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



sub SaveAdd

Loan_Amount=request.Form("Loan_Amount")
pinming=trim(request.Form("pinming"))
use_dep=trim(request.Form("use_dep"))


set rs2=server.createobject("adodb.recordset") 
sql2="select * from JS_out_store"
rs2.open sql2,conn,1,3
rs2.addnew
rs2("pinming") = pinming
rs2("uptime") = request.Form("uptime")
rs2("use_dep") = use_dep
rs2("Loan_Amount") = Loan_Amount
rs2.update
rs2.close
set rs2=nothing

set rs3=server.createobject("adodb.recordset")
sql3="select * from JS_huzong where kehu='"&use_dep&"'"
rs3.open sql3,conn,2,3
if not rs3.eof then
rs3("Loan_Amount")=rs3("Loan_Amount")+Loan_Amount
else 
rs3.addnew
rs3("Loan_Amount")=Loan_Amount
rs3("kehu") = trim(request.Form("use_dep"))            
end if
rs3.update
response.Write "<script language=javascript>alert('���Ӧ�տ�¼��ɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=zanfu.asp"">"
response.end
rs3.close
set rs3=nothing
end sub

 
sub SaveModify  
old_Amount=request.Form("old_Amount")
old_dep=request.Form("old_dep")
pinming=trim(request.Form("pinming"))
use_dep=trim(request.Form("use_dep"))
Loan_Amount=request.Form("Loan_Amount")
num= old_Amount-Loan_Amount


if old_dep<>use_dep then
response.Write "<script language=javascript>alert('�Բ��𣡿ͻ����Ʋ����޸ģ�\n\nֻ���������������븺��������������룡');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=zanfu.asp"">"
response.en
end if



set rs2=server.createobject("adodb.recordset") 
sql2="select * from JS_out_store where id="&request.Form("id")
rs2.open sql2,conn,1,3

rs2("Loan_Amount") =Loan_Amount
rs2("pinming") = pinming
rs2.update
rs2.close
set rs2=nothing

set rs3=server.createobject("adodb.recordset")
sql3="select * from JS_huzong where kehu='"&use_dep&"'"
rs3.open sql3,conn,2,3

rs3("Loan_Amount")=rs3("Loan_Amount")-num
     
rs3.update
response.Write "<script language=javascript>alert('���Ӧ�տ��޸ĳɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=zanfu.asp"">"
response.end
rs3.close
set rs3=nothing
end sub  
 
   sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('��û��ѡ����¼��');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=zanfu.asp"">"
  response.end
else
        conn.execute("delete from JS_out_store where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('ɾ���ɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=zanfu.asp"">"
response.end
end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="xiaoshou" or oskey="admin"  then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from JS_out_store where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�༭����Ӧ�տ���</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�������Ӧ�տ���</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��������</td>
            <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			�� ��</td>
			
			<td width="25%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			��&nbsp; ��</td>
			
			<td width="25%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">�ͻ�����</td>
			
          </tr>
		  <form name="powersearch" method="post" action="zanfu.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"> 
			<select name='pinming'  size="1" style='width:79; height:14'>
			<% if isedit then%>
			
			<option selected value='������'>������</option>

		    <option selected="selected"><%=rs("pinming")%></option>
				<input type="hidden" value="<%=rs("Loan_Amount")%>" name="old_Amount">
			    <input type="hidden" value="<%=rs("use_dep")%>" name="old_dep">
              
                 			<%else%>

		    <option selected value='������'>������</option>
            <option selected value='0000'>��ѡ��</option>

			<%end if%>

			
			
			</select></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="Loan_Amount" type="text" style="ime-mode:disabled"  size="19" value='<% if isedit then
		                                                         response.write rs("Loan_Amount")
																  end if %>'></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">

<select name='use_dep' size="1" style='width:130; height:18'>
			<% if isedit then%>
				<option selected="selected"><%=rs("use_dep")%></option>

			<%else%>
		        <option selected value='0000'>��ѡ��ͻ�</option>
			<%end if%>
			<%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select distinct(kehu) from dindang ",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option <% if isedit then
			if trim(rs3("kehu"))=trim(rs("use_dep")) then
			%> selected="selected"<%end if%><%end if%>><%=rs3("kehu")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%>
			</select></td>
          </tr>
			<tr>
			<td colspan="4" align="center"><input type="submit" value="<% if isedit then
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



			<form action="zanfu.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp;�� �� Ӧ �� �� �� �� �� ��</font></b>

</td>
</tr>
</table>
	  
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">��������</td>
			<td width="21%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">Ʒ ��</td>
			<td width="11%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">���</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�ͻ�����</td>
            <td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�� ��</td>
          </tr>
		 <%
PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 1000          'ÿҳ��ʾ������
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
set rs=server.CreateObject("ADODB.RecordSet") 
search=request("search")
if search<>"" then 
keyword=request("keyword")

rs.Source="select* from JS_out_store where use_dep like '%"&keyword&"%'and loan_num=0 order by -id"
else
rs.Source="select* from JS_out_store where loan_num=0 order by -id"
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
<td height="18" align="center">&nbsp;<a href="zanfu.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="18" align="center"><%=Formatnumber(rs("Loan_Amount"),2,-1,0,0)%><%if rs("Loan_Amount")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("use_dep")%><%if rs("use_dep")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="center" colspan="5" height="22">
<p align="right">�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
			if search<>"" then
			url="zanfu.asp?keyword="&keyword&"&search=search&"
			else
            url="zanfu.asp?"
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" then%>
        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> ѡ��/��ѡ
              <input onClick="{if(confirm('�˲�����ɾ������Ϣ��\n\nȷ��Ҫִ�д��������')){this.document.selform.submit();return true;}return false;}" type=submit value=ɾ�� name=action2> 
              <input type="Hidden" name="action" value='del'><%end if%></td></tr>
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
         
		  <form name="search" method="post" action="zanfu.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">���ͻ������ң� 
				<input name="keyword" type="text"  size="23"></td>
                      
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