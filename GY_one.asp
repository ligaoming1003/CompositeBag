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

		if (powersearch.jinge.value == "")
	{
		alert("����Ϊ�գ�");
		powersearch.jinge.focus();
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

jinge=request.Form("jinge")
pinming=trim(request.Form("pinming"))
gonghuodanwei=trim(request.Form("gonghuodanwei"))


set rs2=server.createobject("adodb.recordset") 
sql2="select * from cailiao_in_store"
rs2.open sql2,conn,1,3
rs2.addnew
rs2("zhuangtai")="ok"
rs2("pinming") = pinming
rs2("uptime") = request.Form("uptime")
rs2("gonghuodanwei") = gonghuodanwei

    if pinming="���ڿͻ�����" then
    rs2("jinge") =jinge
   else
    rs2("fukuang")=jinge
   end if
rs2.update
response.Write "<script language=javascript>alert('����ɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=GY_one.asp"">"
response.end
rs2.close
set rs2=nothing
end sub

 
sub SaveModify  
jinge=request.Form("jinge")
pinming=trim(request.Form("pinming"))
gonghuodanwei=trim(request.Form("gonghuodanwei"))

if pinming="���ڿͻ�Ƿ��" or pinming="���ڿͻ�����" then
set rs2=server.createobject("adodb.recordset") 
sql2="select * from cailiao_in_store where id="&request.Form("id")
rs2.open sql2,conn,1,3
rs2("pinming") = pinming
   if pinming="���ڿͻ�����" then
    rs2("jinge") =jinge
    rs2("fukuang")=0
   elseif pinming="���ڿͻ�Ƿ��" then
    rs2("fukuang")=jinge
    rs2("jinge") =0
   end if
rs2.update
response.Write "<script language=javascript>alert('�����޸ĳɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=GY_one.asp"">"
response.end
rs2.close
set rs2=nothing
else
response.Write "<script language=javascript>alert('�Բ��𣡷ǽ�������޸ģ�');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=GY_one.asp"">"
response.end
end if
end sub  
 
   sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('��û��ѡ����¼��');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=GY_one.asp"">"
  response.end
else
        conn.execute("delete from cailiao_in_store where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('ɾ���ɹ���');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=GY_one.asp"">"
        response.end
end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="xiaoshou" or oskey="admin" then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from cailiao_in_store where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<font color="#ffffff">
<b>&nbsp;&nbsp; �༭��Ӧ�̽���</b></font></td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp;&nbsp; ���</font></b><font color="#FFFFFF"><b>��Ӧ�̽���</b></font></td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			��������</td>
            <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			�� ��</td>
			
			<td width="25%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			����</td>
			
			<td width="25%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			�ͻ�����</td>
			
          </tr>
		  <form name="powersearch" method="post" action="GY_one.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"> 
			<select name='pinming'  size="1" style='width:99; height:14'>
			<% if isedit then%>
		    			
            <option selected value='���ڿͻ�����'>���ڿͻ�����</option>	
            <option selected value='���ڿͻ�Ƿ��'>���ڿͻ�Ƿ��</option>
            <option selected="selected"><%=rs("pinming")%></option>
            <%else%>
            <option selected value='���ڿͻ�����'>���ڿͻ�����</option>	
            <option selected value='���ڿͻ�Ƿ��'>���ڿͻ�Ƿ��</option>
			<%end if%>
			</select></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="jinge" type="text" style="ime-mode:disabled"  size="19" value='<% if isedit then
                                                                 if rs("pinming")="���ڿͻ�����" then
                                                                 response.write rs("jinge")
                                                                 else
		                                                         response.write rs("fukuang")
		                                                         end if
																  end if %>'></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">

<select name='gonghuodanwei' size="1" style='width:130; height:18'>
			<% if isedit then%>
			<option selected="selected"><%=rs("gonghuodanwei")%></option>
			<%else%>
		    <option selected value='0000'>��ѡ��Ӧ��</option>
			<%end if%>
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from gonghuodanwei",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		   gonghuodanwei=rs2("gonghuodanwei")
		   %>
		   <option value="<%=rs2("gonghuodanwei")%>"><%=rs2("gonghuodanwei")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%>			
			</select>
			</td>
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



			<form action="GY.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp;δ �� �� �� ��</font></b>

</td>
</tr>
</table>
	  
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="13%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">����</td>
			<td width="33%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�������</td>
			<td width="18%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			��������</td>
			<td width="11%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			����Ƿ��</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">��Ӧ������</td>
            <td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�� ��</td>
          </tr>
		 <%
PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 100          'ÿҳ��ʾ������
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
set rs=server.CreateObject("ADODB.RecordSet") 
search=request("search")
if search<>"" then 
keyword=request("keyword")

rs.Source="select* from cailiao_in_store where zhuangtai='ok' and gonghuodanwei like '%"&keyword&"%' order by -id"
else
rs.Source="select* from cailiao_in_store where zhuangtai='ok' order by -id"
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
<td height="18" align="center"><a href="GY_one.asp?id=<%=rs("ID")%>&action=edit"><%=rs("uptime")%></a></td>
<td height="18" align="center">&nbsp;<a href="GY_one.asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%><font color="#FF0000">��<%=rs("guige")%>*<%=rs("houdou")%>��</font></a></td>
<td height="18" align="center"><%if rs("jinge")="0" then%>&nbsp;<%else%><%=rs("jinge")%><%end if%></td>
<td height="18" align="center"><%if rs("fukuang")="0" then%>&nbsp;<%else%><%=rs("fukuang")%><%end if%></td>
<td height="18" align="center"><div title="�����ӡ"><a href="GY_in.asp?id=<%=rs("id")%>&action=edit"><%=rs("gonghuodanwei")%></a></div></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="center" colspan="6" height="22">
<p align="right">�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
			if search<>"" then
			url="GY_one.asp?keyword="&keyword&"&search=search&"
			else
            url="GY_one.asp?"
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="xiaoshou" or oskey="admin"  then%>
        <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 
ѡ��/��ѡ
              <input type=submit value=���� name=action1> 
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
<b><font color="#ffffff">&nbsp; ��&nbsp; ��&nbsp; ��&nbsp; ��</font></b>

</td>
</tr>
</table>

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="GY_one.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">���ͻ����ң� 
				<select name="keyword" size="1">
			<option value="">��ѡ��</option>
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(gonghuodanwei) from cailiao_in_store order by gonghuodanwei",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		   gonghuodanwei=rs2("gonghuodanwei")
		   %>
		   <option value="<%=rs2("gonghuodanwei")%>"><%=rs2("gonghuodanwei")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
                      
			<td align="center" width="17%" ><input type="submit"  value="�� ��"> </td>

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

