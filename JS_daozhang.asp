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




end sub

 
sub SaveModify  



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

<TABLE width=820 align=center border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6" height="143">
<tr>
<td >



			<form action="JS_out.asp" method="post" name="selform" >
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
<td height="18" align="center"><%=rs("kehu")%></td>
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" then%>
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
         
		  <form name="search" method="post" action="JS_daozhang.asp">
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