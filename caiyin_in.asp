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
  response.Write "<script language=javascript>alert('��ҳ��ֻ���޸ģ�');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_in.asp"">"
  response.end
end sub

 
sub SaveModify
old_dep=request.Form("old_dep")
old_num=request.Form("old_num")
Loan_oan=request.Form("Loan_oan")
cname=trim(request.Form("cname"))
use_dep=trim(request.Form("use_dep"))
unit=trim(request.Form("unit"))
names=Split(cname,"|")
id=names(0)
cname=names(1)
guige=names(2)
houdou=names(3)
pihao=names(4)
num=old_num-Loan_oan


if use_dep="���ϳ���" and old_dep="���ϳ���" then
         set rs=server.createobject("adodb.recordset") 
          sql="select * from fuhe where cname='"&cname&"'"
          rs.open sql,conn,2,3
          if not rs.eof then
         rs("guige") = guige
         rs("houdou")= houdou
         rs("m_Loan_num")=rs("m_Loan_num")-num
          end if
         rs.update
         rs.close
         set rs=nothing
elseif use_dep="�컯��"  and old_dep="�컯��" then  
             set rs5=server.createobject("adodb.recordset") 
             sql5="select * from matur where cname='"&cname&"'"
             rs5.open sql5,conn,2,3
             if not rs5.eof and not rs5.bof then
             rs5("number")=rs5("number")-num
             end if
             rs5.update 
             rs5.close
             set rs5=nothing
elseif use_dep="�컯��"  and old_dep="���ϳ���" then
         set rs=server.createobject("adodb.recordset") 
          sql="select * from fuhe where cname='"&cname&"'"
          rs.open sql,conn,2,3
          if not rs.eof then
         rs("guige") = guige
         rs("m_Loan_num")=rs("m_Loan_num")-old_num
          end if
         rs.update
         rs.close
         set rs=nothing
             set rs5=server.createobject("adodb.recordset") 
             sql5="select * from matur where cname='"&cname&"'"
             rs5.open sql5,conn,2,3
             if not rs5.eof and not rs5.bof then
             rs5("number")=rs5("number")+Loan_oan
             else
              rs5.addnew
              rs5("cname") = cname
              rs5("guige") = guige
              rs5("number")=loan_oan

              rs5("uptime")=trim(request.Form("uptime"))

             end if
             rs5.update 
             rs5.close
             set rs5=nothing
elseif use_dep="���ϳ���"  and old_dep="�컯��" then
         set rs=server.createobject("adodb.recordset") 
          sql="select * from fuhe where cname='"&cname&"'"
          rs.open sql,conn,2,3
         if not rs.eof and not rs.bof then
         rs("guige") = guige
         rs("m_Loan_num")=rs("m_Loan_num")+Loan_oan
         else 
         rs.addnew
         rs("m_Loan_num")=Loan_oan
         rs("fenlei") = fenlei
         rs("guige") = guige
         rs("houdou")= houdou
         rs("pinming") = pinming
         rs("uptime") = request.Form("uptime")         
         rs("cname")= cname


          end if
         rs.update
         rs.close
         set rs=nothing
             set rs5=server.createobject("adodb.recordset") 
             sql5="select * from matur where cname='"&cname&"'"
             rs5.open sql5,conn,2,3
             if not rs5.eof and not rs5.bof then
             rs5("number")=rs5("number")-old_num
             end if
             rs5.update 
             rs5.close
             set rs5=nothing

             
else 
             response.Write "<script language=javascript>alert('���ó������');</script>"
             response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_in.asp"">"
             response.end 
            
    

end if
          set rs1=server.createobject("adodb.recordset") 
          sql1="select * from caiyin_in where id="&request.Form("id")
          rs1.open sql1,conn,1,3
          if not rs1.eof then
           rs1("unit")=unit
           rs1("loan_oan")=Loan_oan
           rs1("guige") =guige
           rs1("cname") =cname
           rs1("uptime") = request.Form("uptime") 
           rs1("houdou")=houdou

           rs1("use_dep")=use_dep
          end if  
           rs1.update
           rs1.close
           set rs1=nothing

          set rs3=server.createobject("adodb.recordset") 

           sql3="select * from caiyin where cname='"&cname&"'"

           rs3.open sql3,conn,1,3
          if not rs3.eof and not rs3.bof then
           rs3("loan_oan")=rs3("loan_oan")-num
           rs3("qianmi")=rs3("qianmi")+num
           rs3("use_dep")=use_dep
           end if
           rs3.update
             response.Write "<script language=javascript>alert('�޸ĳɹ���');</script>"
             response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_in.asp"">"
             response.end 
           rs3.close
           set rs3=nothing 


 
 
end sub   
 
sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('��û��ѡ����¼��');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_in.asp"">"
  response.end
else
        conn.execute("delete from caiyin_in where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('ɾ���ɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin_in.asp"">"
response.end
end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="admin" then%>
	
	<%
	    
	   if isedit then

	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from caiyin_in where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�༭��ӡ���������Ϣ</font></b>

</td>
</tr>
</table>
<%else%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">��Ӳ�ӡ���������Ϣ</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="19%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��������</td>
			<td width="60%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'"><p align="left">&nbsp;&nbsp;&nbsp;&nbsp; Ʒ �� | ��&nbsp; �� | ����|&nbsp; ��λ|&nbsp; ����</td>
			<td width="12%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">���ó���</td>
			
			<td width="9%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			��λ</td>
			
          </tr>
		  <form name="powersearch" method="post" action="caiyin_in.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='cname' size=1 style='width:481; height:15'>
			<% if isedit then%>
			<option selected="selected" value="0|<%=rs("cname")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("pihao")%>"><%=rs("cname")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("pihao")%></option>
			<input type="hidden" value="<%=rs("Loan_oan")%>" name="old_num">
             <input type="hidden" value="<%=rs("use_dep")%>" name="old_dep">
			<%else%>
			<option>--��ѡ��--</option>
			<%end if%>
			</select>
			</td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="use_dep"  size="1" style='width:68; height:16'>
			<% if isedit then%><option ><%=rs("use_dep")%></option><%else%>
						<option selected value='�컯��'>�컯��</option>
			<option selected value='�˿�'>�˿�</option>
			<option selected value='���ϳ���'>���ϳ���</option><%end if%>
			</select></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="unit"  size="1" style='width:50; height:15'>
			<% if isedit then%><option ><%=rs("unit")%></option><%end if%>
			<option selected value='����'>����</option>
			<option selected value='ǧ��'>ǧ��</option>

			</select></td>
          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			��&nbsp;&nbsp; ��</td>
			<td height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="3" rowspan="2">
			<p align="left">˵��:1�������������󣬵����Ʒ�������ӽ����޸ġ�<br>
&nbsp;&nbsp;&nbsp;&nbsp; 2���˿���������ֻ�����븺����������������롣</td>
          </tr>
		  <tr> 
 <td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
 <input name="Loan_oan" type="text" style="ime-mode:disabled" size="11" value='<% if isedit then
		                                                         response.write rs("Loan_oan")
																  end if %>'><font color="#FFFFFF">ǧ��(����)</font></td>



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
<TABLE width=800 align=center border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td >

  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�� ӡ �� �� �� ��</font></b>

</td>
</tr>
</table>
			<form action="caiyin_in.asp" method="post" name="selform" >
 <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">

        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">�������</td>
			<td width="52%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">Ʒ ��</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">�� ��</td>
			<td width="8%" height="18" align="center"  class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">����</td>
			<td width="4%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">�� λ</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">����</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">
			Ա��</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">����</td>
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

rs.Source="select* from caiyin_in where cname like '%"&keyword&"%' order by -id "
else
rs.Source="select* from caiyin_In order by -id"
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
<td height="18" align="center"><a href="caiyin_in.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("cname")%></a></td>
<td height="18" align="center">&nbsp;<%=rs("guige")%>*<%=rs("houdou")%></td>
<td height="18" align="center"><%if rs("loan_oan")="0" then%>&nbsp;<%else%><%=rs("loan_oan")%><%end if%></td>
<td height="18" align="center"><%=rs("Unit")%><%if rs("Unit")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("use_dep")%></td>
<td height="18" align="center"><%=rs("pihao")%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="right" colspan="8" height="22">
�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
			if search<>"" then
			url="caiyin_In.asp?keyword="&keyword&"&search=search&"
			else
            url="caiyin_In.asp?"
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
              <input onClick="{if(confirm('ȷ��Ҫɾ���ü�¼��')){this.document.selform.submit();return true;}return false;}" type=submit value=ɾ�� name=action2> 
              <input type="Hidden" name="action" value='del'><%end if%></td></tr>
        </table>

	  
</td>
    </tr>
      </table>
</form>


</td>
</tr>
</table>
</td>
</tr></table>
<br>
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>


	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�������</font></b>

</td>
</tr>
</table>

        <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="caiyin_in.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="84%">��Ʒ�����ң� 
			<input name='keyword' type="text"  size="22"></td>
                      
			<td align="center" width="17%" ><input type="submit"  value="�� ��"> </td>

          </tr></form>		   
		</td></tr>
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