<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>�˼�������ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">
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
function Juge(powersearch)
{
	if (powersearch.guige.value == "")
	{
		alert("�����Ϊ�գ�");
		powersearch.guige.focus();
		return (false);
	}

	if (powersearch.loan_oan.value == "")
	{
		alert("��������Ϊ�գ�");
		powersearch.loan_oan.focus();
		return (false);
	}

	if (powersearch.cname.value == "")
	{
		alert("���Ʋ���Ϊ�գ�");
		powersearch.cname.focus();
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
<!--#include file="top.asp"--><% dim sql,rs

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
uptime=request.Form("uptime")
use_dep=trim(request.Form("use_dep")) 
cname= trim(request.Form("cname"))
guige=request.Form("guige")
loan_oan=request.Form("loan_oan")

if use_dep="ӡˢ����" then
       set rs=server.createobject("adodb.recordset") 
       sql="select * from caiyin where  cname='"&cname&"'"
       rs.open sql,conn,2,3
      if not rs.eof and not rs.bof then
         rs("qianmi")=rs("qianmi")+loan_oan
         rs("zhemi")=rs("zhemi")+loan_oan
         rs("guige")=guige
      else
    rs.addnew
       rs("uptime") =uptime
       rs("cname")=cname
       rs("pinming")="����"
       rs("guige") = guige
       rs("houdou")="0"
       rs("qianmi")=loan_oan
       rs("fenlei") ="��Ĥ"
       rs("Unit") ="ǧ��"
       rs("zhemi")=loan_oan
    end if
       rs.update
       rs.close
       set rs=nothing
       
elseif use_dep="���ϳ���" then
       set rs=server.createobject("adodb.recordset") 
       sql="select * from fuhe where  cname='"&cname&"'"
       rs.open sql,conn,2,3
    if not rs.eof and not rs.bof then
         rs("m_Loan_num")=rs("m_Loan_num")+loan_oan
         rs("guige") = guige
      else
         rs.addnew
         rs("fenlei") ="��Ĥ"
         rs("guige") = guige
         rs("houdou")= "0"
         rs("pinming") ="����"
         rs("uptime") =uptime
         rs("cname")= cname
         rs("m_Loan_num") =loan_oan
    end if
       rs.update
       rs.close
       set rs=nothing
       
elseif use_dep="�ƴ�����" then

          set rs=server.createobject("adodb.recordset") 
         sql="select * from matur where  cname='"&cname&"'"
        rs.open sql,conn,2,3
      if not rs.eof and not rs.bof then
         rs("number")=rs("number")+loan_oan
         rs("guige") = guige
      else
         rs.addnew
         rs("cname") = cname
         rs("guige") = guige
         rs("number")=loan_oan
         rs("uptime")=uptime
      end if
         rs.update     
         rs.close
         set rs=nothing       


end if

          set rs3=server.createobject("adodb.recordset") 
           sql3="select * from kucen"
           rs3.open sql3,conn,1,3
           rs3.addnew
           rs3("loan_oan")=Loan_oan
           rs3("guige") =guige
           rs3("cname") =cname
           rs3("uptime") =uptime 
           rs3("use_dep")=use_dep
           rs3.update
           		response.Write "<script language=javascript>alert('¼��ɹ���');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=kucen.asp"">"
                response.end
           rs3.close
           set rs3=nothing
end sub

 
sub SaveModify   
uptime=request.Form("uptime")
use_dep=trim(request.Form("use_dep")) 
cname= trim(request.Form("cname"))
guige=request.Form("guige")
loan_oan=request.Form("loan_oan")
loan_oan1=request.Form("loan_oan1")


num= loan_oan-loan_oan1



if use_dep="ӡˢ����" then
       set rs=server.createobject("adodb.recordset") 
       sql="select * from caiyin where  cname='"&cname&"'"
       rs.open sql,conn,2,3
      if not rs.eof and not rs.bof then
         rs("qianmi")=rs("qianmi")+num
         rs("zhemi")=rs("zhemi")+num
         rs("guige") = guige
    end if
       rs.update
       rs.close
       set rs=nothing
       
elseif use_dep="���ϳ���" then
       set rs=server.createobject("adodb.recordset") 
       sql="select * from fuhe where  cname='"&cname&"'"
       rs.open sql,conn,2,3
    if not rs.eof and not rs.bof then
         rs("m_Loan_num")=rs("m_Loan_num")+num
         rs("guige") = guige
    end if
       rs.update
       rs.close
       set rs=nothing
       
elseif use_dep="�ƴ�����" then

          set rs=server.createobject("adodb.recordset") 
         sql="select * from matur where  cname='"&cname&"'"
        rs.open sql,conn,2,3
      if not rs.eof and not rs.bof then
         rs("number")=rs("number")+num
         rs("guige") = guige
      end if
         rs.update     
         rs.close
         set rs=nothing       


end if




set rs3=server.createobject("adodb.recordset") 
sql3="select * from kucen where id="&request.Form("id")
rs3.open sql3,conn,1,3
           rs3("loan_oan")=Loan_oan
           rs3("guige") =guige
           rs3("cname") =cname
           rs3("uptime") =uptime 
           rs3("use_dep")=use_dep

rs3.update
rs3.close
set rs3=nothing
response.Write "<script language=javascript>alert('�޸ĳɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=kucen.asp"">"
response.end


end sub   
 
  sub delCate()
        conn.execute("delete from kucen where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('ɾ���ɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=kucen.asp"">"
response.end
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=600 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"><%if oskey="supper" then%> 

	<%
	    set rs=server.createobject("adodb.recordset")
	   if isedit then
		   rs.open "select * from kucen where id=" & request("id"),conn,1,1
%>

    <table width="600" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�� �� �� �� �� Ϣ</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="600" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�� �� �� �� �� Ϣ</font></b>

</td>
</tr>
</table>
<%end if %>
       <table width="600" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 

            <td width="12%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			���ʱ��</td>


			
            <td width="56%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			��Ʒ����</td>


			
            <td width="8%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			���</td>


			
            <td width="12%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			�������</td>


			
            <td width="12%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			����</td>


			
          </tr>
		  <form name="powersearch" method="post" action="kucen.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'>			</td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name='cname'  size=1 style='width:335; height:18'>
			<% if isedit then%>
				<option selected="selected" value="<%=rs("cname")%>"><%=rs("cname")%></option> 
                <input type="hidden" value="<%=rs("loan_oan")%>" name="loan_oan1">
                
			<%else%>
		       <option>��ѡ�����</option>
			
			
			<%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from dindang order by -id",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
			%>
			<option value="<%=rs1("pinming")%>"><%=rs1("pinming")%></option>
			<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%><%end if%></select>
			
			
			</td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="guige" type="text" style="ime-mode:disabled"  size="7" value='<% if isedit then
		                                                         response.write rs("guige")
																  end if %>'>
</td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="loan_oan" type="text" style="ime-mode:disabled"  size="9" value='<% if isedit then
		                                                         response.write rs("loan_oan")
																  end if %>'></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name='use_dep'  size="1" style='width:68; height:16'>
<% if isedit then%>
				<option selected="selected"><%=rs("use_dep")%></option>
			<%else%>
            <option selected value='�ƴ�����'>�ƴ�����</option>
			<option selected value='ӡˢ����'>ӡˢ����</option>
			<option selected value='���ϳ���'>���ϳ���</option>
			<%end if%>            
			</select>			</td>


          </tr>
          <tr class="but"> 

            <td width="12%" height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			ʹ��˵��</td>


			
            <td width="88%" height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="4">
			<p align="left">����Ҫ�޸ġ���Ʒ���ơ��򡰳��䡱�����Ƚ���Ⱥ���������޸�Ϊ0��������¼�룡</td>


			
          </tr>
		  <tr> 
            
			<td align="center" colspan="5" height="25"><input type="submit"  value="<% if isedit then
		                                                         response.write "�༭"
																 else
																 response.write "���"
																  end if %>"> </td>
          </tr></form>
		  
           </table>
        <%end if%></td>
</tr>
</table>
<br><br>
<TABLE width=600 align=center border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6">
<tr>
<td >


			<form action="kucen.asp" method="post" name="selform" >
  <table width="600"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="600" align="center" valign="top">
	        <table width="600" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp;�� �� �� ϸ �� ��</font></b>

</td>
</tr>
</table>
	 
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			
                        <td width="12%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
						ʱ��</td>

                        <td width="45%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
						����</td>

                        <td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
						���</td>

                        <td width="15%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
						����</td>

                        <td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
						����</td>

                        <td width="8%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�� ��</td>
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
fenlei1=request("kucen")
keyword=request("keyword")
rs.Source="select* from kucen where kucen like '%"&fenlei1&"%' order by id desc"
else
rs.Source="select* from kucen order by id desc"
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

<td height="18" align="center"><a href="kucen.asp?id=<%=rs("ID")%>&action=edit"><%=rs("cname")%></a></td>

<td height="18" align="center"><%=rs("guige")%></td>

<td height="18" align="center"><%=rs("loan_oan")%></td>

<td height="18" align="center"><%=rs("use_dep")%></td>

<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if

%>
<tr><td align="right" colspan="6" height="22">
�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
			if search<>"" then
			url="kucen.asp?kucen="&fenlei1&"&keyword="&keyword&"&search=search&"
			else
            url="kucen.asp?"
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
%>
&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="cailiao"  then%>
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
<TABLE width=600 align=center border="1" cellspacing="20" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>

	        <table width="600" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp;�� �� �� Ϣ �� ��</font></b>

</td>
</tr>
</table>



        <table width="600" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="kucen.asp">
		  <tr><input name="search" type="hidden" value="search">
            <td align="center" width="20%" height="40" valign="middle">
			<select name="guige" size="1">
			<option value="">���й��</option>
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from guige",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		  guige=rs2("guige")
		   %>
		   <option value="<%=rs2("guige")%>"><%=rs2("guige")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
            <td align="center" width="50%">���ƹؼ��֣� <input name="keyword" type="text"  size="25"></td>
                      
			<td align="center" width="30%" ><input type="submit"  value="�� ��"> </td>
          </tr></form>
		 </fieldset> 
           </table>
       
			</TD>
              </TR>
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


