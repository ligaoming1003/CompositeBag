<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>�˹���ϵͳ��</title>
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
		if (powersearch.Loan_oan.value == "")
	{
		alert("��Ʒ��������Ϊ�գ�");
		powersearch.Loan_oan.focus();
		return (false);
	}
       if (powersearch.one_fu.value == "")
	{
		alert("���Ʒ����Ϊ�գ�");
		powersearch.Loan_oan.focus();
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


<script>
  var po_ca_show = new Array();
  var po_ca_value = new Array();
  var po_detail_show = new Array();
  var po_detail_value = new Array();
  var funtype1;
//Modified by Ryan Gao 2006-03-30 16:01:09
<%
i=0
		
set rs1=server.createobject("adodb.recordset")
rs1.open "select * from chxp",conn,1,1
if not rs1.eof then   
do while not rs1.eof
j=0	
%>
po_ca_show[<%=i%>] = '<%=rs1("chxp")%>';
po_ca_value[<%=i%>] = '<%=rs1("chxp")%>';
po_detail_show[<%=i%>] = new Array();
po_detail_value[<%=i%>] = new Array();

<%
chxp=rs1("chxp")
set rs2=server.createobject("adodb.recordset")
rs2.open "select * from fuhe where chxp='"&chxp&"' order by m_loan_num-Loan_oan desc",conn,1,1
if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("cname")%>|<%=rs2("guige")%>|<%=Formatnumber((rs2("m_loan_num")-rs2("Loan_oan")-rs2("m_oan_num")),2,-1,0,0)%>|km';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("cname")%>|<%=rs2("guige")%>|<%=Formatnumber((rs2("m_loan_num")-rs2("Loan_oan")-rs2("m_oan_num")),2,-1,0,0)%>|km';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '���ް��Ʒ����';
po_detail_value[<%=i%>][<%=j%>] = '';
<%
end if
rs2.close
set rs2=nothing
%>

<%
rs1.movenext
i=i+1
loop
end if
rs1.close
set rs1=nothing
%>

</script>
<script Language="JavaScript">



var psid="";

function DoLoad(form,funtypev){
        var n;
        var i,j,k;
        var num;
        num= GetObjID('funtype[]');
        num1= GetObjID('cname');

        if (!funtypev)
           return;
        k=0;
        for (i=0;i<po_ca_show.length;i++) {
         for(j = 0; j < po_detail_value[i].length; j++){
                if(funtypev.indexOf(po_detail_value[i][j])!=-1) {
                    NewOptionName = new Option(po_detail_show[i][j], po_detail_value[i][j]);
                    form.elements[num].options[k] = NewOptionName;
                    k++;
                    }
         }
        }
}

function Do_po_Change(form){
        var num,n, i, m;
        num= GetObjID('d_position1');
        m = document.powersearch.elements[num].selectedIndex-1;
        n = document.powersearch.elements[num + 1].length;
        for(i = n - 1; i >= 0; i--)
                document.powersearch.elements[num + 1].options[i] = null;

        if (m>=0) {
        for(i = 0; i < po_detail_show[m].length; i++){
                NewOptionName = new Option(po_detail_show[m][i], po_detail_value[m][i]);
                document.powersearch.elements[num + 1].options[i] = NewOptionName;
        }
        document.powersearch.elements[num + 1].options[0].selected = true;
        }
}

function InsertItem(ObjID, Location)
{
	len=document.powersearch.elements[ObjID].length;
	for (counter=len; counter>Location; counter--)
	{
		Value = document.powersearch.elements[ObjID].options[counter-1].value;
		Text2Insert  = document.powersearch.elements[ObjID].options[counter-1].text;
		document.powersearch.elements[ObjID].options[counter] = new Option(trimPrefixIndent(Text2Insert), Value);
	}
}
function GetLocation(ObjID, Value)
{
  total=document.powersearch.elements[ObjID].length;
  for (pp=0; pp<total; pp++)
     if (document.powersearch.elements[ObjID].options[pp].text == "---"+Value+"---")
     { return (pp);
       break;
     }
  return (-1);
}

function GetObjID(ObjName)
{
  for (var ObjID=0; ObjID < window.powersearch.elements.length; ObjID++)
    if ( window.powersearch.elements[ObjID].name == ObjName )
    {  return(ObjID);
       break;
    }
  return(-1);
}




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
pihao=trim(request.Form("pihao"))
use_dep=trim(request.Form("use_dep")) 
Loan_oan=request.Form("Loan_oan")
cname=trim(request.Form("cname"))
one_fu=trim(request.Form("one_fu"))
names=Split(cname,"|")
id=names(0)
cname=names(1)
guige=names(2)


if guige="0" then
response.Write "<script language=javascript>alert('�Բ���,����δ����,�ݲ��ܳ��⣡');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=fh.asp"">"
                response.end
end if
if use_dep<>"�컯��" then
                response.Write "<script language=javascript>alert('�Բ���,����Ʒֻ�ܽ����컯�ң�');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=fh.asp"">"
                response.end
else
         set rs=server.createobject("adodb.recordset") 

         sql="select * from matur where  cname='"&cname&"'"

        rs.open sql,conn,2,3
      if not rs.eof and not rs.bof then
         rs("number")=rs("number")+Loan_oan
      else
         rs.addnew
         rs("cname") = cname
         rs("guige") = guige
         rs("number")=loan_oan
         rs("uptime")=trim(request.Form("uptime"))
      end if
         rs.update     
         rs.close
         set rs=nothing 
          
           set rs3=server.createobject("adodb.recordset") 
           sql3="select * from fuhe_in"
           rs3.open sql3,conn,1,3
           rs3.addnew
           rs3("loan_oan")=Loan_oan
           rs3("guige") =guige
           rs3("cname") =cname
           rs3("uptime") = request.Form("uptime") 
           rs3("one_fu")=one_fu
           rs3("use_dep")=use_dep
           rs3("pihao")=pihao
           rs3.update
           rs3.close
           set rs3=nothing 
           
          set rs4=server.createobject("adodb.recordset") 
           sql4="select * from dindang where pinming='"&cname&"'"
            rs4.open sql4,conn,1,2
        if not rs4.eof then
           rs4("fuhe")="ok"
         end if
          rs4.update
          rs4.close
          set rs4=nothing  
   
           set rs1=server.createobject("adodb.recordset") 
           sql1="select * from fuhe where id="&id&""
            rs1.open sql1,conn,1,2
        if not rs1.eof then
           rs1("Loan_oan")=rs1("Loan_oan")+Loan_oan
           rs1("use_dep")=use_dep
        end if
        rs1.update
                response.Write "<script language=javascript>alert('����ɹ���');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=fh.asp"">"
                response.end
         rs1.close
         set rs1=nothing 
end if

end sub

 
sub SaveModify
end sub   
 
 sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('��û��ѡ����¼��');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=fh.asp"">"
  response.end
else
        conn.execute("delete from fuhe_in where id in ("&Request.Form("selBigClass")&") and loan_oan=0 and one_fu=0")
		response.Write "<script language=javascript>alert('����Ϊ0��ɾ���ɹ���\n\n������Ϊ0��û�б�ɾ����');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=fh.asp"">"
        response.end
end if

  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="admin" or oskey="chejian" then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from fuhe_in where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�༭���ϳ��������Ϣ</font></b>

</td>
</tr>
</table>
<%else%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">��Ӹ��ϳ��������Ϣ</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="40" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			����</td>
            <td width="50" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">�� ��</td>
			<td width="550" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'"><p align="left">&nbsp;&nbsp;&nbsp;&nbsp; 
			Ʒ �� | ��&nbsp; �� | ����|&nbsp; ��λ</td>
			<td width="80" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">���ó���</td>
			
			<td width="80" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			Ա��</td>
			
          </tr>
		  <form name="powersearch" method="post" action="fh.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"> 
			<select name='d_position1' onchange='Do_po_Change(this);' valign=top style='width:58; height:11'>
			<% if isedit then%>
				<option selected="selected"><%=rs("chxp")%></option> 
			<%else%>
		       <option selected value='0000'>��ѡ�����</option>
			<%end if%>
			
			<%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from chxp",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
			%>
			<option value="<%=rs1("chxp")%>"><%=rs1("chxp")%></option>
			<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%></select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='cname' size=1 style='width:515; height:17'>
			<% if isedit then%>
			<option selected="selected" value="0|<%=rs("cname")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("Loan_oan")%>|km"><%=rs("cname")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("Loan_oan")%>|km</option>

			<%end if%>
			</select>
			</td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="use_dep"  size="1"  readonly="readonly "  style='width:59; height:16' value="�컯��">
			</td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="pihao" type="text"  readonly="readonly "  size="6" value="<%=name%>"></td>
          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			���Ʒ</td>
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			��Ʒ</td>
			<td height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" rowspan="2" width="710" colspan="3">
			<p align="justify">˵��:1����λΪǧ�ף��磺20000��Ӧ�20 ��������ƣ�<br>
&nbsp;&nbsp;&nbsp;&nbsp; 2����Ʒ����Ϊ���һ�㸴�ϵ����������컯�ҵ���������&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br>
&nbsp;&nbsp;&nbsp;&nbsp; 3�����ƷΪ�����м�㸴��������<br>
&nbsp;&nbsp;&nbsp;&nbsp; 4�������������������б��е��Ʒ���������޸ġ�</td>
          </tr>
		  <tr> 
 <td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
 			<input name="one_fu" style="ime-mode:disabled"  type="text" size="6" value='0'onfocus="if(this.value=='0')this.value=''" onblur="if(this.value=='')this.value='0'"></td>



 <td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="Loan_oan" type="text" style="ime-mode:disabled"  size="7" value='0'onfocus="if(this.value=='0')this.value=''" onblur="if(this.value=='')this.value='0'">
			</td>



			</tr>
			<tr>
			<td colspan="5" align="center"><input type="submit" value="<% if isedit then
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
<font color="#FFFFFF"><b>&nbsp;&nbsp; �� �� �� �� �� ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><input type="button" name="button" id="button" onclick="window.open('dindang1.asp')" value="�鿴����" />&nbsp;&nbsp;<input type="button" name="button" id="button" onclick="window.open('dindang.asp')" value="�鿴����" /></font></b>

</td>
</tr>
</table>
			<form action="fh.asp" method="post" name="selform" >
 <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">

        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">�������</td>
			<td width="48%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">Ʒ ��</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">�� ��</td>
			<td width="8%" height="18" align="center"  class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">
			���Ʒ</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">��Ʒ</td>
			<td width="14%" height="18" align="center" class="but01" onMouseDown="this.classname='tddown'" onMouseUp="this.classname='but'" onMouseOut="this.classname='but01'">Ա��</td>
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
cname1=request("cname")
rs.Source="select* from fuhe_in where  cname like '%"&cname1&"%' and pihao like '%"&keyword&"%'  and use_dep<>'null'  order by -id "
else
rs.Source="select* from fuhe_In where use_dep<>'null' order by -id"
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
<td height="18" align="center"><a href="fh_shc.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("cname")%></a></td>
<td height="18" align="center"><%=rs("guige")%></td>
<td height="18" align="center"><%if rs("one_fu")="0" then%>&nbsp;<%else%><%=rs("one_fu")%><%end if%></td>
<td height="18" align="center"><%if rs("loan_oan")="0" then%>&nbsp;<%else%><%=rs("loan_oan")%><%end if%></td>
<td height="18" align="center"><%if rs("pihao")="" then%>&nbsp;<%else%><%=rs("pihao")%><%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="right" colspan="7" height="22">
�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
			if search<>"" then
			url="fh.asp?cname="&cname1&"&keyword="&keyword&"&search=search&"
			else
            url="fh.asp?"
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="chejian"  then%>
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

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="fh.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="25%">��Ա���� 
			<select name='keyword'  style='width:73; height:16'>
			<option selected value="">��ѡ������</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(pihao) from fuhe_in order by pihao",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("pihao")%>"><%=rs2("pihao")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
                      
			<td align="center" width="10%" ><input type="submit"  value="�� ��"> </td>
			
			<td align="center" width="75%">��Ʒ���� 
			<select name="cname"  style='width:444; height:13'>
			<option selected value="">��ѡ��Ʒ��</option>
			<%			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select distinct(cname) from fuhe_in order by cname",conn,2,2
		   if not rs4.eof then
		   do while not rs4.eof
			%>
			<option value="<%=rs4("cname")%>"><%=rs4("cname")%></option>
			<%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
			%></select></td>
                      
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