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

	if (powersearch.cname.value == "")
	{
		alert("���Ʋ���Ϊ�գ�");
		powersearch.cname.focus();
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
rs1.open "select * from fenlei",conn,1,1
if not rs1.eof then   
do while not rs1.eof
j=0	
%>
po_ca_show[<%=i%>] = '<%=rs1("fenleiname")%>';
po_ca_value[<%=i%>] = '<%=rs1("fenleiname")%>';
po_detail_show[<%=i%>] = new Array();
po_detail_value[<%=i%>] = new Array();

<%
fenleiname=rs1("fenleiname")
set rs2=server.createobject("adodb.recordset")
rs2.open "select * from mxfuhe where fenlei='"&fenleiname&"' order by id desc",conn,1,1
if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("pinming")%>|<%=rs2("guige")%>|<%=rs2("houdou")%>|<%=rs2("loan_num")%>|<%=rs2("cname")%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("pinming")%>|<%=rs2("guige")%>|<%=rs2("houdou")%>|<%=rs2("loan_num")%>|<%=rs2("cname")%>';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '���޲��ϳ���';
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
        num1= GetObjID('pinming');

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
        response.Write "<script language=javascript>alert('����ֱ������˿⣡\n\n�����б�Ʒ���������˿⡣');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=mxfuhe.asp"">"
        response.end
    
end sub

 
sub SaveModify 
fenlei= trim(request.Form("d_position1")) 
old_num=request.Form("old_num")
Loan_oan=request.Form("Loan_oan")
pinming=trim(request.Form("pinming"))
kufang=trim(request.Form("kufang"))

checkbox=trim(request.Form("checkbox"))
old_yess=trim(request.Form("old_yess"))
names=Split(pinming,"|")
id=names(0)
pinming=names(1)
guige=names(2)
houdou=names(3)
num=names(4)
cname=names(5) 
    



         set rs3=server.createobject("adodb.recordset") 
        sql3="select * from mxfuhe where id="&request.Form("id")
        rs3.open sql3,conn,1,3
        if rs3("yess")<>checkbox or rs3("pinming")<>pinming then
           response.Write "<script language=javascript>alert('��ע���Ƿ�ԳƸ��ϣ�Ҫ��ԭ���������\n\n����ֻѡ�������δѡƷ������ɲ������');</script>"
           response.write "<meta http-equiv=""refresh"" content=""0;url=mxfuhe.asp"">"
           response.end
        elseif rs3("loan_num")-rs3("loan_oan")-loan_oan<0 then
          response.Write "<script language=javascript>alert('�Բ��𣬿�治�㣬����������˿�����');</script>"
          response.write "<meta http-equiv=""refresh"" content=""0;url=mxfuhe.asp"">"
          response.end
        
        end if 
         rs3.update    
         rs3.close
         set rs3=nothing
         
         
       set rs1=server.createobject("adodb.recordset") 
       sql1="select * from cailiao_store where  pinming='"&pinming&"'and fenlei='"&fenlei&"'and guige="&guige&"and houdou="&houdou&" "
       rs1.open sql1,conn,1,3

    if not rs1.eof and not rs1.bof then

          rs1("number")=rs1("number")+Loan_oan
      else 
         rs1.addnew
         
            rs1("fenlei") = trim(request.Form("d_position1"))
            rs1("guige") = guige
            rs1("houdou")=houdou
            rs1("pinming") = pinming
            rs1("number")=Loan_oan
            rs1("unit")="����"

     end if
          rs1.update
          rs1.close
          set rs1=nothing
  
           set rs6=server.createobject("adodb.recordset") 
           sql6="select * from cailiao_in_store"
           rs6.open sql6,conn,1,3
           rs6.addnew
           rs6("fenlei") = trim(request.Form("d_position1"))
           rs6("guige") = guige
           rs6("houdou") = houdou
           rs6("pinming") = pinming
           rs6("uptime") = request.Form("uptime")
           rs6("Unit") ="����"
           rs6("gonghuodanwei") = "�����˿�"
           rs6("use_num") = request.Form("Loan_oan")
           rs6.update
           rs6.close
           set rs6=nothing

          set rs2=server.createobject("adodb.recordset")
          sql2="select * from fuhe where cname='"&cname&"'"           
 
       set rs=server.createobject("adodb.recordset")
             sql="select * from cailiao where pinming='"&pinming&"'"
             rs.open sql,conn,2,3
             bizhong=rs("bizhong")


        rs2.open sql2,conn,1,2
    if not rs2.eof and not rs2.bof then
         if fenlei="��Ĥ" then
         rs2("m_oan_num")=rs2("m_oan_num")+Formatnumber((Loan_oan/guige/houdou/bizhong),2,-1,0,0)
         elseif fenlei="��Ĥ" and checkbox=""  then 
         rs2("zh_oan_num")=Formatnumber((rs2("zh_oan_num")+Loan_oan/guige/houdou/bizhong),2,-1,0,0)
         elseif fenlei="��Ĥ" and checkbox="yess" then 
         rs2("zh1_oan_num")=Formatnumber((rs2("zh1_oan_num")+Loan_oan/guige/houdou/bizhong),2,-1,0,0) 
         elseif fenlei="��Ĥ" then 
         rs2("l_oan_num")=rs2("l_oan_num")+Formatnumber((Loan_oan/guige/houdou/bizhong),2,-1,0,0)
         end if
    end if
        rs2.update
        rs2.close
        set rs2=nothing
            rs.update
            rs.close
            set rs=nothing
            
 set rs=server.createobject("adodb.recordset") 
        sql="select * from mxfuhe where id="&request.Form("id")
        rs.open sql,conn,1,3
        rs("Loan_oan") =rs("Loan_oan")+Loan_oan
        rs.update  
        response.Write "<script language=javascript>alert('�˿�ɹ���');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=mxfuhe.asp"">"
        response.end
        rs.close
        set rs=nothing



end sub  
 sub delCate()
 selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
          response.Write "<script language=javascript>alert('��û��ѡ����¼��');</script>"
          response.write "<meta http-equiv=""refresh"" content=""0;url=mxfuhe.asp"">"
          response.end
  else  
        conn.execute("delete from mxfuhe where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('ɾ���ɹ���');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=mxfuhe.asp"">"
        response.end
  end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="cailiao" or oskey="admin" then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from mxfuhe where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�༭�����˿���Ϣ</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">��Ӹ����˿���Ϣ</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">��������</td>
            <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">�� ��</td>
			<td width="45%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="2">
			<p align="left">������ | ��&nbsp; �� | ������| ��Ʒ��| ����</td>
			
			
          </tr>
		  <form name="powersearch" method="post" action="mxfuhe.asp" onSubmit="return Juge(this)" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"> 
			<select name='d_position1' onchange='Do_po_Change(this);' valign=top style='width:90'>
			<% if isedit then%>
				<option selected="selected"><%=rs("fenlei")%></option>
			<%else%>
		        <option selected value='0000'>��ѡ�����</option>
			<%end if%>
			
			<%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from fenlei",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
			%>
			<option value="<%=rs1("fenleiname")%>"><%=rs1("fenleiname")%></option>
			<%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
			%></select></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="2">
			<p align="left">
			<select name='pinming' size=1 style='width:435; height:13'>
			<% if isedit then%>
			<option selected="selected" value="<%=rs("id")%>|<%=rs("pinming")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("Loan_num")%>|<%=rs("cname")%>"><%=rs("pinming")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("Loan_num")%>|<%=rs("cname")%></option>
			<input type="hidden" value="<%=rs("Loan_oan")%>" name="old_num">

			<input type="hidden" value="<%=rs("yess")%>" name="old_yess">
			<%else%>
			<option>--��ѡ��--</option>
			<%end if%>
			</select>&nbsp;
			</td>

          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="2">
			�˻زֿ�</td>
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			����</td>
			<td height="36" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" rowspan="2">
			<font color="#FFFFFF">����˵��: ����б���Ʒ���������ӣ�<br>
			ѡ��&quot;���ϲֿ�&quot;������������ȷ��</font></td>
          </tr>
		  <tr> 
 <td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="2">
<select name="kufang"  size="1" style='width:89; height:16'>
			<% if isedit then%><option ><%=rs("kufang")%></option><%end if%>
			<%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from cailiaocangku",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option <% if isedit then
			if trim(rs3("cangkuname"))=trim(rs("kufang")) then
			%> selected="selected"<%end if%><%end if%>><%=rs3("cangkuname")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%>
			</select></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="Loan_oan" type="text"  size="19" value=""><font color="#FFFFFF">����</font></td>



			</tr>
			<tr>
			<td colspan="4" align="center"><input type="submit" value="<% if isedit then
		                                                         response.write "ȷ��"
																 else
																 response.write "ȷ��"
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



			<form action="mxfuhe.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">&nbsp;�� �� �� �� �� �� �� ��</font></b>

</td>
</tr>
</table>
	  
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">��������</td>
            <td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�� ��</td>
			<td width="46%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">��Ʒ��</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">Ʒ ��</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�� ��</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�������</td>
			<td width="7%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�˿�����</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">�� λ</td>
            <td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">����</td>
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
rs.Source="select* from mxfuhe where pinming like '"&keyword&"' order by -id"
else
rs.Source="select* from mxfuhe order by -id"
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
<td height="18" align="center"><a href="mxfuhe.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("uptime")%></a></td>
<td height="18" align="center"><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><a href="mxfuhe.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("cname")%></a></td>
<td height="18" align="center"><a href="mxfuhe.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="18"  align="center"><%=rs("guige")%>*<%=rs("houdou")%> </td>
<td height="18" align="center"><%=Formatnumber(rs("Loan_num"),2,-1,0,0)%><%if rs("Loan_num")="" then%><%end if%></td>
<td height="18" align="center"><%if rs("loan_oan")="0" then%>&nbsp;<%else%><%=Formatnumber(rs("Loan_oan"),2,-1,0,0)%><%end if%></td>
<td height="18" align="center">����</td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="center" colspan="9" height="22">
�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
			if search<>"" then
			url="mxfuhe.asp?keyword="&keyword&"&search=search&"
			else
            url="mxfuhe.asp?"
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
              <input onClick="{if(confirm('�˲�����ɾ������Ϣ��\n\nΪ�����ݵ�һ�£����Ȱѳ��������޸�Ϊ0��ɾ��\n\nȷ��Ҫִ�д��������')){this.document.selform.submit();return true;}return false;}" type=submit value=ɾ�� name=action2> 
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
<TABLE width=820 align=center border="1" cellspacing="20" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">�˿�����</font></b>

</td>
</tr>
</table>

        <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		           
<form name="search" method="post" action="mxfuhe.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">��Ʒ�����ң� 
				
				<select name='keyword'  style='width:130'>
			<option selected value="">��ѡ��Ʒ��</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(pinming) from mxfuhe",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
			%>
			<option value="<%=rs2("pinming")%>"><%=rs2("pinming")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
                      
			<td align="center" width="13%" ><input type="submit"  value="�� ��"> </td>
			

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
