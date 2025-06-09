<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>∷嘉美管理系统∷</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="css/main.css" rel="stylesheet" type="text/css">

<SCRIPT language=JavaScript src="css/User_Info_Modify.js"></SCRIPT>
<style type="text/css">
<!--
td {  font-family: "宋体"; font-size: 9pt}
body {  font-family: "宋体"; font-size: 9pt}
select {  font-family: "宋体"; font-size: 9pt}
A {text-decoration: none; color: #336699; font-family: "宋体"; font-size: 9pt}
A:hover {text-decoration: underline; color: #FF0000; font-family: "宋体"; font-size: 9pt} 
-->
</style>
<SCRIPT LANGUAGE=javascript>
<!--
function Juge(powersearch)
{

	if (powersearch.pinming.value == "")
	{
		alert("名称不能为空！");
		powersearch.pinming.focus();
		return (false);
	}
		if (powersearch.number.value == "")
	{
		alert("数量不能为空！");
		powersearch.number.focus();
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
rs2.open "select * from cailiao where fenlei='"&fenleiname&"' order by pinming",conn,1,1
if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("pinming")%>|<%=rs2("number")%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("pinming")%>|<%=rs2("number")%>';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '此类暂无材料';
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

     response.Write "<script language=javascript>alert('不可添加！调整请点击品名链接！.');</script>"
     response.write "<meta http-equiv=""refresh"" content=""0;url=tiaozhen.asp"">"
     response.end

end sub

 
sub SaveModify 
uptime=date()
fenlei= trim(request.Form("d_position1")) 
guige=request.Form("guige")
houdou=request.Form("houdou")
old_num=request.Form("old_num")
old_id=request.Form("old_id")
old_fenlei=request.Form("old_fenlei")
old_guige=request.Form("old_guige")
unit=request.Form("unit")
danjia=request.Form("danjia")
number=request.Form("number")
pinming=trim(request.Form("pinming"))
names=Split(pinming,"|")
id=names(0)
pinming=names(1)

if old_num-number<0 then
          response.Write "<script language=javascript>alert('对不起，库存不足，请调整您的出库量！');</script>"
          response.write "<meta http-equiv=""refresh"" content=""0;url=tiaozhen.asp"">"
          response.end
end if          
       set rs=server.createobject("adodb.recordset") 
       sql="select * from cailiao_store where  id="&old_id&""
       rs.open sql,conn,1,3

    if not rs.eof and not rs.bof then

          rs("number")=rs("number")-number

     end if
          rs.update
          rs.close
          set rs=nothing
    
    set rs2=server.createobject("adodb.recordset") 
    sql2="select * from cailiao_out_store"
    rs2.open sql2,conn,2,3
    rs2.addnew
    rs2("fenlei") = old_fenlei
    rs2("guige") = old_guige
    rs2("houdou") = houdou
    rs2("pinming") = pinming
    rs2("Unit") = unit
    rs2("loan_num") = number
    rs2("cname")="调剂"
    rs2("kufang")=guige
    rs2("uptime") = uptime
    rs2("old_fenlei") =fenlei
    rs2.update
    rs2.close
    set rs2=nothing
    
    set rs3=server.createobject("adodb.recordset") 
    sql3="select * from cailiao_in_store"
    rs3.open sql3,conn,2,3
    rs3.addnew
    rs3("fenlei") =fenlei
    rs3("old_fenlei") = old_fenlei
    rs3("guige") =guige
    rs3("houdou") = houdou
    rs3("pinming") = pinming
    rs3("Unit") = unit
    rs3("use_num") = number
    rs3("gonghuodanwei")="调来"
    rs3("end_time")=old_guige
    rs3("uptime") = uptime
    rs3.update
    rs3.close
    set rs3=nothing



set rs1=server.createobject("adodb.recordset") 
sql1="select * from cailiao_store where fenlei='"&fenlei&"'and pinming='"&pinming&"'and guige="&guige&" and houdou="&houdou&""
rs1.open sql1,conn,2,3
if not rs1.eof and not rs1.bof then
rs1("number")=rs1("number")+number
rs1("jinge")=rs1("jinge")+number*danjia
else
rs1.addnew
rs1("fenlei") = fenlei
rs1("guige") = guige
rs1("houdou") = houdou
rs1("pinming") = pinming
rs1("Unit") = unit
rs1("number") = number

end if
rs1.update  
response.Write "<script language=javascript>alert('调剂成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=tiaozhen.asp"">"
response.end
rs1.close
set rs1=nothing



end sub   
 
  sub delCate()
    selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
          response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
          response.write "<meta http-equiv=""refresh"" content=""0;url=tiaozhen.asp"">"
          response.end
  else
        conn.execute("delete from cailiao_store where id in ("&Request.Form("selBigClass")&")and number=0")
		response.Write "<script language=javascript>alert('数量为0的已删除成功！\n\n数量不为0的没有被删除！');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=tiaozhen.asp"">"
        response.end
  end if
  end sub
  %> <% sub myform(isEdit) %> 	  
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
  <TBODY> 
  <TR> 
    <td align="center" valign="top"> <%if oskey="supper" or oskey="cailiao" or oskey="admin"  then%>
	
	<%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from cailiao_store where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编辑调库信息</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添加调库信息</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			类 别</td>			
			<td width="25%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			品 名 |&nbsp; 余数</td>			
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">宽度</td>		
			
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">厚度</td>
			
			
			<td width="15%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">数量</td>
			
			
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">单位</td>
			
          </tr>
		  <form name="powersearch" method="post" action="tiaozhen.asp" onSubmit="return Juge(this)" >
		  
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			
			<select name='d_position1' onchange='Do_po_Change(this);' valign=top style='width:90'>
			<% if isedit then%>
				<option selected="selected"><%=rs("fenlei")%></option>
			<%else%>
		        <option selected value='0000'>请选择类别</option>
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

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<select name='pinming' size=1 style='width:169; height:17'>
			<% if isedit then%>
			<option selected="selected" value="0|<%=rs("pinming")%>|<%=rs("number")%>"><%=rs("pinming")%>|<%=rs("number")%></option>
			<input type="hidden" value="<%=rs("number")%>" name="old_num">
			<input type="hidden" value="<%=rs("id")%>" name="old_id">
			<input type="hidden" value="<%=rs("fenlei")%>" name="old_fenlei">
			<input type="hidden" value="<%=rs("guige")%>" name="old_guige">
			<%else%>
			<option>--请选择--</option>
			<%end if%>
			</select></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="guige" type="text" style="ime-mode:disabled"  size="8" value='<% if isedit then
		                                                         response.write rs("guige")
																  end if %>'></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="houdou" type="text" style="ime-mode:disabled"  size="7" value='<% if isedit then
		                                                         response.write rs("houdou")
																  end if %>'></td>


<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="number" type="text" style="ime-mode:disabled"  size="13" value=''></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<input name="unit" type="text"  size="8" value='<% if isedit then
		                                                         response.write rs("unit")
																  end if %>'></td>

          </tr><tr> 
<td height="54" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="6"> 
			<p align="left"><font color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 操作说明: 
			点击列表中品名栏的链接，选择类别—选择品名—输入数量—确定<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			1、如：PET登记在面膜中，复合时PET作为中膜，请选择“类别”，将面膜PET调整为中膜pet，再出库。<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
			2、宽度大的料子需要分切成两个小规格的，请在修改宽度，分配好重量，厚度不变，</font></td>

          </tr>
			<tr>
			<td colspan="6" align="center"><input type="submit" value="<% if isedit then
		                                                         response.write "确定"
																 else
																 response.write "确定"
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
<b><font color="#ffffff">库 存 材 料 列 表</font></b>

</td>
</tr>
</table>
			<form action="tiaozhen.asp" method="post" name="selform" >
  <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
    <tr> 
      
      <td width="100%" align="center" valign="top">
	
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="20%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">类 别</td>
			<td width="25%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">品 名</td>
			<td width="15%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">规 格</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">单 位</td>
			<td width="15%" height="18" align="center"  class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">库存数量</td>
            <td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">操 作</td>
          </tr>
		 <%
PageShowSize = 10            '每页显示多少个页
MyPageSize   = 20          '每页显示多少条
If Not IsNumeric(Request("page")) Or IsEmpty(Request("page")) Or Request("page") <=0 Then
MyPage=1
Else
MyPage=Int(Abs(Request("page")))
End if
set rs=server.CreateObject("ADODB.RecordSet")
search=request("search")
if search<>"" then 
fenlei1=request("fenlei")
keyword=request("keyword")
rs.Source="select* from cailiao_Store where number<>0 and fenlei like '%"&fenlei1&"%' and pinming like '"&keyword&"' order by -id"
else
rs.Source="select* from cailiao_Store where number<>0 order by -id"
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
<td height="18" align="center"><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td height="18"  align="center"><a href="tiaozhen.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="18" align="center">&nbsp;<%=rs("guige")%>*<%=rs("houdou")%></td>
<td height="18" align="center"><%=rs("Unit")%><%if rs("Unit")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%if rs("number")="0" then%><font color="#FF0000">没有库存</font><%else%><%=Formatnumber(rs("number"),2,-1,0,0)%><%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="right" colspan="10" height="22">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="tiaozhen.asp?fenlei="&fenlei1&"&keyword="&keyword&"&search=search&"
			else
            url="tiaozhen.asp?"
			end if				
PageNextSize=int((MyPage-1)/PageShowSize)+1
Pagetpage=int((total-1)/rs.PageSize)+1

if PageNextSize >1 then
PagePrev=PageShowSize*(PageNextSize-1)
Response.write "<a class=black href='" & Url & "page=" & PagePrev & "' title='上" & PageShowSize & "页'>上一翻页</a> "
Response.write "<a class=black href='" & Url & "page=1' title='第1页'>页首</a> "
end if
if MyPage-1 > 0 then
Prev_Page = MyPage - 1
Response.write "<a class=black href='" & Url & "page=" & Prev_Page & "' title='第" & Prev_Page & "页'>上一页</a> "
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
Response.write "<a class=black href='" & Url & "page=" & Next_Page & "' title='第" & Next_Page & "页'>下一页</A>"
end if

if MaxPages > PageShowSize*PageNextSize then
PageNext = PageShowSize * PageNextSize + 1
Response.write " <A class=black href='" & Url & "page=" & Pagetpage & "' title='第"& Pagetpage &"页'>页尾</A>"
Response.write " <a class=black href='" & Url & "page=" & PageNext & "' title='下" & PageShowSize & "页'>下一翻页</a>"
End if
rs.Close
set rs=nothing
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" then%>
       <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input onClick="{if(confirm('此操作只能删除数量为0的信息！\n\n确定要执行此项操作吗？')){this.document.selform.submit();return true;}return false;}" type=submit value=删除 name=action2> 
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
<TABLE width=800 align=center border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">库存搜索</font></b>

</td>
</tr>
</table>

        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="tiaozhen.asp">
		  <tr><input name="search" type="hidden" value="search">
            <td  height="40" valign="middle" align="center" width="20%">
			<select name="fenlei" size="1">
			<option value="">所有分类</option>
			<%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from fenlei",conn,1,1
		   if not rs2.eof then
		   do while not rs2.eof
		   fenleiname=rs2("fenleiname")
		   %>
		   <option value="<%=rs2("fenleiname")%>"><%=rs2("fenleiname")%></option>
			<%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
			%></select></td>
            <td align="center" width="65%">按品名查询： 
            <select name="keyword"  style='width:192; height:14'>
			<option selected value="">请选择品名</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(pinming) from cailiao_store order by pinming",conn,2,2
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
			%></select>
            
            
            </td>
                      
			<td align="center" width="15%" ><input type="submit"  value="查 找"> </td>
          </tr></form>
		  
           </table>
        
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

