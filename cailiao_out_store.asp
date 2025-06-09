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
		if (powersearch.cname.value == "0000")
	{
		alert("请选择名称！");
		powersearch.cname.focus();
		return (false);
	}

		if (powersearch.Loan_num.value == "")
	{
		alert("数量不能为空！");
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
rs2.open "select * from cailiao_store where fenlei='"&fenleiname&"'and number<>0 order by pinming",conn,1,1
if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("pinming")%>|<%=rs2("guige")%>|<%=rs2("houdou")%>|<%=Formatnumber(rs2("number"),2,-1,0,0)%>|<%=rs2("unit")%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("pinming")%>|<%=rs2("guige")%>|<%=rs2("houdou")%>|<%=Formatnumber(rs2("number"),2,-1,0,0)%>|<%=rs2("unit")%>';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '暂无材料出库';
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
fenlei= trim(request.Form("d_position1"))
Loan_num=request.Form("Loan_num")
pinming=trim(request.Form("pinming"))
cname=trim(request.Form("cname"))
pihao=trim(request.Form("pihao"))
checkbox=trim(request.Form("checkbox"))
use_dep=trim(request.Form("use_dep"))
names=Split(pinming,"|")
id=names(0)
pinming=names(1)
guige=names(2)
houdou=names(3)
num=names(4)
unit=names(5)

        if fenlei<>"面膜" and use_dep="彩印车间" then 
          response.Write "<script language=javascript>alert('对不起，彩印车间只能接受面膜！');</script>"
          response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_out_store.asp"">"
          response.end
        end if
       

       set rs1=server.createobject("adodb.recordset") 
       sql1="select * from cailiao_store where id="&id&""
       rs1.open sql1,conn,1,3
          if rs1("number")-Loan_num<-0.001 then
          response.Write "<script language=javascript>alert('对不起，库存不足，请调整您的出库量！');</script>"
          response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_out_store.asp"">"
          response.end
          end if
          rs1.update
          rs1.close
          set rs1=nothing

if use_dep="彩印车间" then
             set rs=server.createobject("adodb.recordset")
             sql="select * from cailiao where pinming='"&pinming&"'"
             rs.open sql,conn,2,3
             bizhong=rs("bizhong")
          set rs2=server.createobject("adodb.recordset")
          sql2="select * from caiyin where cname='"&cname&"'"
       rs2.open sql2,conn,2,3
    if not rs2.eof and not rs2.bof then
       rs2("guige") = guige
       rs2("houdou")=houdou
       rs2("pinming") = pinming
       rs2("Loan_num")=rs2("Loan_num")+Loan_num
       rs2("qianmi")=Formatnumber((rs2("qianmi")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
       rs2("zhemi")=Formatnumber((rs2("zhemi")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
    else
       rs2.addnew
       rs2("fenlei") = fenlei
       rs2("guige") = guige
       rs2("houdou")=houdou
       rs2("pinming") = pinming
       rs2("uptime") = request.Form("uptime")
       rs2("Unit") = unit
       rs2("Loan_num") =Loan_num
       rs2("qianmi")=Formatnumber((rs2("loan_num")/guige/houdou/bizhong),2,-1,0,0)
       rs2("zhemi")=Formatnumber((rs2("loan_num")/guige/houdou/bizhong),2,-1,0,0)
       rs2("cname") = cname
       rs2("pihao") = pihao
    end if
       rs2.update
       rs2.close
       set rs2=nothing
            rs.update
            rs.close
            set rs=nothing

elseif use_dep="复合车间" then
       set rs=server.createobject("adodb.recordset")
             sql="select * from cailiao where pinming='"&pinming&"'"
             rs.open sql,conn,2,3
             bizhong=rs("bizhong")

       
       set rs3=server.createobject("adodb.recordset")
       sql3="select * from fuhe where cname='"&cname&"'"
        rs3.open sql3,conn,2,3
    if not rs3.eof and not rs3.bof then
         if fenlei="面膜" then
         rs3("m_Loan_num")=Formatnumber((rs3("m_Loan_num")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
         rs3("guige") = guige
         rs3("houdou")= houdou 
         rs3("pinming") = pinming        
         elseif fenlei="中膜" and checkbox=""  then 
         rs3("zh_Loan_num")=Formatnumber((rs3("zh_Loan_num")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
         elseif fenlei="中膜" and checkbox="yess" then 
         rs3("zh1_Loan_num")=Formatnumber((rs3("zh1_Loan_num")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
         elseif fenlei="内膜" then 
         rs3("l_Loan_num")=Formatnumber((rs3("l_Loan_num")+Loan_num/guige/houdou/bizhong),2,-1,0,0)
         end if
         
      else
         rs3.addnew
         
         if fenlei="面膜" then
         rs3("m_Loan_num")=Formatnumber((Loan_num/guige/houdou/bizhong),2,-1,0,0)
         rs3("fenlei") = fenlei
         rs3("guige") = guige
         rs3("houdou")= houdou
         rs3("pinming") = pinming
         rs3("uptime") = request.Form("uptime")
         rs3("cname")= cname
         rs3("pihao") = pihao
         
         elseif fenlei="中膜" and checkbox=""  then 
         rs3("zh_Loan_num")=Formatnumber((Loan_num/guige/houdou/bizhong),2,-1,0,0)
         rs3("uptime") = request.Form("uptime")
         rs3("cname")= cname
         rs3("pihao") = pihao
                  
         elseif fenlei="中膜" and checkbox="yess" then 
         rs3("zh1_Loan_num")=Formatnumber((Loan_num/guige/houdou/bizhong),2,-1,0,0)
         rs3("uptime") = request.Form("uptime")
         rs3("cname")= cname
         rs3("pihao") = pihao
                  
         elseif fenlei="内膜" then 
         rs3("l_Loan_num")=Formatnumber((Loan_num/guige/houdou/bizhong),2,-1,0,0)
         rs3("uptime") = request.Form("uptime")
         rs3("cname")= cname
         rs3("pihao") = pihao
         end if
      end if
         rs3.update
         rs3.close
         set rs3=nothing
            rs.update
            rs.close
            set rs=nothing
         
        set rs4=server.createobject("adodb.recordset") 
        sql4="select * from mxfuhe"
        rs4.open sql4,conn,1,3
        rs4.addnew
        rs4("mid")=id
        rs4("fenlei") = fenlei
        rs4("guige") = guige
        rs4("houdou") = houdou
        rs4("pinming") = pinming
        rs4("uptime") = request.Form("uptime")
        rs4("Unit") = unit

        rs4("use_dep") = use_dep
        rs4("Loan_num") = Loan_num
        rs4("cname") = cname
        rs4("pihao") =pihao
        rs4("yess")=checkbox
        rs4.update
        rs4.close
        set rs4=nothing

         

else
         response.Write "<script language=javascript>alert('领用车间只能是彩印或复合两车间！');</script>"
         response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_out_store.asp"">"
         response.end


end if


        set rs1=server.createobject("adodb.recordset") 
        sql1="select * from cailiao_store where id="&id&""
        rs1.open sql1,conn,1,3
     if not rs1.eof then
        rs1("number")=rs1("number")-Loan_num
        rs1("anpai")=0
     end if
        rs1.update
        rs1.close
        set rs1=nothing 

         



        set rs=server.createobject("adodb.recordset") 
        sql="select * from cailiao_out_store"
        rs.open sql,conn,1,3
        rs.addnew          
        rs("fenlei") = fenlei
        rs("guige") = guige
        rs("houdou") = houdou
        rs("pinming") = pinming
        rs("uptime") = request.Form("uptime")
        rs("Unit") = unit
        rs("content") = trim(request.Form("content"))
        rs("use_dep") = trim(request.Form("use_dep"))
        rs("Loan_num") = Loan_num
        rs("cname") = trim(request.Form("cname"))
        rs("pihao") =trim(request.Form("pihao"))
        rs("yess")=checkbox
        rs.update
        response.Write "<script language=javascript>alert('出库成功！');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_out_store.asp"">"
        response.end
        rs.close
        set rs=nothing
end sub

 
sub SaveModify 
fenlei= trim(request.Form("d_position1")) 
old_num=request.Form("old_num")
Loan_num=request.Form("Loan_num")
pinming=trim(request.Form("pinming"))
cname=trim(request.Form("cname"))
cname1=trim(request.Form("cname1"))
pihao=trim(request.Form("pihao"))
pihao1=trim(request.Form("pihao1"))
use_dep=trim(request.Form("use_dep"))
use_dep1=trim(request.Form("use_dep1"))
checkbox=trim(request.Form("checkbox"))

names=Split(pinming,"|")

id=names(0)
pinming=names(1)
guige=names(2)
houdou=names(3)
num=names(4)
unit=names(5)


     if  use_dep<>use_dep1 or cname<>cname1 then
     response.Write "<script language=javascript>alert('只有数量才能修改！\n\n如果修改其它内容,\n\n请先数量改为0,再正确输入.');</script>"
     response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_out_store.asp"">"
     response.end
     end if

set rs5=server.createobject("adodb.recordset") 
sql5="select * from cailiao_out_store where id="&request.Form("id")
rs5.open sql5,conn,1,3 
 if rs5("yess")<>checkbox or rs5("pinming")<>pinming then
 response.Write "<script language=javascript>alert('请注意是否对称复合，要与原出库相符！\n\n或者只选择了类别，未选品名，造成不相符。');</script>"
     response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_out_store.asp"">"
     response.end
     end if

rs5.update
rs5.close
set rs5=nothing 
     
num= old_num-Loan_num


    
    
       set rs1=server.createobject("adodb.recordset") 
       sql1="select * from cailiao_store where pinming='"&pinming&"'and guige="&guige&" and houdou="&houdou&""
       rs1.open sql1,conn,1,3

    if not rs1.eof then
       if rs1("number")+num<0 then
          response.Write "<script language=javascript>alert('对不起，库存不足，请调整您的出库量！');</script>"
          response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_out_store.asp"">"
          response.end
       else
          rs1("number")=rs1("number")+num
          rs1("anpai")=0
       end if
     end if
          rs1.update
          rs1.close
          set rs1=nothing
  
if use_dep="彩印车间" then
             set rs=server.createobject("adodb.recordset")
             sql="select * from cailiao where pinming='"&pinming&"'"
             rs.open sql,conn,2,3
             bizhong=rs("bizhong")
      
       set rs2=server.createobject("adodb.recordset") 
       sql2="select * from caiyin where cname='"&cname&"'"
numss=Formatnumber((num/guige/houdou/bizhong),2,-1,0,0)

       rs2.open sql2,conn,2,3
    if not rs2.eof then
       rs2("pinming") = pinming
       rs2("guige") = guige
       rs2("houdou")=houdou
       rs2("Loan_num")=rs2("Loan_num")-num
       rs2("qianmi")=rs2("qianmi")-numss
       rs2("zhemi")=rs2("zhemi")-numss
    end if
       rs2.update
       rs2.close
       set rs2=nothing
            rs.update
            rs.close
            set rs=nothing
            
elseif use_dep="复合车间" then
         set rs=server.createobject("adodb.recordset")
             sql="select * from cailiao where pinming='"&pinming&"'"
             rs.open sql,conn,2,3
             bizhong=rs("bizhong")

       set rs3=server.createobject("adodb.recordset") 
       sql3="select * from fuhe where cname='"&cname&"'"
       rs3.open sql3,conn,2,3
numss=Formatnumber((num/guige/houdou/bizhong),2,-1,0,0)
      if not rs3.eof then
         if fenlei="面膜" then
          rs3("m_Loan_num")=rs3("m_Loan_num")-numss
          rs3("guige") =guige
          rs3("houdou")=houdou        
          rs3("pinming")=pinming
         elseif fenlei="中膜"and checkbox=""   then
         rs3("zh_Loan_num")=rs3("zh_Loan_num")-numss
         elseif fenlei="中膜"  and checkbox="yess" then
         rs3("zh1_Loan_num")=rs3("zh1_Loan_num")-numss
         elseif fenlei="内膜" then 
         rs3("l_Loan_num")=rs3("l_Loan_num")-numss
         end if
      end if
         rs3.update
         rs3.close
         set rs3=nothing
            rs.update
            rs.close
            set rs=nothing

         set rs4=server.createobject("adodb.recordset") 
        if pihao="" then
          sql4="select * from mxfuhe where cname='"&cname&"'"
       else
          sql4="select * from mxfuhe where  pihao='"&pihao&"'and cname='"&cname&"'"
        end if
        rs4.open sql4,conn,1,3
        rs4("fenlei") = fenlei
        rs4("guige") = guige
        rs4("houdou") = houdou
        rs4("pinming") = pinming
        rs4("uptime") = request.Form("uptime")
        rs4("Unit") = unit

        rs4("use_dep") =use_dep
        rs4("Loan_num") = Loan_num
        rs4("cname") = cname
        rs4("pihao") =pihao
        rs4("yess")=checkbox
        rs4.update
        rs4.close
        set rs4=nothing
end if


set rs=server.createobject("adodb.recordset") 
sql="select * from cailiao_out_store where id="&request.Form("id")
rs.open sql,conn,1,3 
rs("fenlei") = fenlei
rs("guige") = guige
rs("houdou") = houdou
rs("pinming") = pinming
rs("Unit") = unit
rs("uptime") = request.Form("uptime")
rs("use_dep") = use_dep
rs("Loan_num") = Loan_num
rs("content") = trim(request.Form("content"))
rs("cname") = cname

rs.update
rs.close
set rs=nothing  
response.Write "<script language=javascript>alert('修改成功！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_out_store.asp"">"
response.end



end sub   
 
  sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_out_store.asp"">"
  response.end
else
        conn.execute("delete from cailiao_out_store where id in ("&Request.Form("selBigClass")&")and loan_num=0")
		response.Write "<script language=javascript>alert('数量为0的已删除成功！\n\n数量不为0的没有被删除！');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=cailiao_out_store.asp"">"
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
		   rs.open "select * from cailiao_out_store where id=" & request("id"),conn,1,1
%>

    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">编辑出库信息</font></b>

</td>
</tr>
</table>
<%
else
%>
    <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">添加出库信息</font></b>

</td>
</tr>
</table>
<%end if %>
        <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">出库日期</td>
            <td width="10%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">类 别</td>
			<td width="60%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			<p align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 品 名 | 规&nbsp; 格 | 余数|&nbsp; 单位</td>
			
			
			<td width="20%" height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">领用车间</td>
			
			
          </tr>
		  <form name="powersearch" method="POST" action="cailiao_out_store.asp" onSubmit="return Juge(this)" style="width: 200" >
		  <tr> <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'> 
		  <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>"> <% End If %>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
			<input name="uptime" onFocus="show_cele_date(uptime,'','',uptime)" type="text"  size="10"  value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'"> 
			<select name='d_position1' onchange='Do_po_Change(this);' valign=top style='width:64; height:17'>
			<% if isedit then%>
				<option selected="selected"><%=rs("fenlei")%></option>
			<%else%>
		        <option selected value='0000'>请选择</option>
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
			<select name='pinming' size=1 style='width:257; height:13'>
			<% if isedit then%>
			<option selected="selected" value="<%=rs("id")%>|<%=rs("pinming")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("Loan_num")%>|<%=rs("unit")%>"><%=rs("pinming")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=rs("unit")%></option>
			<input type="hidden" value="<%=rs("Loan_num")%>" name="old_num">
			<input type="hidden" value="<%=rs("use_dep")%>" name="use_dep1">
			<input type="hidden" value="<%=rs("cname")%>" name="cname1">
			<input type="hidden" value="<%=rs("pihao")%>" name="pihao1">
			<%else%>
			<option>--请选择--</option>
			<%end if%>
			</select>
			</td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
<select name="use_dep"  size="1" style='width:100; height:17'>
			<% if isedit then%><option ><%=rs("use_dep")%></option><%end if%>
			<%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from chejian",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
			%>
			<option <% if isedit then
			if trim(rs3("chejianname"))=trim(rs("use_dep")) then
			%> selected="selected"<%end if%><%end if%>><%=rs3("chejianname")%></option>
			<%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
			%>
			</select></td>

          </tr>
		  <tr class="but">
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'" colspan="2">数量</td>
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">产品名称&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			备注</td>			
          </tr>
		  <tr> 
 <td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'" colspan="2">
 <input name="Loan_num" type="text" style="ime-mode:disabled"  size="19" value='<% if isedit then
		                                                         response.write rs("Loan_num")
																  end if %>'></td>

<td height="18" align="center" class="but11" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but11'">
	      
	      <%
               set rs1=server.createobject("adodb.recordset")
               exec="select * from dindang order by -id"
               rs1.open exec,conn,1,1
               %><select name="cname" id="cname" style='width:480; height:15'>
			<% if isedit then%><option ><%=rs("cname")%></option><%else%>
		        <option selected value='0000'>请选择名称</option>
			<%end if%>

			<%do while not rs1.eof%>
                 <option  value="<%=rs1("pinming")%>"><%=rs1("pinming")%></option>
                 <%
rs1.movenext
loop
rs1.close
set rs1=nothing
%>
                            </select>			</td>																		  
																  





			<td height="18" align="center" class="but" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but'">
			<input name="content" type="text"  size="19" value='<% if isedit then
		                                                         response.write rs("content")
																  end if %>'></td>			



			</tr>
			<tr>
			<td colspan="4" align="center"><input type="submit" value="<% if isedit then
		                                                         response.write "编辑"
																 else
																 response.write "添加"
																  end if %>"> </td>
          </tr></form>
		  
           </table>
        <%end if%>


</td>
</tr>
</table>
<br>
<TABLE width=820 align=center border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6" height="143">
<tr>
<td >



			<form action="cailiao_out.asp" method="post" name="selform" >
  <table width="800"  border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

    <tr> 
      
      <td width="800" align="center" valign="top">
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">材 料 出 库 列 表</font></b>

</td>
</tr>
</table>
	  
        <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
          <tr class="but"> 
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">出库日期</td>
            <td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">类 别</td>
			<td width="8%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">品 名</td>
			<td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">规 格</td>
			<td width="10%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">数 量</td>
			<td width="5%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">单 位</td>
			<td width="9%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">领用车间</td>
			<td width="12%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			产品名称</td>
            <td width="7%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">
			鸳鸯复合</td>
            <td width="8%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">批号</td>
            <td width="6%" height="18" align="center" class="but01" onMouseDown="this.className='tddown'" onMouseUp="this.className='but'" onMouseOut="this.className='but01'">操 作</td>
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
akeyword=request("akeyword")
keyword=request("keyword")
rs.Source="select* from cailiao_out_store where cname like '%"&akeyword&"%' and pinming like '%"&keyword&"%'and loan_num<>0 order by -id"
else
rs.Source="select* from cailiao_out_store where  loan_num<>0 order by -id"
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
<td height="18" align="center"><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><a href="cailiao_out_store.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
<td height="18"  align="center"><%=rs("guige")%>*<%=rs("houdou")%> </td>
<td height="18" align="center"><%=Formatnumber(rs("Loan_num"),2,-1,0,0)%><%if rs("Loan_num")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%=rs("Unit")%><%if rs("Unit")="" then%>&nbsp;<%end if%></td>

<td height="18" align="center"><%if rs("use_dep")<>"" then%><%=rs("use_dep")%><%else%>调为<%=rs("kufang")%><%=rs("old_fenlei")%><%end if%></td>

<td height="18" align="center"><%=rs("cname")%><%if rs("cname")="" then%>&nbsp;<%end if%></td>
<td height="18" align="center"><%if rs("yess")<>"" then%>是<%else%>&nbsp;<%end if%></td>
<td height="18" align="center"><%if rs("pihao")<>"" then%><%=rs("pihao")%><br><u><%=rs("content")%></u><%else%><%=rs("content")%><%end if%></td>
<td height="18" align="center"><input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
</tr>
<%
rs.MoveNext
end if
next
end if
%>
<tr><td align="right" colspan="11" height="22">
共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="cailiao_out_store.asp?keyword="&keyword&"&akeyword="&akeyword&"&search=search&"
			else
            url="cailiao_out_store.asp?"
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
%>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper"or oskey="cailiao" or oskey="admin" then%>
     <input type="checkbox" name="checkbox" value="checkbox" onClick="javascript:SelectAll()"> 选择/反选
              <input type=submit value=打印 name=action> 
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
<TABLE width=820 align=center border="1" cellspacing="20" cellpadding="0" bordercolor="#0055E6">
<tr>
<td>
	        <table width="800" border="0" cellpadding="0" cellspacing="0" >
          <tr >  
<td height=26 background="images/background.gif">
<b><font color="#ffffff">出库搜索</font></b>

</td>
</tr>
</table>

        <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">
         
		  <form name="search" method="post" action="cailiao_out_store.asp">
		  <tr><input name="search" type="hidden" value="search">
             <td align="center" width="36%">按品名查找： 
				<select name='keyword'  style='width:130'>
			<option selected value="">请选择品名</option>
			<%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(pinming) from cailiao_out_store where guige<>0 order by pinming",conn,2,2
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
                      
			<td align="center" width="13%" ><input type="submit"  value="查 找"> </td>
			<td align="center" width="35%">按产品查找： 
			<input name="akeyword" type="text"  size="22"></td>
                      
			<td align="center" width="17%" ><input type="submit"  value="查 找"> </td>

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
