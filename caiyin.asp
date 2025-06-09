<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
    <title>∷嘉美管理系统∷</title>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <link href="css/main.css" rel="stylesheet" type="text/css">

    <script language="JavaScript" src="css/User_Info_Modify.js"></script>
    <style type="text/css">
        <!--
        td {
            font-family: "宋体";
            font-size: 9pt
        }

        body {
            font-family: "宋体";
            font-size: 9pt
        }

        select {
            font-family: "宋体";
            font-size: 9pt
        }

        A {
            text-decoration: none;
            color: #336699;
            font-family: "宋体";
            font-size: 9pt
        }

            A:hover {
                text-decoration: underline;
                color: #FF0000;
                font-family: "宋体";
                font-size: 9pt
            }
        -->
    </style>

    <script>
// 记录页面开始加载时间
var startTime = new Date().getTime();

window.onload = function() {
    var endTime = new Date().getTime();
    var loadTime = (endTime - startTime) / 1000;
    document.getElementById("loadTime").innerHTML = "完整页面加载时间: " + loadTime.toFixed(2) + " 秒";
};
    </script>


    <script language="javascript">
<!--
    function Juge(powersearch) {
        if (powersearch.Loan_oan.value == "") {
            alert("数量不能为空！");
            powersearch.Loan_oan.focus();
            return (false);
        }

    }


    function SelectAll() {
        for (var i = 0; i < document.selform.selBigClass.length; i++) {
            var e = document.selform.selBigClass[i];
            e.checked = !e.checked;
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
rs2.open "select * from caiyin where fenlei='"&fenleiname&"' order by qianmi desc",conn,1,1
if not rs2.eof then   
do while not rs2.eof
%>
po_detail_show[<%=i%>][<%=j%>] = '<%=rs2("cname")%>|<%=rs2("guige")%>|<%=rs2("houdou")%>|<%=Formatnumber(rs2("qianmi"),2,-1,0,0)%>|km|<%=rs2("pihao")%>|<%=rs2("pinming")%>';
po_detail_value[<%=i%>][<%=j%>] = '<%=rs2("id")%>|<%=rs2("cname")%>|<%=rs2("guige")%>|<%=rs2("houdou")%>|<%=Formatnumber(rs2("qianmi"),2,-1,0,0)%>|km|<%=rs2("pihao")%>|<%=rs2("pinming")%>';
<%
rs2.movenext
j=j+1
loop
else
%>
po_detail_show[<%=i%>][<%=j%>] = '暂无半成品出库';
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
    <script language="JavaScript">



        var psid = "";

        function DoLoad(form, funtypev) {
            var n;
            var i, j, k;
            var num;
            num = GetObjID('funtype[]');
            num1 = GetObjID('cname');

            if (!funtypev)
                return;
            k = 0;
            for (i = 0; i < po_ca_show.length; i++) {
                for (j = 0; j < po_detail_value[i].length; j++) {
                    if (funtypev.indexOf(po_detail_value[i][j]) != -1) {
                        NewOptionName = new Option(po_detail_show[i][j], po_detail_value[i][j]);
                        form.elements[num].options[k] = NewOptionName;
                        k++;
                    }
                }
            }
        }

        function Do_po_Change(form) {
            var num, n, i, m;
            num = GetObjID('d_position1');
            m = document.powersearch.elements[num].selectedIndex - 1;
            n = document.powersearch.elements[num + 1].length;
            for (i = n - 1; i >= 0; i--)
                document.powersearch.elements[num + 1].options[i] = null;

            if (m >= 0) {
                for (i = 0; i < po_detail_show[m].length; i++) {
                    NewOptionName = new Option(po_detail_show[m][i], po_detail_value[m][i]);
                    document.powersearch.elements[num + 1].options[i] = NewOptionName;
                }
                document.powersearch.elements[num + 1].options[0].selected = true;
            }
        }

        function InsertItem(ObjID, Location) {
            len = document.powersearch.elements[ObjID].length;
            for (counter = len; counter > Location; counter--) {
                Value = document.powersearch.elements[ObjID].options[counter - 1].value;
                Text2Insert = document.powersearch.elements[ObjID].options[counter - 1].text;
                document.powersearch.elements[ObjID].options[counter] = new Option(trimPrefixIndent(Text2Insert), Value);
            }
        }
        function GetLocation(ObjID, Value) {
            total = document.powersearch.elements[ObjID].length;
            for (pp = 0; pp < total; pp++)
                if (document.powersearch.elements[ObjID].options[pp].text == "---" + Value + "---") {
                    return (pp);
                    break;
                }
            return (-1);
        }

        function GetObjID(ObjName) {
            for (var ObjID = 0; ObjID < window.powersearch.elements.length; ObjID++)
                if (window.powersearch.elements[ObjID].name == ObjName) {
                    return (ObjID);
                    break;
                }
            return (-1);
        }




    </script>
</head>

<!--<div id="loadTime">-->

<body topmargin="0" rightmargin="0" leftmargin="0">
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
use_dep=trim(request.Form("use_dep")) 
Loan_oan=request.Form("Loan_oan")
cname=trim(request.Form("cname"))
unit=trim(request.Form("unit"))
names=Split(cname,"|")
id=names(0)
cname=names(1)
guige=names(2)
houdou=names(3)
qianmi=names(4)
pihao=names(6)
pinming=names(7)

if use_dep="复合车间" and unit<>"千米" then
                response.Write "<script language=javascript>alert('向复合出库，只能是千米！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin.asp"">"
                response.end
elseif use_dep="熟化室" and unit<>"千米" then
                response.Write "<script language=javascript>alert('向熟化室出库，只能是千米！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin.asp"">"
                response.end
elseif use_dep="退库" and unit<>"公斤" then
                response.Write "<script language=javascript>alert('退库，只能是千米！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin.asp"">"
                response.end
end if

if use_dep="复合车间" then
          set rs=server.createobject("adodb.recordset")
          if pihao="" then
          sql="select * from fuhe where cname='"&cname&"'"
          else
          sql="select * from fuhe where pihao='"&pihao&"'and cname='"&cname&"'"
          end if
        rs.open sql,conn,2,3
    if not rs.eof and not rs.bof then

         rs("m_Loan_num")=rs("m_Loan_num")+Loan_oan
         rs("guige")=guige
         rs("fenlei")=fenlei
         rs("houdou")= houdou
         rs("pinming") = pinming
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
       

       set rs3=server.createobject("adodb.recordset") 
           sql3="select top 1 * from caiyin_in"
           rs3.open sql3,conn,1,3
           rs3.addnew
           rs3("unit")=unit
           rs3("loan_oan")=Loan_oan
           rs3("guige") =guige
           rs3("cname") =cname
           rs3("uptime") = request.Form("uptime") 
           rs3("houdou")=houdou
           rs3("pihao") =pihao
           rs3("use_dep")=use_dep

           rs3.update
           rs3.close
           set rs3=nothing     
        
        if pinming<>"存料" then
          set rs4=server.createobject("adodb.recordset") 
           sql4="select * from dindang where pinming='"&cname&"'"
            rs4.open sql4,conn,1,2
        if not rs4.eof then
           rs4("caiyin")="ok"
         end if
          rs4.update
          rs4.close
       end if

           set rs1=server.createobject("adodb.recordset") 
           sql1="select * from caiyin where id="&id&""
            rs1.open sql1,conn,1,2

        if not rs1.eof then
           rs1("Loan_oan")=rs1("Loan_oan")+Loan_oan
           rs1("qianmi")=rs1("qianmi")-Loan_oan 
           rs1("use_dep") = use_dep
        end if
        rs1.update
                response.Write "<script language=javascript>alert('出库成功！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin.asp"">"
                response.end
         rs1.close
         set rs1=nothing
         
elseif use_dep="熟化室" then
             set rs5=server.createobject("adodb.recordset") 
             if pihao=""  then
             sql5="select * from matur where cname='"&cname&"'"
             else
             sql5="select * from matur where  pihao='"&pihao&"'and cname='"&cname&"'"
             end if
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
         
          set rs3=server.createobject("adodb.recordset") 
           sql3="select * from caiyin_in"
           rs3.open sql3,conn,1,3
           rs3.addnew
           rs3("unit")=unit
           rs3("loan_oan")=Loan_oan
           rs3("guige") =guige
           rs3("cname") =cname
           rs3("uptime") = request.Form("uptime") 
           rs3("houdou")=houdou

           rs3("use_dep")=use_dep
           rs3.update
           rs3.close
           set rs3=nothing
           
          
          set rs4=server.createobject("adodb.recordset") 
           sql4="select * from dindang where pinming='"&cname&"'"
            rs4.open sql4,conn,1,2
        if not rs4.eof then
           rs4("caiyin")="ok"
         end if
          rs4.update
          rs4.close

         
          set rs1=server.createobject("adodb.recordset") 
           sql1="select * from caiyin where id="&id&""
            rs1.open sql1,conn,1,2

         if not rs1.eof and not rs1.bof then
           rs1("Loan_oan")=rs1("Loan_oan")+Loan_oan
           rs1("qianmi")=rs1("qianmi")-Loan_oan 
           rs1("use_dep") = use_dep
        end if
        rs1.update
                response.Write "<script language=javascript>alert('出库成功！');</script>"
                response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin.asp"">"
                response.end
         rs1.close
         set rs1=nothing 
end if

if use_dep="退库" then
           set rs6=server.createobject("adodb.recordset") 
           sql6="select * from cailiao_in_store"
           rs6.open sql6,conn,1,3
           rs6.addnew
           rs6("fenlei") = trim(request.Form("d_position1"))
           rs6("guige") = guige
           rs6("houdou") = houdou
           rs6("pinming") = pinming
           rs6("uptime") = request.Form("uptime")
           rs6("Unit") = trim(request.Form("unit"))
           rs6("gonghuodanwei") = "彩印退库"
           rs6("use_num") = request.Form("Loan_oan")
           rs6.update
           rs6.close
           set rs6=nothing

            
            set rs2=server.createobject("adodb.recordset") 
            sql2="select * from cailiao_store where pinming='"&pinming&"'and guige="&guige&" and houdou="&houdou&" and fenlei='"&fenlei&"' "
            rs2.open sql2,conn,1,3
          if not rs2.eof and not rs2.bof then
            rs2("number")=rs2("number")+Loan_oan
          else
            rs2.addnew
            rs2("fenlei") = trim(request.Form("d_position1"))
            rs2("guige") = guige
            rs2("houdou")=houdou
            rs2("pinming") = pinming
            rs2("number")=Loan_oan
            rs2("unit")=unit
          end if  
           rs2.update
           rs2.close
           set rs2=nothing
           
            set rs=server.createobject("adodb.recordset")
             sql="select * from cailiao where pinming='"&pinming&"'"
             rs.open sql,conn,2,3
             bizhong=rs("bizhong")
                     set rs1=server.createobject("adodb.recordset") 
           sql1="select * from caiyin where id="&id&""
    %><%   
            rs1.open sql1,conn,1,2
            if not rs1.eof then
               rs1("oan_num")=rs1("oan_num")+Loan_oan
               rs1("loan_num")=rs1("loan_num")-loan_oan
               rs1("zhemi")=rs1("zhemi")-(Loan_oan/guige/houdou/bizhong)
               rs1("qianmi")=rs1("qianmi")-(Loan_oan/guige/houdou/bizhong)
            end if 
            rs1.update
            rs1.close
            set rs1=nothing
            rs.update
            rs.close
            set rs=nothing
            response.Write "<script language=javascript>alert('退库成功！');</script>"
            response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin.asp"">"
            response.end
else

          response.Write "<script language=javascript>alert('彩印只能向退库、熟化室、复合车间出库！');</script>"
          response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin.asp"">"
          response.end
end if
end sub

sub SaveModify 
end sub   
  sub delCate() 
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
          response.Write "<script language=javascript>alert('您没有选定记录！');</script>"
          response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin.asp"">"
          response.end
  else
        conn.execute("delete from caiyin where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('已删除成功！');</script>"
        response.write "<meta http-equiv=""refresh"" content=""0;url=caiyin.asp"">"
        response.end
  end if
    end sub

    %> <% sub myform(isEdit) %>
    <table width="800" align="center" border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
        <tbody>
            <tr>
                <td align="center" valign="top"><%if oskey="supper" or oskey="cailiao"  or oskey="admin" then%>

                    <%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from caiyin where id=" & request("id"),conn,1,1
                    %>

                    <table width="800" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td height="26" background="images/background.gif">
                                <b><font color="#ffffff">编辑彩印车间出库信息</font></b>

                            </td>
                        </tr>
                    </table>
                    <%else%>
                    <table width="800" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td height="26" background="images/background.gif">
                                <b><font color="#ffffff">添加彩印车间出库信息</font></b>

                            </td>
                        </tr>
                    </table>
                    <%end if %>
                    <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
                        <tr class="but">
                            <td width="10%" height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'">出库日期</td>
                            <td width="9%" height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'">类 别</td>
                            <td width="60%" height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'">
                                <p align="left">
                                    &nbsp;&nbsp;&nbsp;&nbsp; 品 名 | 规&nbsp; 格 | 余数|&nbsp; 单位|&nbsp; 批号</td>
                            <td width="12%" height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'">领用车间</td>

                            <td width="9%" height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'">单位</td>

                        </tr>
                        <form name="powersearch" method="post" action="caiyin.asp" onsubmit="return Juge(this)">
                            <tr>
                                <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'>
                                <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>">
                                <% End If %>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="uptime" onfocus="show_cele_date(uptime,'','',uptime)" type="text" size="10" value='<% if isedit then
		                                                         response.write rs("uptime")
																 else
																 response.write date()
																  end if %>'></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name='d_position1' onchange='Do_po_Change(this);' valign="top" style='width: 61; height: 25'>
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
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name='cname' size="1" style='width: 481; height: 15'>
                                        <% if isedit then%>
                                        <option selected="selected" value="0|<%=rs("cname")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=Formatnumber(rs("qianmi"),2,-1,0,0)%>|km|<%=rs("pihao")%>|<%=rs("pinming")%>"><%=rs("cname")%>|<%=rs("guige")%>|<%=rs("houdou")%>|<%=Formatnumber(rs("qianmi"),2,-1,0,0)%>|km|<%=rs("pihao")%>|<%=rs("pinming")%></option>
                                        <input type="hidden" value="<%=rs("Loan_oan")%>" name="old_num">
                                        <input type="hidden" value="<%=rs("use_dep")%>" name="old_dep">
                                        <%else%>
                                        <option>--请选择--</option>
                                        <%end if%>
                                    </select>
                                </td>

                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="use_dep" size="1" style='width: 68; height: 16'>
                                        <option selected value='熟化室'>熟化室</option>
                                        <option selected value='退库'>退库</option>
                                        <option selected value='复合车间'>复合车间</option>

                                    </select></td>

                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="unit" size="1" style='width: 50; height: 15'>
                                        <option selected value='公斤'>公斤</option>
                                        <option selected value='千米'>千米</option>
                                    </select></td>
                            </tr>
                            <tr class="but">
                                <td height="18" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'" colspan="2">数&nbsp;&nbsp; 量</td>
                                <td height="36" align="center" class="but" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but'" colspan="3" rowspan="2">
                                    <p align="left">
                                        说明:1、出库,数量直接输入千米数，领用车间选择复合车间；如果不复合，请选择&quot;熟化室&quot;出库&nbsp;。 
			<br>
                                        &nbsp;&nbsp;&nbsp;&nbsp; 2、材料退回，选择正确品名,数量直接输入公斤数;领用车间选择“退库”，单位选择“公斤”。<br>
                                        &nbsp;&nbsp;&nbsp;&nbsp; 3、出库输入有误请在“出库明细”中修改，退库有误只能输入负数冲减。&nbsp;&nbsp;&nbsp;&nbsp; </td>
                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <input name="Loan_oan" type="text" style="ime-mode: disabled" size="11" value='<% if isedit then
		                                                         response.write rs("Loan_oan")
																  end if %>'><font color="#FFFFFF">千米(公斤)</font></td>



                            </tr>
                            <tr>
                                <td colspan="5" align="center">
                                    <input type="submit" value="<% if isedit then
		                                                         response.write "编辑"
																 else
																 response.write "添加"
																  end if %>">
                                </td>
                            </tr>
                        </form>

                    </table>
                    <%end if%>


                </td>
            </tr>
    </table>
    <br>
    <br>
    <table width="820" align="center" border="0" cellspacing="20" cellpadding="2" bordercolor="#0055E6" height="143">
        <tr>
            <td>



                <form action="caiyin.asp" method="post" name="selform">
                    <table width="800" border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

                        <tr>

                            <td width="800" align="center" valign="top">
                                <table width="800" border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td height="26" background="images/background.gif">
                                            <b><font color="#ffffff">印&nbsp; 刷&nbsp; 明&nbsp; 细&nbsp; 列&nbsp; 表</font></b>

                                        </td>
                                    </tr>
                                </table>

                                <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
                                    <tr class="but">
                                        <td width="8%" height="36" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">出库日期</td>
                                        <td width="4%" height="36" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">类 别</td>
                                        <td width="31%" height="36" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">产品名称</td>
                                        <td width="8%" height="36" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">规 格</td>
                                        <td width="11%" height="18" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" colspan="2">领&nbsp;&nbsp;&nbsp;&nbsp; 用</td>
                                        <td width="7%" height="36" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">出库数</td>
                                        <td width="7%" height="36" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">现存</td>
                                        <td width="5%" height="36" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">退库</td>
                                        <td width="6%" height="36" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">损耗率</td>
                                        <td width="7%" height="36" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">领用车间</td>
                                        <td width="6%" height="36" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'" rowspan="2">操 作</td>
                                    </tr>
                                    <tr class="but">
                                        <td width="6%" height="18" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'">公斤</td>
                                        <td width="5%" height="18" align="center" class="but01" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but01'">千米</td>
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

keyword=request("keyword")
rs.Source="select top 1000 * from caiyin where cname like '%"&keyword&"%' order by id desc "
else
rs.Source="select top 1000 * from caiyin order by id desc "
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
                                    <!--<tr "onmouseover="this.bgColor='#73A475';" onmouseout="this.bgColor=''"">-->
                                    <tr onmouseover="this.style.backgroundColor='#73A475';" onmouseout="this.style.backgroundColor='';">
                                        <td height="18" align="center"><%=rs("uptime")%></td>
                                        <td height="18" align="center"><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
                                        <td height="18" align="center"><%=rs("cname")%></td>
                                        <td height="18" align="center"><%=rs("guige")%>*<%=rs("houdou")%></td>
                                        <td height="18" align="center"><%if rs("Loan_num")="" then%>&nbsp;<%else%><%=rs("Loan_num")%><%end if%></td>
                                        <td height="18" align="center"><%if rs("zhemi")="0" then%>&nbsp;<%else%><%=Formatnumber(rs("zhemi"),2,-1,0,0)%><%end if%></td>
                                        <td height="18" align="center"><%if rs("Loan_oan")="0" then%><font color="#FF0000">在印</font><%else%><%=rs("Loan_oan")%>km<%end if%></td>
                                        <td height="18" align="center" bgcolor="#FFFFCC"><%if rs("Loan_oan")="0" then%>&nbsp;<%else%><%=Formatnumber(rs("qianmi"),2,-1,0,0)%>km<%end if%></td>
                                        <td height="18" align="center"><%if rs("oan_num")="0" then%>&nbsp;<%else%><%=rs("oan_num")%>kg<%end if%></td>
                                        <td height="18" align="center"><%if rs("Loan_oan")="0" then%>&nbsp;<%elseif rs("zhemi")="0" then%>&nbsp;<%else%><%=Formatnumber(rs("qianmi")/rs("zhemi")*100,2,-1,0,0)%>%<%end if%></td>
                                        <td height="18" align="center"><%if rs("use_dep")="" then%>&nbsp;<%else%><%=rs("use_dep")%><%end if%></td>
                                        <td height="18" align="center">
                                            <input name="selBigClass" type="checkbox" id="selBigClass" value="<%=rs("ID")%>"></td>
                                    </tr>
                                    <%
rs.MoveNext
end if
next
end if
                                    %>
                                    <tr>
                                        <td align="center" colspan="12" height="22">共 <%=total%> 条，当前第 <%=Mypage%>/<%=Maxpages%> 
            页，每页 <%=MyPageSize%> 条 
            <%
			if search<>"" then
			url="caiyin.asp?keyword="&keyword&"&search=search&"
			else
            url="caiyin.asp?"
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
                                            <input type="checkbox" name="checkbox" value="checkbox" onclick="javascript:SelectAll()">
                                            选择/反选
              <input onclick="{if(confirm('确定要删除吗？')){this.document.selform.submit();return true;}return false;}" type="submit" value="删除" name="action2">
                                            <input type="Hidden" name="action" value='del'><%end if%></td>
                                    </tr>
                                </table>


                            </td>
                        </tr>

                    </table>


                </form>

            </td>
        </tr>
    </table>
    <br>
    <br>
    <table width="820" align="center" border="1" cellspacing="20" cellpadding="0" bordercolor="#0055E6">
        <tr>
            <td>
                <table width="800" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                        <td height="26" background="images/background.gif">
                            <b><font color="#ffffff">出库搜索</font></b>

                        </td>
                    </tr>
                </table>

                <table width="90%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">

                    <form name="search" method="post" action="caiyin.asp">
                        <tr>
                            <input name="search" type="hidden" value="search">
                            <td height="40" valign="middle" align="center" width="20%"></td>
                            <td align="center" width="65%">品名关键字：
                                <input name="keyword" type="text" size="25">
                                关键字为空则搜索所有</td>

                            <td align="center" width="15%">
                                <input type="submit" value="查 找">
                            </td>
                        </tr>
                    </form>
            </td>
        </tr>
    </table>

    <table width="100%" border="0" cellspacing="0" cellpadding="4">
    </table>
    </TBODY>
</TABLE>
    <%end sub%>
    <p>
        <table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
            <tbody>
                <tr valign="center">
                    <td align="middle" width="100%">
                        <!--#include file="footer.htm"-->
                    </td>
                </tr>
            </tbody>
        </table>
    </p>
</body>
</html>

<!--</div>-->


