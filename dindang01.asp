<!--#include file="conn.asp"-->
<!--#include file="checkuser.asp"-->
<html>
<head>
    <title>�˼�������ϵͳ��</title>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312">
    <link href="css/main.css" rel="stylesheet" type="text/css">
    <script language="JavaScript" src="css/User_Info_Modify.js"></script>
    <style type="text/css">
        <!--
        td {
            font-family: "����";
            font-size: 9pt
        }

        body {
            font-family: "����";
            font-size: 9pt
        }

        select {
            font-family: "����";
            font-size: 9pt
        }

        A {
            text-decoration: none;
            color: #336699;
            font-family: "����";
            font-size: 9pt
        }

            A:hover {
                text-decoration: underline;
                color: #FF0000;
                font-family: "����";
                font-size: 9pt
            }
        -->
    </style>
    <script language="javascript">
<!--
    function Juge(powersearch) {
        if (powersearch.uptime.value == "") {
            alert("��ѡ�����ڣ�");
            powersearch.uptime.focus();
            return (false);
        }

        if (powersearch.pinming.value == "") {
            alert("Ʒ������Ϊ�գ�");
            powersearch.pinming.focus();
            return (false);
        }
        if (powersearch.chang.value == "") {
            alert("�����Ϊ�գ�");
            powersearch.chang.focus();
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


</head>
<body topmargin="0" rightmargin="0" leftmargin="0">
Response.Write "<!-- Executed in " & Timer() - startTime & " seconds -->"
    <!--#include file="top.asp"-->
    <% dim sql,rs
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

response.Write "<script language=javascript>alert('��ӵ�����¼�룡');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=dindang.asp"">"
response.end

end sub

 
sub SaveModify
fenlei=request.Form("fenlei")
doupian=request.Form("file1")  
pinming=trim(request.Form("pinming"))
chang=trim(request.Form("chang"))
kuang=trim(request.Form("kuang"))
kehu=request.Form("kehu")
uptime=request.form("uptime")
diangjia=request.form("diangjia")
bian=trim(request.Form("bian"))
di=trim(request.Form("di"))
checkbox=trim(request.Form("checkbox"))

set rs=server.createobject("adodb.recordset") 
sql="select * from dindang where id="&request.Form("id")
rs.open sql,conn,1,3
rs("uptime") = uptime
rs("fenlei") = fenlei
rs("chang") = chang 
rs("kuang")=kuang
rs("pinming") = pinming
rs("kehu")=kehu
rs("doupian")=doupian
rs("pihao")=request.form("pihao")
rs("name1")=request.form("name1")
rs("nianbang")=request.form("nianbang")
rs("cunshu")=request.form("cunshu")
rs("nianti")=request.form("nianti")
rs("unit1")=request.form("unit1")
rs("chati")=replace(request("chati"),vbCrlf,"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
rs("name2")=request.form("name2")
rs("nianbang2")=request.form("nianbang2")
rs("cunshu2")=request.form("cunshu2")
rs("nianti2")=request.form("nianti2")
rs("unit2")=request.form("unit2")
rs("chati2")=replace(request("chati2"),vbCrlf,"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
rs("name3")=request.form("name3")
rs("nianbang3")=request.form("nianbang3")
rs("cunshu3")=request.form("cunshu3")
rs("nianti3")=request.form("nianti3")
rs("unit3")=request.form("unit3")
rs("chati3")=replace(request("chati3"),vbCrlf,"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
rs("name4")=request.form("name4")
rs("nianbang4")=request.form("nianbang4")
rs("cunshu4")=request.form("cunshu4")
rs("nianti4")=request.form("nianti4")
rs("unit4")=request.form("unit4")
rs("chati4")=replace(request("chati4"),vbCrlf,"<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;")
rs("diangjia")=diangjia
rs("di")=di
rs("bian")=bian
if checkbox="yess" then
rs("anpai1")="0"
rs("anpai2")="0"
rs("anpai3")="0"
rs("anpai4")="0"
rs("shueming1")=""
rs("shueming2")=""
rs("shueming3")=""
rs("shueming4")=""
rs("uptime")=date()
end if
rs("zhuantai")="ok"
rs.update
rs.close
set rs=nothing

response.Write "<script language=javascript>alert('������ӳɹ���');</script>"
response.write "<meta http-equiv=""refresh"" content=""0;url=dindang.asp"">"
response.end

end sub   
 
  sub delCate()
  selBigClass=Request.Form("selBigClass")
  if selBigClass="" then 
  response.Write "<script language=javascript>alert('��û��ѡ����¼��');</script>"
  response.write "<meta http-equiv=""refresh"" content=""0;url=dindang.asp"">"
  response.end
else
        conn.execute("delete from dindang where id in ("&Request.Form("selBigClass")&")")
		response.Write "<script language=javascript>alert('ɾ���ɹ���');</script>"

		response.write "<meta http-equiv=""refresh"" content=""0;url=dindang.asp"">"
		end if
response.end
  end sub
    %> <% sub myform(isEdit) %>
    <table width="800" align="center" border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
        <tbody>
            <tr>
                <td align="center" valign="top"><%if oskey="supper" or oskey="admin" then%>

                    <%
	    
	   if isedit then
	   set rs=server.createobject("adodb.recordset")
		   rs.open "select * from dindang where id=" & request("id"),conn,1,1
                    %>

                    <table width="800" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td height="26" background="images/background.gif">
                                <b><font color="#ffffff">�༭������Ϣ</font></b>

                            </td>
                        </tr>
                    </table>
                    <%
else
                    %>
                    <table width="800" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td height="26" background="images/background.gif">
                                <b><font color="#ffffff">��Ӷ�����Ϣ</font></b>

                            </td>
                        </tr>
                    </table>
                    <%end if %>
                    <table width="800" border="0" cellpadding="0" cellspacing="0" bordercolor="#0055E6">
                        <tr class="but">
                            <td width="9%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">�������</td>
                            <td width="8%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">�ͻ�</td>
                            <td width="9%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">�� ��</td>
                            <td width="18%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">Ʒ ��</td>
                            <td width="4%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">�� ��</td>
                            <td width="4%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">�� ��</td>
                            <td width="6%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">�۱�</td>
                            <td width="6%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">�۵�</td>
                            <td width="4%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">����</td>
                            <td width="4%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">�� ��</td>

                            <td width="4%" height="18" align="center" class="but" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but'">���µ�</td>

                        </tr>
                        <form name="powersearch" method="post" action="dindang.asp" onsubmit="return Juge(this)">
                            <tr>
                                <input type="Hidden" name="action" value='<% If isedit then%>modify<% Else %>add<% End If %>'>
                                <% If isedit then%><input type="Hidden" name="id" value="<%=rs("id")%>">
                                <% End If %>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="uptime" onfocus="show_cele_date(uptime,'','',uptime)" type="text" size="10" value='<% if isedit then
		                                                        
																
																 response.write rs("uptime")
																
																  end if %>'></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="kehu" size="1" style='width: 65; height: 16'>
                                        <% if isedit then%>
                                        <option selected="selected"><%=rs("kehu")%></option>
                                        <%else%>
                                        <option selected value="">���пͻ�</option>
                                        <%end if%>

                                        <%
	       set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from kehu  order by kehu",conn,1,1
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
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="pihao" type="text" size="12" value='<% if isedit then
		                                                         response.write rs("pihao")
																  end if %>'></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="pinming" type="text" size="28" value='<% if isedit then
		                                                         response.write rs("pinming")
																  end if %>'
                                        maxlength="42"></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="chang" type="text" style="ime-mode: disabled" size="5" value='<% if isedit then
		                                                         response.write rs("chang")
																  end if %>'></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="kuang" type="text" style="ime-mode: disabled" size="5" value='<% if isedit then
		                                                         response.write rs("kuang")
																  end if %>'></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="bian" type="text" style="ime-mode: disabled" size="5" value='<% if isedit then
		                                                         response.write rs("bian")
		                                                         else response.write 0
																  end if %>'></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="di" type="text" style="ime-mode: disabled" size="5" value='<% if isedit then
		                                                         response.write rs("di")
		                                                         else response.write 0
																  end if %>'></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="diangjia" type="text" style="ime-mode: disabled" size="5" value='<% if isedit then
		                                                         response.write rs("diangjia")
		                                                         else response.write 0
																  end if %>'></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="fenlei" size="1" style='width: 58; height: 16'>
                                        <% if isedit then%>
                                        <option><%=rs("fenlei")%></option>
                                        <%end if%>

                                        <%			
			set rs1=server.createobject("adodb.recordset")
		   rs1.open "select * from fenlei1",conn,1,1
		   if not rs1.eof then
		   do while not rs1.eof
                                        %>
                                        <option <% if isedit then
			if trim(rs1("fenlei"))=trim(rs("fenlei")) then
			%>
                                            selected="selected" <%end if%><%end if%>><%=rs1("fenlei")%></option>
                                        <%
			rs1.movenext
			loop
			end if
			rs1.close
			set rs1=nothing
                                        %>
                                    </select></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <% if isedit then%><input type="checkbox" name="checkbox" value="yess"><%else%><input type="checkbox" name="checkbox" value=""><%end if%></td>


                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">��ӡ��Ĥ</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">��Ĥ���</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">��Ĥ���</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">
                                        <input class="iText" type="text" name="file1" size="20" value='<% if isedit then
		                                                         response.write rs("doupian")
		                                                         else  response.write "�ϴ���ά��"

																 end if %>'>
                                        <input class="iButton" type="button" onclick="window.open('Upload.asp','_blank','Width=480 Height=50 top=300 left=300 status=1');" value="�ϴ�" /></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <font color="#FFFFFF">��λ</font></td>
                                <td height="12" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="5">
                                    <font color="#FFFFFF">ӡˢҪ��</font></td>


                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name='name1' style='width: 55; height: 16'>
                                        <% if isedit then%>
                                        <option selected value='0'>��ѡ��</option>
                                        <option selected="selected"><%=rs("name1")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from cailiao where fenlei='��Ĥ' order by pinming",conn,2,2
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
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="nianbang" size="1" style='width: 54; height: 18'>
                                        <% if isedit then%>
                                        <option selected value='0'>��ѡ��</option>
                                        <option><%=rs("nianbang")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from guige order by guige",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
                                        %>
                                        <option <% if isedit then
			if rs3("guige")=rs("nianbang") then
			%>
                                            selected="selected" <%end if%><%end if%>><%=rs3("guige")%></option>
                                        <%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="cunshu" size="1" style='width: 55; height: 16'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option><%=rs("cunshu")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%
			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select * from houdou order by houdou",conn,1,1
		   if not rs4.eof then
		   do while not rs4.eof
                                        %>
                                        <option <% if isedit then
			if rs4("houdou")=rs("cunshu") then
			%>
                                            selected="selected" <%end if%><%end if%>><%=rs4("houdou")%></option>
                                        <%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="nianti" type="text" style="ime-mode: disabled" size="18" value='<% if isedit then
		                                                         response.write rs("nianti")
		                                                         else response.write 0
																  end if %>'></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <select name='unit1' size="1" style='width: 64; height: 16'>
                                        <% if isedit then%>
                                        <option selected value='0'>��ѡ��</option>
                                        <option selected="selected"><%=rs("unit1")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <option value='��'>��</option>
                                        <option value='��'>��</option>
                                        <option value='ǧ��'>ǧ��</option>
                                        <option value='����'>����</option>


                                    </select>
                                </td>
                                <td height="36" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="5" rowspan="3">
                                    <textarea name="chati" cols="35" rows="3"><% if isedit then
		                                                         response.write rs("chati")
																  end if %></textarea></td>


                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">������Ĥ</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">��Ĥ���</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">��Ĥ���</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">����</font></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <font color="#FFFFFF">��λ</font></td>


                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name='name2' style='width: 55; height: 16'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option selected="selected"><%=rs("name2")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from cailiao where fenlei='��Ĥ' order by pinming",conn,2,2
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
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="nianbang2" size="1" style='width: 54; height: 18'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option><%=rs("nianbang2")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from guige order by guige",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
                                        %>
                                        <option <% if isedit then
			if rs3("guige")=rs("nianbang2") then
			%>
                                            selected="selected" <%end if%><%end if%>><%=rs3("guige")%></option>
                                        <%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="cunshu2" size="1" style='width: 55; height: 16'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option><%=rs("cunshu2")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%
			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select * from houdou order by houdou",conn,1,1
		   if not rs4.eof then
		   do while not rs4.eof
                                        %>
                                        <option <% if isedit then
			if rs4("houdou")=rs("cunshu2") then
			%>
                                            selected="selected" <%end if%><%end if%>><%=rs4("houdou")%></option>
                                        <%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="nianti2" type="text" style="ime-mode: disabled" size="18" value='<% if isedit then
		                                                         response.write rs("nianti2")
		                                                         else response.write 0
																  end if %>'></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <select name='unit2' size="1" style='width: 64; height: 16'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option selected="selected"><%=rs("unit2")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <option value='��'>��</option>
                                        <option value='��'>��</option>
                                        <option value='ǧ��'>ǧ��</option>
                                        <option value='����'>����</option>

                                    </select>
                                </td>


                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">����ԧĤ</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">ԧĤ���</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">ԧĤ���</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">����</font></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <font color="#FFFFFF">��λ</font></td>
                                <td height="12" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="5">
                                    <font color="#FFFFFF">����Ҫ��</font></td>


                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name='name3' style='width: 55; height: 16'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option selected="selected"><%=rs("name3")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from cailiao where fenlei='��Ĥ' order by pinming",conn,2,2
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
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="nianbang3" size="1" style='width: 54; height: 18'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option><%=rs("nianbang3")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from guige order by guige",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
                                        %>
                                        <option <% if isedit then
			if rs3("guige")=rs("nianbang3") then
			%>
                                            selected="selected" <%end if%><%end if%>><%=rs3("guige")%></option>
                                        <%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="cunshu3" size="1" style='width: 55; height: 16'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option><%=rs("cunshu3")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%
			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select * from houdou order by houdou",conn,1,1
		   if not rs4.eof then
		   do while not rs4.eof
                                        %>
                                        <option <% if isedit then
			if rs4("houdou")=rs("cunshu3") then
			%>
                                            selected="selected" <%end if%><%end if%>><%=rs4("houdou")%></option>
                                        <%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="nianti3" type="text" style="ime-mode: disabled" size="18" value='<% if isedit then
		                                                         response.write rs("nianti3")
		                                                          else response.write 0

																  end if %>'></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <select name='unit3' size="1" style='width: 64; height: 16'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option selected="selected"><%=rs("unit3")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <option value='��'>��</option>
                                        <option value='��'>��</option>
                                        <option value='ǧ��'>ǧ��</option>
                                        <option value='����'>����</option>

                                    </select>
                                </td>
                                <td height="36" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="5" rowspan="3">
                                    <textarea name="chati4" cols="35" rows="3"><% if isedit then
		                                                         response.write rs("chati4")
																  end if %></textarea></td>


                            </tr>
                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">������Ĥ</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">��Ĥ���</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">��Ĥ���</font></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">����</font></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <font color="#FFFFFF">��λ</font></td>


                            </tr>

                            <tr>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name='name4' style='width: 55; height: 16'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option selected="selected"><%=rs("name4")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select * from cailiao where fenlei='��Ĥ' order by pinming",conn,2,2
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
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="nianbang4" size="1" style='width: 54; height: 18'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option><%=rs("nianbang4")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%
			
			set rs3=server.createobject("adodb.recordset")
		   rs3.open "select * from guige order by guige",conn,1,1
		   if not rs3.eof then
		   do while not rs3.eof
                                        %>
                                        <option <% if isedit then
			if rs3("guige")=rs("nianbang4") then
			%>
                                            selected="selected" <%end if%><%end if%>><%=rs3("guige")%></option>
                                        <%
			rs3.movenext
			loop
			end if
			rs3.close
			set rs3=nothing
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <select name="cunshu4" size="1" style='width: 55; height: 16'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option><%=rs("cunshu4")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <%
			
			set rs4=server.createobject("adodb.recordset")
		   rs4.open "select * from houdou order by houdou",conn,1,1
		   if not rs4.eof then
		   do while not rs4.eof
                                        %>
                                        <option <% if isedit then
			if rs4("houdou")=rs("cunshu4") then
			%>
                                            selected="selected" <%end if%><%end if%>><%=rs4("houdou")%></option>
                                        <%
			rs4.movenext
			loop
			end if
			rs4.close
			set rs4=nothing
                                        %>
                                    </select></td>
                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <input name="nianti4" type="text" style="ime-mode: disabled" size="18" value='<% if isedit then
		                                                         response.write rs("nianti4")
		                                                          else response.write 0

																  end if %>'></td>


                                <td height="18" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <select name='unit4' size="1" style='width: 64; height: 16'>
                                        <% if isedit then%><option selected value='0'>��ѡ��</option>
                                        <option selected="selected"><%=rs("unit4")%></option>
                                        <%else%>
                                        <option selected value='0'>��ѡ��</option>
                                        <%end if%>
                                        <option value='��'>��</option>
                                        <option value='��'>��</option>
                                        <option value='ǧ��'>ǧ��</option>
                                        <option value='����'>����</option>

                                    </select>
                                </td>


                            </tr>
                            <tr>
                                <td height="40" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'">
                                    <font color="#FFFFFF">��ƷҪ��</font></td>
                                <td height="40" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="3">
                                    <textarea name="chati2" cols="42" rows="4"><% if isedit then
		                                                         response.write rs("chati2")
																  end if %></textarea></td>


                                <td height="40" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="2">
                                    <font color="#FFFFFF">��װҪ��</font></td>
                                <td height="40" align="center" class="but11" onmousedown="this.className='tddown'" onmouseup="this.className='but'" onmouseout="this.className='but11'" colspan="5">
                                    <textarea name="chati3" cols="35" rows="3"><% if isedit then
		                                                         response.write rs("chati3")
																  end if %></textarea></td>


                            </tr>
                            <tr>
                                <td colspan="11" align="center">
                                    <input type="submit" value="<% if isedit then
		                                                         response.write "�༭"
																 else
																 response.write "���"
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
    <table width="800" align="center" border="0" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
        <tr>
            <td>

                <table width="800" border="1" cellpadding="0" cellspacing="0" bordercolor="#0055E6" align="center">

                    <tr>

                        <td width="800" align="center" valign="top">
                            <table width="800" border="0" cellpadding="0" cellspacing="0">
                                <tr>
                                    <td height="26" background="images/background.gif">
                                        <b><font color="#ffffff">&nbsp; �� �� �� ��</font></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                                        <%if oskey="chejian" then%>
                                        <input type="button" name="close" value="�ر�" onclick="window.close();" /><%end if%>
                                    </td>
                                </tr>
                            </table>
                            <form action="dindang.asp" method="post" name="selform">
                                <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
                                    <tr>

                                        <td width="100%" align="center" valign="top">

                                            <table width="100%" border="1" cellpadding="0" cellspacing="1" bordercolor="#EFEADD">
                                                <tr class="but">
                                                    <td width="9%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">�������</td>
                                                    <td width="6%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">�� ��</td>
                                                    <td width="34%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">Ʒ ��</td>
                                                    <td width="7%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">�� ��</td>
                                                    <td width="5%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">��Ĥ</td>
                                                    <td width="20%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">����</td>
                                                    <td width="10%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">�ͻ�</td>
                                                    <td width="5%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">
                                                        <span lang="zh-cn">����</span></td>
                                                    <td width="4%" height="18" align="center" class="but01" onmousedown="this.classname='tddown'" onmouseup="this.classname='but'" onmouseout="this.classname='but01'">����</td>
                                                </tr>
                                                <%
PageShowSize = 10            'ÿҳ��ʾ���ٸ�ҳ
MyPageSize   = 50          'ÿҳ��ʾ������
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
bkeyword=request("bkeyword")
rs.Source="select top 7000 * from dindang where kehu like '%"&akeyword&"%' and pinming like '%"&keyword&"%' and nianbang like '%"&bkeyword&"%' order by -id"
else
rs.Source="SELECT * FROM dindang ORDER BY id DESC OFFSET ("&MyPage&"-1)*"&MyPageSize&" ROWS FETCH NEXT "&MyPageSize&" ROWS ONLY"
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
                                                <tr onmouseover="this.style.backgroundColor='#73A475';" onmouseout="this.style.backgroundColor='';">
                                                    <td height="18" align="center"><a href="dindang.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("uptime")%></a></td>
                                                    <td height="18" align="center"><%=rs("fenlei")%><%if rs("fenlei")="" then%>&nbsp;<%end if%></td>
                                                    <td height="18" align="center">&nbsp;<a href="dindang.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("pinming")%></a></td>
                                                    <td height="18" align="center"><%=rs("chang")%>*<%=rs("kuang")%></td>
                                                    <td height="18" align="center"><%if rs("nianbang")=0 then%>����<%else%><%=rs("nianbang")%><%end if%></td>
                                                    <td height="18" align="center">
                                                        <p align="left">
                                                            &nbsp;<%=rs("name1")%>��<%if rs("name2")="0" then%><%else%><%=rs("name2")%>��<%end if%><%if rs("name3")="0" then%><%else%><%=rs("name3")%>��<%end if%><%if rs("name4")="0" then%><%else%><%=rs("name4")%><%end if%>
                                                    </td>
                                                    <td height="18" align="center"><a href="zhidai.Asp?id=<%=rs("ID")%>&action=edit"><%=rs("kehu")%><%if rs("kehu")="" then%>&nbsp;<%end if%></a></td>
                                                    <td height="18" align="center"><a href="dindang_in.asp?id=<%=rs("id")%>&action=edit">�� ��</a></td>
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
                                                    <td align="right" colspan="9" height="22">�� <%=total%> ������ǰ�� <%=Mypage%>/<%=Maxpages%> 
            ҳ��ÿҳ <%=MyPageSize%> �� 
            <%
			if search<>"" then
			url="dindang.asp?keyword="&keyword&"&akeyword="&akeyword&"&bkeyword="&bkeyword&"&search=search&"
			else
            url="dindang.asp?"
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
            %>&nbsp;&nbsp;&nbsp;&nbsp;<%if oskey="supper" or oskey="admin" then%>
                                                        <input type="checkbox" name="checkbox" value="checkbox" onclick="javascript:SelectAll()">
                                                        ѡ��/��ѡ
              <input onclick="{if(confirm('�˲�����ɾ������Ϣ��\n\nȷ��Ҫִ�д��������')){this.document.selform.submit();return true;}return false;}" type="submit" value="ɾ��" name="action2">
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
            </td>
        </tr>
    </table>
    <br>
    <table width="800" align="center" border="1" cellspacing="0" cellpadding="0" bordercolor="#0055E6">
        <tr>
            <td>


                <table width="800" border="0" cellpadding="0" cellspacing="0">
                    <tr>
                        <td height="26" background="images/background.gif">
                            <b><font color="#ffffff">�������</font></b>

                        </td>
                    </tr>
                </table>

                <table width="100%" border="0" cellpadding="0" cellspacing="0" bordercolor="#D4D0C8">

                    <form name="search" method="post" action="dindang.asp">
                        <tr>
                            <input name="search" type="hidden" value="search">
                            <td align="center" width="29%">��Ʒ�����ң� 
				<input name="keyword" type="text" size="23"></td>

                            <td align="center" width="9%">
                                <input type="submit" value="�� ��">
                            </td>
                            <td align="center" width="26%">���ͻ����ң� 
					<select name='akeyword' style='width: 130'>
                        <option selected value="">��ѡ��ͻ�</option>
                        <%			
			set rs2=server.createobject("adodb.recordset")
		   rs2.open "select distinct(kehu) from dindang order by kehu",conn,2,2
		   if not rs2.eof then
		   do while not rs2.eof
                        %>
                        <option value="<%=rs2("kehu")%>"><%=rs2("kehu")%></option>
                        <%
			rs2.movenext
			loop
			end if
			rs2.close
			set rs2=nothing
                        %>
                                </select></td>

                            <td align="center" width="7%">
                                <input type="submit" value="�� ��">
                            </td>

                            <td align="center" width="20%">����ȣ�
			      <input name="bkeyword" type="text" size="11"></td>

                            <td align="center" width="10%">
                                <input type="submit" value="�� ��">
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
