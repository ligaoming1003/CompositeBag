<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>�˼�������ϵͳ��</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="pragma" content="no-cache">
<link href="qq/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
    body {
        font-size: 9pt;
        line-height: 15pt;
        font-style: normal;
        font-family: "����";
        text: #000000;
        leftmargin: 0;
        rightmargin: 0;
        background: url(images/xiao-04.gif);
        topmargin: 0;
        overflow-y: auto;
        overflow-x: auto;
    }
    
    td {
        font-size: 9pt;
        line-height: 15pt;
        font-style: normal;
        font-family: "����";
    }
    
    a:hover {
        color: red;
        text-decoration: underline;
    }
    
    a:link {
        color: black;
        text-decoration: none;
    }
    
    a:visited {
        color: black;
        text-decoration: none;
    }
    
    .divback {
        border: 1px solid #9c9c9c;
        border-bottom: 0;
        background-color: #efeadd;
    }
    
    .divhover {
        background-color: #caccc9;
    }
    
    .tdback {
        border: 1px solid #efeadd;
        padding-left: 3px;
        line-height: 14px;
        background-color: #efeadd;
    }
    
    .tdhover {
        border: 1px solid #9c9c9c;
        padding-left: 3px;
        line-height: 14px;
        background-color: #caccc9;
    }
    
    .tdlineyb {
        background-color: #dddddd;
    }
    
    .menu_title {
        padding: 6px 0 3px 10px;
        left: 1px;
        width: 100%;
        cursor: pointer;
    }
    
    .menu_title2 {
        /* ��ͣ��ʽ */
    }
    
    .submenu_container {
        padding-right: 0;
        border-top: 0;
        overflow-y: auto;
        padding-left: 0;
        overflow-x: auto;
        border-left: 1px solid #9c9c9c;
        width: 100%;
        padding-top: 0;
        border-bottom: 0;
        background-color: #efeadd;
        center: 0;
    }
</style>



<script type="text/javascript">
    function showsubmenu(sid) {
        var whichel = document.getElementById("submenu" + sid);
        if (whichel.style.display == "none") {
            whichel.style.display = "";
        } else {
            whichel.style.display = "none";
        }
    }

    function hiddensubmenu(sid) {
        var whichel = document.getElementById("submenu" + sid);
        whichel.style.display = "none";
    }
</script>







</head>

<body>
<div style="overflow-y: auto; overflow-x: auto;">
    <!-- ��������˵� -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(1);hiddensubmenu(2);hiddensubmenu(3);hiddensubmenu(4);hiddensubmenu(5);hiddensubmenu(6);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;��������&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu1">
        <%if oskey="supper" then%>
        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiaocangku.asp" target="main">�ֿ�����</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="unit.asp" target="main">������λ</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="department.asp" target="main">������Ϣ</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

      <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="yinhang.asp" target="main">������Դ</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chejian.asp" target="main">������Ϣ</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>
        
        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="kucen.asp" target="main">���¼��</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="zanfu.asp" target="main">����Ӧ��</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="user.asp" target="main">Ȩ�޷���</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="admin_data_backup.asp" target="main">���ݱ���</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_in_del.asp" target="main">�������</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_out_del.asp" target="main">��������</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="caiyin_del.asp" target="main">��ӡ����</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
         <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fuhe_del.asp" target="main">��������</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="matur_del.asp" target="main">�ƴ�����</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="jsm_out_del.asp" target="main">��������</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="daozhang_del.asp" target="main">��������</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="js_one_del.asp" target="main">��������</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="js_one1.asp" target="main">���㷵��</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="gy_one1.asp" target="main">��Ӧ����</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- �����Ӳ˵���... -->
        <%end if%>
    </div>
    
    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>
    
    <!-- ��������˵� -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(2);hiddensubmenu(1);hiddensubmenu(3);hiddensubmenu(4);hiddensubmenu(5);hiddensubmenu(6);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;��������&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu2">
        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="kehu.asp" target="main">�����ͻ�</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="jijia.asp" target="main">�Ƽ�ϵͳ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="dangang.asp" target="main">����¼��</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="dindang.asp" target="main">����¼��</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="dindang1.asp" target="main">���ϰ���</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

         <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="dindang3.asp" target="main">��������</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_dindang.asp" target="main">�������</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_compiles.asp" target="main">�����ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- �����Ӳ˵���... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- ���Ϲ���˵� -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(3);hiddensubmenu(1);hiddensubmenu(2);hiddensubmenu(4);hiddensubmenu(5);hiddensubmenu(6);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;���Ϲ���&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu3">
        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fenlei.asp" target="main">�������</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao.asp" target="main">��������</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="guige.asp" target="main">�����</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="houdou.asp" target="main">�����</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="gonghuodanwei.asp" target="main">������Ϣ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_in_store.asp" target="main">�������</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fuliao_in_store.asp" target="main">�������</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_in_compiles.asp" target="main">����ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_out_store.asp" target="main">���ϳ���</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_out_compiles.asp" target="main">�����ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="tiaozhen.asp" target="main">���ϵ���</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_out_ss.asp" target="main">��������</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_store.asp" target="main">���Ͽ��</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ���Ϲ����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_store_compiles.asp" target="main">����ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        <!-- �����Ӳ˵���... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- �������˵� -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(4);hiddensubmenu(1);hiddensubmenu(3);hiddensubmenu(2);hiddensubmenu(5);hiddensubmenu(6);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;�������&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu4">
        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="caiyin.asp" target="main">��ӡ����</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="caiyin_in.asp" target="main">������ϸ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="caiyin_compiles.asp" target="main">��ӡ��ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fuhe.asp" target="main">���ϳ���</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fuhe_in.asp" target="main">������ϸ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fuhe_compiles.asp" target="main">���ϲ�ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>


        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="mxfuhe.asp" target="main">�����˿�</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="shigu.asp" target="main">�����¹�</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- �����Ӳ˵���... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- ��Ʒ����˵� -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(5);hiddensubmenu(1);hiddensubmenu(3);hiddensubmenu(4);hiddensubmenu(2);hiddensubmenu(6);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;��Ʒ����&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu5">
        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fenlei1.asp" target="main">��Ʒ���</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_qita.asp" target="main">��Ʒ���</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="matur1.asp" target="main">Ʒ������</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="matur.asp" target="main">��Ʒ���</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_in_store.asp" target="main">�����ϸ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_in_compiles.asp" target="main">����ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_out_store.asp" target="main">���۳���</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="jiagong.asp" target="main">�ӹ�����</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="qida_out_store.asp" target="main">��������</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_out_compiles.asp" target="main">���۲�ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_store.asp" target="main">��Ʒ���</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��Ʒ�����Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_store_compiles.asp" target="main">����ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>


        <!-- �����Ӳ˵���... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- �������˵� -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(6);hiddensubmenu(1);hiddensubmenu(3);hiddensubmenu(4);hiddensubmenu(5);hiddensubmenu(2);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;�������&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu6">
        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="daozhang.asp" target="main">�����</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="daozhang_compiles.asp" target="main">���˲�ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="gy_fk.asp" target="main">��Ӧ����</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="jsm_out_store.asp" target="main">������ϸ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="js_daozhang.asp" target="main">�տ���ϸ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>


        <!-- �����Ӳ˵���... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- �������˵� -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(7);hiddensubmenu(1);hiddensubmenu(3);hiddensubmenu(4);hiddensubmenu(5);hiddensubmenu(6);hiddensubmenu(2)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;�������&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu7">
        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="js_one.asp" target="main">���۽���</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="js_compiles.asp" target="main">���۲�ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="guhu.asp" target="main">������ȡ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="gy_one.asp" target="main">��������</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- ��������Ӳ˵��� -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="gy_compiles.asp" target="main">������ѯ</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>


        <!-- �����Ӳ˵���... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- �������˵���... -->
</div>
</body>
</html>