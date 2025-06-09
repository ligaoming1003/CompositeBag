<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="checkuser.asp"-->
<html>
<head>
<title>∷嘉美管理系统∷</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="pragma" content="no-cache">
<link href="qq/main.css" rel="stylesheet" type="text/css">
<style type="text/css">
    body {
        font-size: 9pt;
        line-height: 15pt;
        font-style: normal;
        font-family: "宋体";
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
        font-family: "宋体";
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
        /* 悬停样式 */
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
    <!-- 基础管理菜单 -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(1);hiddensubmenu(2);hiddensubmenu(3);hiddensubmenu(4);hiddensubmenu(5);hiddensubmenu(6);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;基础管理&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu1">
        <%if oskey="supper" then%>
        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiaocangku.asp" target="main">仓库名称</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="unit.asp" target="main">计量单位</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="department.asp" target="main">部门信息</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

      <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="yinhang.asp" target="main">款项来源</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chejian.asp" target="main">车间信息</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>
        
        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="kucen.asp" target="main">库存录入</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="zanfu.asp" target="main">上年应收</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="user.asp" target="main">权限分配</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="admin_data_backup.asp" target="main">数据备份</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_in_del.asp" target="main">入库清零</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_out_del.asp" target="main">出库清零</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="caiyin_del.asp" target="main">彩印清零</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
         <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fuhe_del.asp" target="main">复合清零</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="matur_del.asp" target="main">制袋清零</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="jsm_out_del.asp" target="main">销售清零</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="daozhang_del.asp" target="main">到账清零</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="js_one_del.asp" target="main">结算清零</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="js_one1.asp" target="main">结算返回</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 基础管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="middle" height="20">
                        <td class="tdlineyb" valign="middle" align="center" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="gy_one1.asp" target="main">供应返回</a>
                        </td>
                    </tr>
                </tbody>
            </table>            
        </div>

        <!-- 更多子菜单项... -->
        <%end if%>
    </div>
    
    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>
    
    <!-- 订单管理菜单 -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(2);hiddensubmenu(1);hiddensubmenu(3);hiddensubmenu(4);hiddensubmenu(5);hiddensubmenu(6);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;订单管理&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu2">
        <!-- 订单管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="kehu.asp" target="main">新增客户</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 订单管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="jijia.asp" target="main">计价系统</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 订单管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="dangang.asp" target="main">档案录入</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 订单管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="dindang.asp" target="main">订单录入</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 订单管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="dindang1.asp" target="main">材料安排</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

         <!-- 订单管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="dindang3.asp" target="main">生产进度</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 订单管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_dindang.asp" target="main">拨后材料</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 订单管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_compiles.asp" target="main">拨后查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 更多子菜单项... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- 材料管理菜单 -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(3);hiddensubmenu(1);hiddensubmenu(2);hiddensubmenu(4);hiddensubmenu(5);hiddensubmenu(6);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;材料管理&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu3">
        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fenlei.asp" target="main">材料类别</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao.asp" target="main">材料名称</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="guige.asp" target="main">规格宽度</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="houdou.asp" target="main">规格厚度</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="gonghuodanwei.asp" target="main">供方信息</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_in_store.asp" target="main">材料入库</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fuliao_in_store.asp" target="main">辅料入库</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_in_compiles.asp" target="main">入库查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_out_store.asp" target="main">材料出库</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_out_compiles.asp" target="main">出库查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="tiaozhen.asp" target="main">材料调剂</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_out_ss.asp" target="main">其它出库</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_store.asp" target="main">材料库存</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 材料管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="cailiao_store_compiles.asp" target="main">库存查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
        <!-- 更多子菜单项... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- 车间管理菜单 -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(4);hiddensubmenu(1);hiddensubmenu(3);hiddensubmenu(2);hiddensubmenu(5);hiddensubmenu(6);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;车间管理&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu4">
        <!-- 车间管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="caiyin.asp" target="main">彩印出库</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 车间管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="caiyin_in.asp" target="main">出库明细</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 车间管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="caiyin_compiles.asp" target="main">彩印查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 车间管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fuhe.asp" target="main">复合出库</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 车间管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fuhe_in.asp" target="main">出库明细</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 车间管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fuhe_compiles.asp" target="main">复合查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>


        <!-- 车间管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="mxfuhe.asp" target="main">复合退库</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 车间管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="shigu.asp" target="main">责任事故</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 更多子菜单项... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- 成品管理菜单 -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(5);hiddensubmenu(1);hiddensubmenu(3);hiddensubmenu(4);hiddensubmenu(2);hiddensubmenu(6);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;成品管理&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu5">
        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="fenlei1.asp" target="main">成品类别</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_qita.asp" target="main">成品结存</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="matur1.asp" target="main">品名调整</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="matur.asp" target="main">成品入库</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_in_store.asp" target="main">入库明细</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_in_compiles.asp" target="main">入库查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_out_store.asp" target="main">销售出库</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="jiagong.asp" target="main">加工销售</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="qida_out_store.asp" target="main">其它销售</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_out_compiles.asp" target="main">销售查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_store.asp" target="main">成品库存</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 成品管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="chengping_store_compiles.asp" target="main">库存查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>


        <!-- 更多子菜单项... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- 财务管理菜单 -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(6);hiddensubmenu(1);hiddensubmenu(3);hiddensubmenu(4);hiddensubmenu(5);hiddensubmenu(2);hiddensubmenu(7)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;财务管理&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu6">
        <!-- 财务管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="daozhang.asp" target="main">货款到账</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 财务管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="daozhang_compiles.asp" target="main">到账查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 财务管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="gy_fk.asp" target="main">供应付款</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 财务管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="jsm_out_store.asp" target="main">发货明细</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 财务管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="js_daozhang.asp" target="main">收款明细</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>


        <!-- 更多子菜单项... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- 结算管理菜单 -->
    <div class="menu_title" 
         onmouseover="this.className='menu_title2'" 
         onmouseout="this.className='menu_title'" 
         onclick="showsubmenu(7);hiddensubmenu(1);hiddensubmenu(3);hiddensubmenu(4);hiddensubmenu(5);hiddensubmenu(6);hiddensubmenu(2)">
        <b><img src="images/d-1.gif" width="8" height="12" alt="">&nbsp;结算管理&nbsp;</b>
    </div>
    
    <div style="display:none" id="submenu7">
        <!-- 结算管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="js_one.asp" target="main">销售结算</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 结算管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="js_compiles.asp" target="main">销售查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 结算管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="guhu.asp" target="main">过户领取</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 结算管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="gy_one.asp" target="main">供方结算</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>

        <!-- 结算管理子菜单项 -->
        <div class="submenu_container" name="div103">
            <table width="100%" border="0">
                <tbody>
                    <tr valign="center" height="20">
                        <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
                    </tr>
                    <tr>
                        <td class="tdback" onmouseover="this.className='tdhover';" onmouseout="this.className='tdback';" nowrap align="middle" height="20">
                            <a href="gy_compiles.asp" target="main">供方查询</a>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>


        <!-- 更多子菜单项... -->
    </div>

    <table width="100%" border="0">
        <tbody>
            <tr>
                <td class="tdlineyb" valign="center" align="middle" width="100%" height="1"></td>
            </tr>
        </tbody>
    </table>

    <!-- 更多主菜单项... -->
</div>
</body>
</html>