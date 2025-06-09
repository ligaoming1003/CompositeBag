minimizebar="image/skin/min.gif";   //窗口右上角最小化“按钮”的图片
minimizebar2="image/skin/min.gif"; //鼠标悬停时最小化“按钮”的图片
closebar="image/skin/close.gif";         //窗口右上角关闭“按钮”的图片
closebar2="image/skin/close.gif";       //鼠标悬停时关闭“按钮”的图片
icon="image/skin/close.gif";              //窗口左上角的小图标

function noBorderWin(fileName,w,h,titleBg,moveBg,titleColor,titleWord,scr)  //定义一个弹出无边窗口的函数，能数意义见下面“参数说明”，实际使用见最后的实例。
/*
------------------参数说明-------------------
fileName   ：无边窗口中显示的文件。
w　　　　  ：窗口的宽度。
h　　　　  ：窗口的高度。
titleBg    ：窗口“标题栏”的背景色以及窗口边框颜色。
moveBg     ：窗口拖动时“标题栏”的背景色以及窗口边框颜色。
titleColor ：窗口“标题栏”文字的颜色。
titleWord  ：窗口“标题栏”的文字。
scr        ：是否出现滚动条。取值yes/no或者1/0。
--------------------------------------------
*/
{
  var contents="<html>"+
               "<head>"+
	    	   "<title>"+titleWord+"</title>"+
			   "<link href='nobody.css' rel='stylesheet' type='text/css'>"+
			   "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=gb2312\">"+
			   "<link href='nobody.css' rel='stylesheet' type='text/css'>"+
			   "<object id=hhctrl type='application/x-oleobject' classid='clsid:adb880a6-d8ff-11cf-9377-00aa003b7a11'><param name='Command' value='minimize'></object>"+
			   "</head>"+
               "<body topmargin=0 leftmargin=0 scroll=no onselectstart='return false' ondragstart='return false'  oncontextmenu='return false' onkeydown='if(event.ctrlKey)return false'>"+
			   "  <table height=100% width=100% cellpadding=0 cellspacing=0 bgcolor="+titleBg+" id=mainTab class='body1'>"+
			   "    <tr bordercolor="+titleBg+" height=18 style=cursor:default; onmousedown='x=event.x;y=event.y;setCapture();mainTab.bgColor=\""+moveBg+"\";' onmouseup='releaseCapture();mainTab.bgColor=\""+titleBg+"\";' onmousemove='if(event.button==1)self.moveTo(screenLeft+event.x-x,screenTop+event.y-y);'>"+
			   "      <td width=18 align=center background='image/skin/title.gif'><img height=12 width=12 border=0 src="+icon+"></td>"+
			   "      <td width="+w+" background='image/skin/title.gif'>"+titleWord+"</td>"+
			   "      <td width=14 background='image/skin/title.gif'><img border=0 alt=最小化 src="+minimizebar+" onmousedown=hhctrl.Click(); onmouseover=this.src='"+minimizebar2+"' onmouseout=this.src='"+minimizebar+"'></td>"+
			   "      <td width=13 background='image/skin/title.gif'><img border=0 alt=关闭 src="+closebar+" onmousedown=nbw_v6_iframe.exitask(); onmouseover=this.src='"+closebar2+"' onmouseout=this.src='"+closebar+"'></td>"+
			   "    </tr>"+
			   "    <tr height=*>"+
			   "      <td colspan=4 align=center>"+
			   "        <iframe name=nbw_v6_iframe src="+fileName+" scrolling="+scr+" width=100% height=100% frameborder=0></iframe>"+
			   "      </td>"+
			   "    </tr>"+
			   "  </table>"+
			   "</body>"+
			   "</html>";

  pop=window.open("","_blank","fullscreen=yes");
  pop.resizeTo(w,h);
  pop.moveTo((screen.width-w)*8/9,(screen.height-h)/2);
  pop.document.writeln(contents);
}