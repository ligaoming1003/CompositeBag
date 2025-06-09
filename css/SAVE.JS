if ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >= 4))
{
	var url="http://meet.lc.gov/";
	var title="【中共柳城县委、县人民政府文件视频会议系统】";

	document.write('<A HREF="javascript:window.ext');
	document.write('ernal.AddFavorite(url,title);" ');
	document.write('onMouseOver=" window.status=');
	document.write("'收藏本站'; return true ");
	document.write('"onMouseOut=" window.status=');
	document.write("' '; return true ");
	document.write('">收 藏 本 站</a>');}
else
{
	var msg = "收藏本站";
	if(navigator.appName == "Netscape") msg += "(CTRL-D)";
	document.write(msg);
}

