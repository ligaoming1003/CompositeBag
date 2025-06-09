<!--#include file="upload_class.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn" lang="zh-cn">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>上传</title>
<style type="text/css">
TABLE {border:1px green solid;margin-top:5px;}
TD{border-bottom:1px #dddddd solid;height:20px;padding:3px 0 0 5px;}
.head{background-color:#eeeeee;}
</style>
</head>
<body style="font-size:12px">
请选择一个文件进行上传:<br />
<form name="upload" method="post" action="upload.asp?act=upload" enctype="multipart/form-data">
选择文件:<input class="iFile" id="file1" type="file" name="file1" size="30" />
<input class="iButton" type="submit" value="上传" />
</form>

<%
if request.querystring("act")="upload" then
 Dim Upload,successful,str
 str=""
'===============================================================================
 set Upload=new AnUpLoad 		'创建类实例
 Upload.SingleSize=200*1024            			'设置单个文件最大上传限制,按字节计；默认为不限制
 Upload.MaxSize=200*1024            			'设置最大上传限制,按字节计；默认为不限制
 Upload.Exe="rar|jpg|bmp|gif|png"          				'设置合法扩展名,以|分割,忽略大小写
 Upload.GetData()						'获取并保存数据,必须调用本方法
'===============================================================================
 if Upload.Err>0 then						'判断错误号,如果myupload.Err<=0表示正常
 	response.write Upload.Description 			'如果出现错误,获取错误描述
 else
 	if Upload.forms("file1")<>"" then 			'这里判断你file1是否选择了文件
    		path=server.mappath("files") 			'文件保存路径(这里是files文件夹)
    		set tempCls=Upload.files("file1") 
    		successful=tempCls.SaveToFile(path,0)		'以时间+随机数字为文件名保存
    		'successful=tempCls.SaveToFile(path,1)		'如果想以原文件名保存,请使用本句
		    if successful then
    			str="files/" & tempCls.FileName
		    end if
    		set tempCls=nothing
 	end if
 end if
 set Upload=nothing                   '销毁类实例
 if str<>"" then
%>
<script type="text/javascript">
window.opener.document.all.file1.value='<%=str%>';
window.opener=null;
window.close();
</script>
<%
 end if
end if
%>
</body>
</html>

