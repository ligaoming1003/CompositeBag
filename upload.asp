<!--#include file="upload_class.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="zh-cn" lang="zh-cn">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title>�ϴ�</title>
<style type="text/css">
TABLE {border:1px green solid;margin-top:5px;}
TD{border-bottom:1px #dddddd solid;height:20px;padding:3px 0 0 5px;}
.head{background-color:#eeeeee;}
</style>
</head>
<body style="font-size:12px">
��ѡ��һ���ļ������ϴ�:<br />
<form name="upload" method="post" action="upload.asp?act=upload" enctype="multipart/form-data">
ѡ���ļ�:<input class="iFile" id="file1" type="file" name="file1" size="30" />
<input class="iButton" type="submit" value="�ϴ�" />
</form>

<%
if request.querystring("act")="upload" then
 Dim Upload,successful,str
 str=""
'===============================================================================
 set Upload=new AnUpLoad 		'������ʵ��
 Upload.SingleSize=200*1024            			'���õ����ļ�����ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
 Upload.MaxSize=200*1024            			'��������ϴ�����,���ֽڼƣ�Ĭ��Ϊ������
 Upload.Exe="rar|jpg|bmp|gif|png"          				'���úϷ���չ��,��|�ָ�,���Դ�Сд
 Upload.GetData()						'��ȡ����������,������ñ�����
'===============================================================================
 if Upload.Err>0 then						'�жϴ����,���myupload.Err<=0��ʾ����
 	response.write Upload.Description 			'������ִ���,��ȡ��������
 else
 	if Upload.forms("file1")<>"" then 			'�����ж���file1�Ƿ�ѡ�����ļ�
    		path=server.mappath("files") 			'�ļ�����·��(������files�ļ���)
    		set tempCls=Upload.files("file1") 
    		successful=tempCls.SaveToFile(path,0)		'��ʱ��+�������Ϊ�ļ�������
    		'successful=tempCls.SaveToFile(path,1)		'�������ԭ�ļ�������,��ʹ�ñ���
		    if successful then
    			str="files/" & tempCls.FileName
		    end if
    		set tempCls=nothing
 	end if
 end if
 set Upload=nothing                   '������ʵ��
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

