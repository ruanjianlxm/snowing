
<!--#include file="../inc/upload.asp"-->
<%
Server.ScriptTimeOut=5000
%>
<html>

<head>
<meta http-equiv="content-type" content="text/html;charset=gb2312">
<title>上传文件</title>
<style>
*{margin:0px;padding:0px;}
body{
font-size:12px;
}
.border{
font-size:12px;
border:#000 solid 1px;
}
</style>
<link rel="stylesheet" rev="stylesheet" href="images/css.css" type="text/css" media="all" />

</head>

<body leftmargin=0 topmargin=0>

<%
Dim go:go=Request.QueryString("go")
If Request.QueryString("action")="upload" Then 

Set upload=new my_upload
Dim filepath
	filepath=trim(upload.form("filepath"))
For each formName in upload.File
	set file=upload.File(formName)
	Dim o,txt,FileExt:FileExt=file.FileExt
	
	txt="gif|jpg|bmp|png|swf|flv"
	
If InStr(txt,LCase(FileExt))=0 Then 
	response.write "<script>alert('您上传的格式错误！');location.href='upload.asp';</script>"
	response.end
End if

	
	Randomize
	ranNum=int(90000*rnd)+10000
	filename=filepath&year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)&ranNum&"."&FileExt
	'开始保存文件
	If file.FileSize>0 Then
	   file.SaveToFile Server.mappath(filename) 
	Response.Write("<script type='text/javascript'>")
	If go="down" Then 
		Response.Write("parent.document.getElementById('bUrl').value='"&Replace(filename,"../","")&"';")
		Response.Write("parent.document.getElementById('uploadurl').style.display='none';")
		Response.Write("parent.document.getElementById('boxurl').checked=false;")
	
	Else
		Response.Write("parent.document.getElementById('bPic').value='"&Replace(filename,"../","")&"';")
		Response.Write("parent.document.getElementById('upload').style.display='none';")
		Response.Write("parent.document.getElementById('box').checked=false;")
	
	End if
	Response.Write("</script>")
	End if
	Set file=Nothing 
next
Set upload=Nothing 
End if
If go="down" Then%><form name="form1" method="post" action="?go=down&action=upload" enctype="multipart/form-data">
	<%Else%></form>
<form name="form1" method="post" action="?action=upload" enctype="multipart/form-data">
	<%End if%><input type="hidden" name="filepath" value="../Upload/"><input size="49" class="btn" type="file" name="file1" />
	<input type="submit" name="Submit" class="button" value="开始上传" /></form>

</body>

</html>
