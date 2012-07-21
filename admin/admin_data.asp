<%@ Language=VBScript %> <%
Response.Buffer = true
'禁用缓存
Response.Expires = -10000
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private"
Response.CacheControl = "no-cache"

if session("admin")="" then
    response.Redirect("index.asp?go=body")
end if
%>
<html>
<head>
<title>数据库管理页面</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 6.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
dim action
action=trim(request("action"))

dim dbpath,bkfolder,bkdbname,fso,fso1

select case action
case "CompressData"		'压缩数据
	
		dim tmprs
		dim allarticle
		dim Maxid
		dim topic,dateandtime,body
		call CompressData()


case "BackupData"		'备份数据
		if request("act")="Backup" then
			call updata()
		else
			call BackupData()
		end if

case "RestoreData"		'恢复数据
	dim backpath
		if request("act")="Restore" then
			Dbpath=request.form("Dbpath")
			backpath=request.form("backpath")
			if dbpath="" then
			response.write "请输入您要恢复成的数据库全名"	
			else
			Dbpath=server.mappath(Dbpath)
			end if
			backpath=server.mappath(backpath)
		
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then  					
			fso.copyfile Dbpath,Backpath
			response.write "成功恢复数据！"
			else
			response.write "备份目录下并无您的备份文件！"	
			end if
		else
		call RestoreData()
		end if
case else
	response.Write "请选取相应的操作"

end select

response.write"</body></html>"

'====================恢复数据库=========================
sub RestoreData()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<tr>
	<th height=25 >
   					&nbsp;&nbsp;<B>恢复整站数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
				<form method="post" action="ADMIN_data.asp?action=RestoreData&act=Restore">
  				
  				<tr>
  					<td height=100 class="txlrow">
  						&nbsp;&nbsp;备份数据库路径(相对)：<input type=text size=30 name=DBpath value="../data/backup/dbbak.mdb">
  						&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;目标数据库路径(相对)：<input type=text size=30 name=backpath value="../data/#$%~~data.mdb">
  						<BR>&nbsp;&nbsp;填写您当前使用的数据库路径，如不想覆盖当前文件，可自行命名（注意路径是否正确），然后修改conn.asp文件，如果目标文件名和当前使用数据库名一致的话，不需修改conn.asp文件<BR>
						&nbsp;&nbsp;<input type=submit value="恢复数据"> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认备份数据库文件为data\backup\db.mdb，请按照您的备份文件自行修改。<br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

'====================备份数据库=========================
sub BackupData()
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;<B>备份整站数据</B>( 需要FSO支持，FSO相关帮助请看微软网站 )
  					</th>
  				</tr>
  				<form method="post" action="ADMIN_data.asp?action=BackupData&act=Backup">
  				<tr>
  					<td height=100 class="txlrow">
  						&nbsp;&nbsp;
						当前数据库路径(相对路径)：<input type=text size=30 name=DBpath value="../data/#$%~~data.mdb">
						<BR>
						&nbsp;&nbsp;
						备份数据库目录(相对路径)：<input type=text size=30 name=bkfolder value="../data/backup">
						&nbsp;如目录不存在，程序将自动创建<BR>
						&nbsp;&nbsp;
						备份数据库名称(填写名称)：<input type=text size=30 name=bkDBname value=dbbak.mdb>
						&nbsp;如备份目录有该文件，将覆盖，如没有，将自动创建<BR>
						&nbsp;&nbsp;<input type=submit value="确定"><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;在上面填写本程序的数据库路径全名，本程序的默认数据库文件为../data/#$%~~data.mdb，<B>请一定不能用默认名称命名备份数据库</B><br>
  						&nbsp;&nbsp;您可以用这个功能来备份您的法规数据，以保证您的数据安全！<br>
  						&nbsp;&nbsp;注意：所有路径都是相对与程序空间根目录的相对路径				</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

sub updata()
		Dbpath=request.form("Dbpath")
		Dbpath=server.mappath(Dbpath)
		bkfolder=request.form("bkfolder")
		bkdbname=request.form("bkdbname")
		Set Fso=server.createobject("scripting.filesystemobject")
		if fso.fileexists(dbpath) then
			If CheckDir(bkfolder) = True Then
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			else
			MakeNewsDir bkfolder
			fso.copyfile dbpath,bkfolder& "\"& bkdbname
			end if
			response.write "备份数据库成功，您备份的数据库路径为" & bkdbname
		Else
			response.write "找不到您所需要备份的文件。"
		End if
end sub
'------------------检查某一目录是否存在-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '存在
       CheckDir = True
    Else
       '不存在
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------根据指定名称生成目录-----------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function


'====================压缩数据库 =========================
sub CompressData()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<form action="Admin_data.asp?action=CompressData" method="post">
<tr>
<td class="txlrow" height=25><b>注意：</b><br>输入数据库所在相对路径,并且输入数据库名称（正在使用中数据库不能压缩，请选择备份数据库进行压缩操作） </td>
</tr>
<tr>
<td class="txlrow">压缩数据库：<input name="dbpath" type="text" value="../data/#$%~~data.mdb" size="30">
&nbsp;
<input type="submit" value="开始压缩"></td>
</tr>
<tr>
<td class="txlrow"><input type="checkbox" name="boolIs97" value="True">如果使用 Access 97 数据库请选择
(默认为 Access 2000 数据库)<br><br></td>
</tr>
<form>
</table>
<%
dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")

If dbpath <> "" Then
dbpath = server.mappath(dbpath)
	response.write(CompactDB(dbpath,boolIs97))
End If

end sub

'=====================压缩参数=========================
Function CompactDB(dbPath, boolIs97)

Dim fso, Engine, strDBPath,JET_3X
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")

	If boolIs97 = "True" Then
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb;" _
		& "Jet OLEDB:Engine Type=" & JET_3X
	Else
		Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
		"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
	End If

fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing

	CompactDB = "你的数据库, " & dbpath & ", 已经压缩成功!" & vbCrLf

Else
	CompactDB = "数据库名称或路径不正确. 请重试!" & vbCrLf
End If

End Function
%>                                                                        
<%
Function IsObjInstalled(strClassString)
On Error Resume Next
IsObjInstalled = False
Err = 0
Dim xTestObj
Set xTestObj = Server.CreateObject(strClassString)
If 0 = Err Then IsObjInstalled = True
Set xTestObj = Nothing
Err = 0
End Function
%>