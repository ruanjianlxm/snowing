<%@ Language=VBScript %> <%
Response.Buffer = true
'���û���
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
<title>���ݿ����ҳ��</title>
<link href="images/css.css" rel="stylesheet" type="text/css">
<meta NAME=GENERATOR Content="Microsoft FrontPage 6.0" CHARSET=GB2312>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="0" marginheight="0" marginwidth="0">
<%
dim action
action=trim(request("action"))

dim dbpath,bkfolder,bkdbname,fso,fso1

select case action
case "CompressData"		'ѹ������
	
		dim tmprs
		dim allarticle
		dim Maxid
		dim topic,dateandtime,body
		call CompressData()


case "BackupData"		'��������
		if request("act")="Backup" then
			call updata()
		else
			call BackupData()
		end if

case "RestoreData"		'�ָ�����
	dim backpath
		if request("act")="Restore" then
			Dbpath=request.form("Dbpath")
			backpath=request.form("backpath")
			if dbpath="" then
			response.write "��������Ҫ�ָ��ɵ����ݿ�ȫ��"	
			else
			Dbpath=server.mappath(Dbpath)
			end if
			backpath=server.mappath(backpath)
		
			Set Fso=server.createobject("scripting.filesystemobject")
			if fso.fileexists(dbpath) then  					
			fso.copyfile Dbpath,Backpath
			response.write "�ɹ��ָ����ݣ�"
			else
			response.write "����Ŀ¼�²������ı����ļ���"	
			end if
		else
		call RestoreData()
		end if
case else
	response.Write "��ѡȡ��Ӧ�Ĳ���"

end select

response.write"</body></html>"

'====================�ָ����ݿ�=========================
sub RestoreData()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<tr>
	<th height=25 >
   					&nbsp;&nbsp;<B>�ָ���վ����</B>( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
  					</th>
  				</tr>
				<form method="post" action="ADMIN_data.asp?action=RestoreData&act=Restore">
  				
  				<tr>
  					<td height=100 class="txlrow">
  						&nbsp;&nbsp;�������ݿ�·��(���)��<input type=text size=30 name=DBpath value="../data/backup/dbbak.mdb">
  						&nbsp;&nbsp;<BR>
  						&nbsp;&nbsp;Ŀ�����ݿ�·��(���)��<input type=text size=30 name=backpath value="../data/#$%~~data.mdb">
  						<BR>&nbsp;&nbsp;��д����ǰʹ�õ����ݿ�·�����粻�븲�ǵ�ǰ�ļ���������������ע��·���Ƿ���ȷ����Ȼ���޸�conn.asp�ļ������Ŀ���ļ����͵�ǰʹ�����ݿ���һ�µĻ��������޸�conn.asp�ļ�<BR>
						&nbsp;&nbsp;<input type=submit value="�ָ�����"> <br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;��������д����������ݿ�·��ȫ�����������Ĭ�ϱ������ݿ��ļ�Ϊdata\backup\db.mdb���밴�����ı����ļ������޸ġ�<br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ��Ŀ¼�����·��</font>
  					</td>
  				</tr>	
  				</form>
  			</table>
<%
end sub

'====================�������ݿ�=========================
sub BackupData()
%>
	<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
  				<tr>
  					<th height=25 >
  					&nbsp;&nbsp;<B>������վ����</B>( ��ҪFSO֧�֣�FSO��ذ����뿴΢����վ )
  					</th>
  				</tr>
  				<form method="post" action="ADMIN_data.asp?action=BackupData&act=Backup">
  				<tr>
  					<td height=100 class="txlrow">
  						&nbsp;&nbsp;
						��ǰ���ݿ�·��(���·��)��<input type=text size=30 name=DBpath value="../data/#$%~~data.mdb">
						<BR>
						&nbsp;&nbsp;
						�������ݿ�Ŀ¼(���·��)��<input type=text size=30 name=bkfolder value="../data/backup">
						&nbsp;��Ŀ¼�����ڣ������Զ�����<BR>
						&nbsp;&nbsp;
						�������ݿ�����(��д����)��<input type=text size=30 name=bkDBname value=dbbak.mdb>
						&nbsp;�籸��Ŀ¼�и��ļ��������ǣ���û�У����Զ�����<BR>
						&nbsp;&nbsp;<input type=submit value="ȷ��"><br>
  						-----------------------------------------------------------------------------------------<br>
  						&nbsp;&nbsp;��������д����������ݿ�·��ȫ�����������Ĭ�����ݿ��ļ�Ϊ../data/#$%~~data.mdb��<B>��һ��������Ĭ�����������������ݿ�</B><br>
  						&nbsp;&nbsp;������������������������ķ������ݣ��Ա�֤�������ݰ�ȫ��<br>
  						&nbsp;&nbsp;ע�⣺����·��������������ռ��Ŀ¼�����·��				</font>
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
			response.write "�������ݿ�ɹ��������ݵ����ݿ�·��Ϊ" & bkdbname
		Else
			response.write "�Ҳ���������Ҫ���ݵ��ļ���"
		End if
end sub
'------------------���ĳһĿ¼�Ƿ����-------------------
Function CheckDir(FolderPath)
	folderpath=Server.MapPath(".")&"\"&folderpath
    Set fso1 = CreateObject("Scripting.FileSystemObject")
    If fso1.FolderExists(FolderPath) then
       '����
       CheckDir = True
    Else
       '������
       CheckDir = False
    End if
    Set fso1 = nothing
End Function
'-------------����ָ����������Ŀ¼-----------------------
Function MakeNewsDir(foldername)
	dim f
    Set fso1 = CreateObject("Scripting.FileSystemObject")
        Set f = fso1.CreateFolder(foldername)
        MakeNewsDir = True
    Set fso1 = nothing
End Function


'====================ѹ�����ݿ� =========================
sub CompressData()
%>
<table border="0"  cellspacing="1" cellpadding="5" height="1" align=center width="95%" class="tableBorder">
<form action="Admin_data.asp?action=CompressData" method="post">
<tr>
<td class="txlrow" height=25><b>ע�⣺</b><br>�������ݿ��������·��,�����������ݿ����ƣ�����ʹ�������ݿⲻ��ѹ������ѡ�񱸷����ݿ����ѹ�������� </td>
</tr>
<tr>
<td class="txlrow">ѹ�����ݿ⣺<input name="dbpath" type="text" value="../data/#$%~~data.mdb" size="30">
&nbsp;
<input type="submit" value="��ʼѹ��"></td>
</tr>
<tr>
<td class="txlrow"><input type="checkbox" name="boolIs97" value="True">���ʹ�� Access 97 ���ݿ���ѡ��
(Ĭ��Ϊ Access 2000 ���ݿ�)<br><br></td>
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

'=====================ѹ������=========================
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

	CompactDB = "������ݿ�, " & dbpath & ", �Ѿ�ѹ���ɹ�!" & vbCrLf

Else
	CompactDB = "���ݿ����ƻ�·������ȷ. ������!" & vbCrLf
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