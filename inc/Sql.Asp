<%
'--------定义部份------------------
Dim Ay_Application_Value
Dim Ay_Post,Ay_Get,Ay_FilterKeys,Ay_FilterKeys_Keyword,Ay_IsKill,Ay_IsWriteLog
Dim Ay_AlertUrl,Ay_AlertInfo,Ay_KillInfo,Ay_ErrorHandle
Dim Ay_SafePage,Ay_IsSafeOpen,Ay_Sec_Form

'获取配置信息
msql="select * from Ay_SqlConfig"
Set rss=Server.CreateObject("Adodb.Recordset")
rss.Open msql, Conn
if not rss.eof then
	Ay_FilterKeys = trim(rss("bFilterKeys")&"")
	Ay_IsKill = trim(rss("bIsKill")&"") 
	Ay_IsWriteLog = trim(rss("bIsWriteLog")&"")
	Ay_AlertUrl = trim(rss("bAlertUrl")&"")
	Ay_AlertInfo = trim(rss("bAlertInfo")&"")
	Ay_KillInfo = trim(rss("bKillInfo")&"")
	Ay_ErrorHandle = trim(rss("bErrorHandle")&"")
	Ay_SafePage = trim(rss("bSafePage")&"")
	Ay_IsSafeOpen = trim(rss("bIsSafeOpen")&"")
end if
'安全页面参数
Ay_Sec_Form = split(Ay_SafePage,"|")
Ay_FilterKeys_Keyword = split(Ay_FilterKeys,"|")

If Ay_IsKill=1 Then Stop_IP

If Request.Form<>"" Then StopInjection(Request.Form)

If Request.QueryString<>"" Then StopInjection(Request.QueryString)

Function Stop_IP()
	Dim Sqlin_IP,rsKill_IP,Kill_IPsql
	Sqlin_IP=Request.ServerVariables("REMOTE_ADDR")
	Kill_IPsql="select * from Ay_SqlIn where bIPAddress='"&Sqlin_IP&"' and bIsKill=1"
	Set rsKill_IP=conn.execute(Kill_IPsql)
	If Not(rsKill_IP.eof or rsKill_IP.bof) Then
		N_Alert(Ay_KillInfo)
	Response.End
	End If
	rsKill_IP.close	
End Function



'sql通用防注入主函数
Function StopInjection(values)
	Dim Ay_Get,Ay_i
	For Each Ay_Get In values

		If Ay_IsSafeOpen = 1 Then 
			For Ay_i=0 To UBound(Ay_Sec_Form)
				If Instr(LCase(SelfName),Ay_Sec_Form(Ay_i))> 0 Then 
					Exit Function
				else
					Call Select_BadChar(values,Ay_Get)
				End If 
			Next
			
		Else
			Call Select_BadChar(values,Ay_Get)
		End If 
	Next
End Function 

Function Select_BadChar(values,Ay_Get)
	Dim Ay_Xh
	Dim Ay_ip,Ay_url,Ay_sql
	Ay_ip = GetIPAddress
	Ay_url = Request.ServerVariables("URL")

	For Ay_Xh=0 To Ubound(Ay_FilterKeys_Keyword)
		If Instr(LCase(values(Ay_Get)),Ay_FilterKeys_Keyword(Ay_Xh))<>0 Then
			If Ay_IsWriteLog = 1 Then				
				Ay_sql = "insert into Ay_SqlIn(bIPAddress,bPage,bMethod,bParameter,bData,bTime) values('"&Ay_ip&"','"&Ay_url&"','"&intype(values)&"','"&Ay_Get&"','"&N_Replace(values(Ay_Get))&"',now)"
				response.write Ay_sql
				conn.Execute(Ay_sql)
							
			End If			
			N_Alert(Ay_AlertInfo)
			Response.End
		End If
	Next
End Function

'输出警告信息
Function N_Alert(Ay_AlertInfo)
	Dim str
	'response.write "test"
	str = "<"&"Script Language=JavaScript"&">"
	Select Case Ay_ErrorHandle
		Case 1
			str = str & "window.opener=null; window.close();"
		Case 2
			str = str & "alert('"&Ay_AlertInfo&"');window.opener=null; window.close();"
		Case 3
			str = str & "location.href='"&Ay_AlertUrl&"';"
		Case 4
			str = str & "alert('"&Ay_AlertInfo&"');location.href='"&Ay_AlertUrl&"';"
	end Select
	str = str & "<"&"/Script"&">"
	response.write  str
End Function 

'判断注入类型函数
Function intype(values)
	Select Case values
		Case Request.Form
			intype = "Post"
		Case Request.QueryString
			intype = "Get"
	end Select
End Function

'干掉脚本
Function N_Replace(N_urlString)
	N_urlString = Replace(N_urlString,"'","''")
    N_urlString = Replace(N_urlString, ">", "&gt;")
    N_urlString = Replace(N_urlString, "<", "&lt;")
    N_Replace = N_urlString
End Function

'获取本页文件名
Function SelfName()
    SelfName = Mid(Request.ServerVariables("URL"),InstrRev(Request.ServerVariables("URL"),"/")+1)
End Function

%>

    