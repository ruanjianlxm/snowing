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
<!--#include file="conn.asp"-->
<!--#include file="cls_page.asp" -->

<html>
<head>
<title>��վ����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<meta http-equiv="Expires" content="0">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="Javascript">
function display(ID)
{
	if (document.getElementById(ID).style.display == "none") {
		document.getElementById(ID).style.display = "";
	}else{
		document.getElementById(ID).style.display = "none";
	}
}

function checkAll(bCheck)
{
    var arr = document.all.listid;
    if (typeof(arr)!="undefined")
    {
	    for(var i=0; i<arr.length; i++)
	    {
	        if(!arr[i].disabled == true)
	           arr[i].checked = bCheck;
	    }
    }
}
function checkform(form)
	{
		var flag=true;				
		if(form("bClassID").value=="0"){alert("��ѡ�����!");form("bClassID").focus();return false;}	
		if(form("bTitle").value==""){alert("�������Ʋ���Ϊ��!");form("bTitle").focus();return false;}		
		return flag;
	}

</script>

</head>

<body topmargin="5" leftmargin="5" bgcolor="#ffffff">

<%
select case request("go")
	case "unlock"
		call UnLock()
	case "lock"
	    call Lock()
	case "edit"
	    call EditSetting()	
	case "saveedit"
	    call UpdateSetting()	
	case "batchdelete"
	    call BatchDelete()
	case else
		call ListItem()
end select
call CloseConn()
%>
</body>
</html>
<%
private sub UnLock()
	sql="update  Ay_SqlIn set bIsKill=0 where bId=" & request("id")	
	conn.execute sql
	response.write "<script>window.location.href='admin_sql.asp';</script>"
	response.end
end sub
%>
<%
private sub Lock()
	sql="update  Ay_SqlIn set bIsKill=1 where bId=" & request("id")	
	conn.execute sql
	response.write "<script>window.location.href='admin_sql.asp';</script>"
	response.end
end sub
%>
<%
private sub UpdateSetting()
    dim ed_bIsKill,ed_bIsSafeOpen,ed_bSafePage,ed_bKillInfo	
	dim ed_bAlertInfo,ed_bAlertUrl,ed_bIsWriteLog,ed_bErrorHandle,ed_bFilterKeys
		
	ed_bIsKill=request("bIsKill")		
	ed_bIsSafeOpen=request("bIsSafeOpen")		
	ed_bSafePage=request("bSafePage")	
	ed_bKillInfo=request("bKillInfo")	
	ed_bAlertInfo=request("bAlertInfo")	
	ed_bAlertUrl=request("bAlertUrl")		
	ed_bIsWriteLog=request("bIsWriteLog")
	ed_bErrorHandle=request("bErrorHandle")
	ed_bFilterKeys=request("bFilterKeys")	
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_SqlConfig "
	rs.open sql,conn,1,3
	if not rs.eof then	
		rs("bIsKill")=ed_bIsKill		
		rs("bIsSafeOpen")=ed_bIsSafeOpen	
		rs("bSafePage")=ed_bSafePage	
		rs("bKillInfo")=ed_bKillInfo	
		rs("bAlertInfo")=ed_bAlertInfo	
		rs("bAlertUrl")=ed_bAlertUrl	
		rs("bIsWriteLog")=ed_bIsWriteLog
		rs("bErrorHandle")=ed_bErrorHandle
		rs("bFilterKeys")=ed_bFilterKeys		
		
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('SQL��ע�����óɹ��������������');window.location.href='admin_sql.asp?go=edit';</script>"
		response.end
	end if
end sub
%>
<%
private sub BatchDelete()
    dim mm_ndelid
	mm_ndelid = request.Form("listid")
	if trim(mm_ndelid)  = "" then
		response.write "<script language=javascript>alert('û���κ�ѡ��!');window.location.href='admin_sql.asp';</script>"
		response.end
	end if
	sql="delete from Ay_SqlIn where bId in (" & Replace(mm_ndelid, "'", "''") & ")"
	
	conn.execute sql
	response.write "<script language=javascript>alert('����ɾ���ɹ�!');window.location.href='admin_sql.asp';</script>"
end sub
%>

<%
private sub ListItem()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<form autocomplete="off" name="form1" id="form1" method="post" action>
	<tr valign="bottom">
		<td>�����ڵ�λ�ã�<font color="DarkSlateGray" style="font-size:12px"><b>SQLע�����</b></font>&nbsp;<a href="index.asp?go=body">[����]</a>&nbsp;<a href="admin_sql.asp">[ˢ���б�]</a></td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td><hr size="1"></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="1">
	<tr bgcolor="527c72" style="color:#ffffff" align="center" height="23">
		<td width="10" align="center"></td>
		<td class="t2" width="40">���</td>				
		<td class="t2" align="center">�����ɣ�</td>
        <td class="t2" align="center">��ǰ״̬</td>
        <td class="t2" align="center">�Ƿ�����</td>
        <td class="t2" align="center">����ҳ��</td>
        <td class="t2" align="center">����ʱ��</td>
        <td class="t2" align="center">�ύ��ʽ</td>
        <td class="t2" align="center">�ύ����</td>
        <td class="t2" align="center">�ύ����</td>

	</tr>
	<%
	dim curpage
	
	sql="select *  from Ay_SqlIn a order by a.bId desc"
	Set rs=Server.CreateObject("Adodb.Recordset")
	Set mypage=new xdownpage
	mypage.getconn=conn
	mypage.getsql=sql
	mypage.pagesize=16
	set rs=mypage.getrs()	
	if request("page")<>"" then
		curpage=clng(request("page")+0)
	else
		curpage=1
	end if		
	if rs.eof and rs.bof then
		Response.Write("<tr height='25' bgcolor=efefef>")
		response.write("<td align='center' colspan=15>�Ҳ����κμ�¼��</td>")
		response.write("</tr>")
	end if
	for i=1 to mypage.pagesize
		if rs.eof or rs.bof then exit for		
			if (i mod 2)=0 then
				Response.Write("<tr height='23' bgcolor='#fefefe' ")
			else
				Response.Write("<tr height='23' bgcolor='#efefef' ")
			end if				
			response.write ">"
	%>			
		<td width="10" height="20" align="center">
		<input name="listid" type="checkbox" id="listid" value="<%=(rs("bId"))%>">
		</td>
		<td width="40" align="center"><%=trim(cstr(i + (curpage -1) * mypage.pagesize))%></td>	
		
		<td align="center"><%=trim(rs("bIPAddress") & "")%></td>
		<td align="center">
			<%	if rs("bIsKill")=1 then 
					response.write "<font color='red'>������</font>"
				else
					response.write "<font color='green'>�ѽ���</font>"
				end if
			%>
		</td>		
		<td align="center">
			<%	if rs("bIsKill")=1 then 
					response.write "<a href=admin_sql.asp?go=unlock&id="&rs("bId")&" style=""color:#FF0000"">����IP</a>"
				else
					response.write "<a href=admin_sql.asp?go=lock&id="&rs("bId")&" style=""color:#006600"">����IP</a>"
				end if
			%>
		</td>		
		<td align="center"><%=trim(rs("bPage") & "")%></td>	
		<td align="center"><%=trim(rs("bTime") & "")%></td>	
		<td align="center"><%=trim(rs("bMethod") & "")%></td>
		<td align="center"><%=trim(rs("bParameter") & "")%></td>
		<td align="center"><%=N_Replace(trim(rs("bData") & ""))%></td>
	</tr>
	<%	
		rs.MoveNext 
	next		
	if rs.State <>0 then rs.Close 
	set rs=nothing
	%>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="30">
	<tr>
		<td colspan="2" bgcolor="696969" height="1"></td>
	</tr>
	<tr>
		<td>
		<input type="submit" name="mydele" class="button" value="����ɾ��" onclick="javascript:this.form.action='admin_sql.asp?go=batchdelete';return confirm('��ȷ��ɾ������ ?')">
		<input type="checkbox" name="all" id="all" onclick="checkAll(this.checked)"><label for="all">ȫѡ</label>
		</td>
		<td align="right" height="30"><%mypage.showpage()%></td>
	</tr>
</table>
</form>
<%
end sub
%> 
<%
private sub EditSetting()
	dim mvarbIsKill,mvarbIsSafeOpen,mvarbSafePage,mvarbKillInfo,mvarbAlertInfo
	dim mvarbAlertUrl,mvarbIsWriteLog,mvarbErrorHandle,mvarbFilterKeys	
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_SqlConfig "
	rs.open sql,conn,1,1
	if not rs.eof Then	
	    mvarbIsKill=trim(rs("bIsKill")&"")		   
	    mvarbIsSafeOpen=trim(rs("bIsSafeOpen")&"")	    
	    mvarbSafePage=trim(rs("bSafePage")&"")
	    mvarbKillInfo=trim(rs("bKillInfo")&"")
	    mvarbAlertInfo=trim(rs("bAlertInfo")&"") 
	    mvarbAlertUrl=trim(rs("bAlertUrl")&"")	    
	    mvarbIsWriteLog=trim(rs("bIsWriteLog")&"")
		mvarbErrorHandle=trim(rs("bErrorHandle")&"")
	    mvarbFilterKeys=trim(rs("bFilterKeys")&"")	   	
	end if
	if rs.state<>0 then rs.close
	set rs=nothing
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>�����ڵ�λ�ã�<font color="DarkSlateGray" style="font-size:12px"><b>SQL��ע������</b></font>&nbsp;->&nbsp;�༭����&nbsp;&nbsp;��&nbsp;<a href="admin_sql.asp?go=edit">�����б�</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td><hr size="1"></td></tr>
	<tr bgcolor="#898989"><td height="23"><font class="t2">&nbsp;��ϸ����</font></td></tr>
	<tr><td height="10"></td></tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">	
<form autocomplete="off" name="editform" id="editform" method="post">	 					
	<tr>
		<td align="right" class="td1" valign="middle">��Ҫ���˵Ĺؼ��֣���</td>
		<td class="td2">
		<input name="bFilterKeys" type="text" value="<%=mvarbFilterKeys%>" id="bFilterKeys" style=" " size="50">
                  ��&quot;|&quot;�ֿ�
		</td>			
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle">�Ƿ��¼��������Ϣ��</td>
		<td class="td2">
		    <select name="bIsWriteLog" id="bIsWriteLog">
              <option value="1" <%if mvarbIsWriteLog="1" Then response.write "selected"%>>��</option>
              <option value="0" <%if mvarbIsWriteLog="0" Then response.write "selected"%>>��</option>
          </select>
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">�Ƿ���������IP��</td>
		<td class="td2">
		<select name="bIsKill" id="bIsKill">
          <option value="1" <%if mvarbIsKill="1" Then response.write "selected"%>>��</option>
          <option value="0" <%if mvarbIsKill="0" Then response.write "selected"%>>��</option>
      </select>
        </td>
	</tr>	
	<tr>
		<td width="15%" class="td1" align="right">�Ƿ����ð�ȫҳ�棺</td>
		<td width="85%" class="td2">
		<select name="bIsSafeOpen" id="bIsSafeOpen">
                      <option value="1" <%if mvarbIsSafeOpen="1" Then response.write "selected"%>>��</option>
                      <option value="0" <%if mvarbIsSafeOpen="0" Then response.write "selected"%>>��</option>
                    </select>
                  ����������ܣ��������ȷ�ϴ�ҳ��������ˣ���ȷ���԰�ȫûӰ�죡
		</td>
 	</tr> 
  	<tr>
		<td align="right" class="td1">����Ϊ��ȫ��ҳ�棺</td>
		<td class="td2">
		<input name="bSafePage" type="text" value="<%=mvarbSafePage%>" id="bSafePage" style=" " size="50">
                  ��&quot;|&quot;�ֿ�
		</td>
  	</tr> 
	<tr>
		<td align="right" class="td1">�����Ĵ���ʽ��</td>
		<td class="td2">
		<select name="bErrorHandle" id="bErrorHandle">
          <option value="1" <%if mvarbErrorHandle="1" Then response.write "selected"%>>ֱ�ӹر���ҳ</option>
          <option value="2" <%if mvarbErrorHandle="2" Then response.write "selected"%>>�����ر�</option>
          <option value="3" <%if mvarbErrorHandle="3" Then response.write "selected"%>>��ת��ָ��ҳ��</option>
          <option value="4" <%if mvarbErrorHandle="4" Then response.write "selected"%>>�������ת</option>
      </select>
		</td>
	</tr>	
	<tr>
		<td align="right" class="td1" valign="top">�������תUrl��</td>
		<td class="td2">			
			<input name="bAlertUrl" type="text" value="<%=mvarbAlertUrl%>" id="bAlertUrl"  size="30">
		</td>
	</tr>
 	<tr>
		<td align="right" class="td1" valign="top">������ʾ��Ϣ��</td>
		<td class="td2">			
			<textarea name="bAlertInfo" cols="45" rows="4" id="bAlertInfo"><%=mvarbAlertInfo%></textarea>                  
                  \n\n����
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">��ֹ������ʾ��Ϣ��</td>
		<td class="td2">			
			<textarea name="bKillInfo" cols="45" rows="4" id="bKillInfo"><%=mvarbKillInfo%></textarea>
                  \n\n����
		</td>
	</tr>

</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td width="150" align="right" height="40"></td>
		<td>
		<input type="button" class="button" name="submit1" value="ȷ���ύ" onclick="this.form.action='admin_sql.asp?go=saveedit';this.form.submit();">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="reset" class="button" value="������д" name="Button">
		</td>
	</tr>				
</table>
</form>  
<script language="Javascript">
	editform.bFilterKeys.focus()
</script>
<%
end sub
Function N_Replace(N_urlString)
	N_urlString = Replace(N_urlString,"'","''")
    N_urlString = Replace(N_urlString, ">", "&gt;")
    N_urlString = Replace(N_urlString, "<", "&lt;")
    N_Replace = N_urlString
End Function

%>