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
<html>

<head>
<title>��վ��Ϣ</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
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
</script>
</head>

<body topmargin="5" leftmargin="5" bgcolor="#ffffff">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>�����ڵ�λ�ã�<b><%=Request.ServerVariables("Http_HOST")%> -&gt; ��վ����</b></td>
		<td></td>
	</tr>
</table>
<table width="100%" align="center" border="0" cellspacing="2" cellpadding="0">
	<tr>
		<td rowspan="3" width="120" align="center"><img src="images/admin_p.gif" width="90" height="100"></td>
		<td rowspan="3" width="100">��</td>
		<td style="color:#191970;" height="30"><%=year(now())%>��<%=month(now())%>��<%=day(now())%>��<%=hour(now())%>:<%=minute(now())%></td>
	</tr>
	<tr>
		<td class="font2" height="50">��վ��ز�������</td>
	</tr>
	<tr>
		<td class="font1" height="30"><a href="admin_system.asp?go=base" class="r1">[��������]</a> <a href="admin_system.asp?go=master" class="r1">[վ������]</a></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr height="12" bgcolor="#EEEEEE">
		<td></td>
	</tr>
	<tr height="25" bgcolor="#31615A">
		<td  style="color:#FFFFFF;padding-left:10px;" valign="middle">�����Ϣ</td>
	</tr>
</table> 
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="0">
	<form autocomplete="off" name="form1" id="form1" method="post" action="admin_system.asp">
		<tr>
			<td align="left" valign="top">			
			<%
			select case request("go")	
				case "base"
				    call BaseInfomation()
				case "updatebase"
					call UpdateBaseInfo()
				Case "master"
					Call StationMaster()			
				Case "updatemaster"
					Call UpdateMaster()
				Case "updatenotice"
					Call UpdateNotice()
				case else
					call BaseInfomation()
			end select
			call CloseConn()
			%> </td>
		</tr>
	</form>
</table>

</body>

</html>
<%
private sub UpdateBaseInfo()
   	dim ed_bName,ed_bTitle,ed_bUrl,ed_bAuthor,ed_bKeywords,ed_bDescriptions
    dim ed_bMiibeian
    
	ed_bName=request("bName")
	ed_bTitle=request("bTitle")
	ed_bUrl=request("bUrl")
	ed_bAuthor=request("bAuthor")
	ed_bKeywords=request("bKeywords")
	ed_bDescriptions=request("bDescriptions")
	ed_bMiibeian=request("bMiibeian")
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_System" 
	rs.open sql,conn,1,3
	if not rs.eof then	
		rs("bName")=ed_bName
		rs("bTitle")=ed_bTitle
		rs("bUrl")=ed_bUrl
		rs("bAuthor")=ed_bAuthor
		rs("bKeywords")=ed_bKeywords
		rs("bDescriptions")=ed_bDescriptions
			
		rs("bMiibeian")=ed_bMiibeian
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('��վ���óɹ��������������');window.location.href='admin_system.asp?go=base';</script>"
		response.end
	end if
end sub
%> <%
private sub UpdateMaster()
   	dim ed_bUserName,ed_bEmail,ed_bPhone,ed_bAddress
    dim ed_bReplacewords
	dim ed_bInformation,ed_bPhoto
	
    ed_bUserName=request("bUserName")
	ed_bEmail=request("bEmail")
	ed_bPhone=request("bPhone")
	ed_bAddress=request("bAddress")
	ed_bReplacewords=request("bReplacewords")	
	ed_bInformation=request("bInformation")
	ed_bPhoto=request("bPic")
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_System" 
	rs.open sql,conn,1,3
	if not rs.eof then	
		rs("bUserName")=ed_bUserName
		rs("bEmail")=ed_bEmail
		rs("bPhone")=ed_bPhone
		rs("bAddress")=ed_bAddress		
		rs("bReplacewords")=ed_bReplacewords		
		rs("bInformation")=ed_bInformation
		rs("bPhoto")=ed_bPhoto
		
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('վ�����ϸ��³ɹ��������������');window.location.href='admin_system.asp?go=master';</script>"
		response.end
	end if
end sub
%> 
<%
private sub BaseInfomation()
	dim mvarbName,mvarbTitle,mvarbUrl,mvarbAuthor,mvarbKeywords,mvarbDescriptions
    dim mvarbMiibeian,mvarbNeedPass
 
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_System"
	rs.open sql,conn,1,1
	if not rs.eof then
		mvarbName=trim(rs("bName")&"")
	    mvarbTitle=trim(rs("bTitle")&"")
	    mvarbUrl=trim(rs("bUrl")&"")
	    mvarbAuthor=trim(rs("bAuthor")&"")
	    mvarbKeywords=trim(rs("bKeywords")&"")
	    mvarbDescriptions=trim(rs("bDescriptions")&"")	
	    mvarbMiibeian=trim(rs("bMiibeian")&"")	 
	    mvarbNeedPass=trim(rs("bNeedPass")&"")
	end if
	if rs.state<>0 then rs.close
	set rs=nothing

%>
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">
	<tr>
		<td class="td1" align="right" valign="middle">վ �� �� �ƣ�</td>
		<td class="td2">
		<input type="text" name="bName" id="bName" size="50" value="<%=mvarbName%>">
		</td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle">վ �� �� �⣺</td>
		<td class="td2">
		<input type="text" name="bTitle" id="bTitle" size="70" value="<%=mvarbTitle%>">
		</td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle">վ �� �� ����</td>
		<td class="td2">
		<input type="text" name="bUrl" id="bUrl" size="40" value="<%=mvarbUrl%>">
		<button class="button" onclick="bUrl.value='http://'">http://</button>
		</td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle">�� Ȩ �� �ߣ�</td>
		<td class="td2">
		<input type="text" name="bAuthor" id="bAuthor" size="50" value="<%=mvarbAuthor%>">
		</td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle">�� վ �� ����</td>
		<td class="td2">
		<input type="text" name="bMiibeian" id="bMiibeian" size="40" value="<%=mvarbMiibeian%>">
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">��վ �ؼ��֣�</td>
		<td class="td2">
		<textarea name="bKeywords" cols="60" rows="4" id="bKeywords"><%=mvarbKeywords%></textarea>
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">�� վ �� ����</td>
		<td class="td2">
		<textarea name="bDescriptions" cols="60" rows="3" id="bDescriptions"><%=mvarbDescriptions%></textarea>
		</td>
	</tr>	
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td width="150" align="right" height="40"></td>
		<td><input type="hidden" name="go" id="go" value="updatebase">
		<input type="submit" class="button" name="submit" value="ȷ���ύ">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="reset" class="button" value="������д" name="Button"> </td>
	</tr>
</table>
<%
end sub
%> <%
private sub StationMaster()
	dim mvarbUserName,mvarbEmail,mvarbPhone,mvarbAddress
    dim mvarbReplacewords
 	dim mvarbInformation,mvarbPhoto
 	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_System"
	rs.open sql,conn,1,1
	if not rs.eof then
		mvarbUserName=trim(rs("bUserName")&"")
	    mvarbEmail=trim(rs("bEmail")&"")
	    mvarbPhone=trim(rs("bPhone")&"")
	    mvarbAddress=trim(rs("bAddress")&"")	    
	    mvarbReplacewords=trim(rs("bReplacewords")&"")	   
	    mvarbInformation=trim(rs("bInformation") & "")
	    mvarbPhoto=trim(rs("bPhoto")&"")
	end if
	if rs.state<>0 then rs.close
	set rs=nothing

%>
<table width="100%">
	<tr>
		<td valign="top">
		<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">
			<tr>
				<td class="td1" align="right" valign="middle">վ�����ƣ�</td>
				<td class="td2">
				<input type="text" name="bUserName" id="bUserName" size="40" value="<%=mvarbUserName%>">
				</td>
			</tr>
			<tr>
				<td align="right" class="td1" valign="top">���˽��գ�</td>
				<td class="td2">
				<input type="text" class="input" id="bPic" name="bPic" style="width:300px;" />
				<input type="checkbox" onclick="display('upload');" id="box" /><label for="box">�ϴ�ͼƬ</label>
				<br>
				<div id="upload" style="display:none;" class="td2">
					<iframe src="upload.asp?go=pic" frameborder="0" style="height:22px;width:100%;" scrolling="no">
					</iframe></div>
				</td>
			</tr>
			<tr>
				<td class="td1" align="right" valign="top">���˽��ܣ�</td>
				<td class="td2">
				<textarea name="bInformation" cols="60" rows="5" id="bInformation"><%=mvarbInformation%></textarea>
				</td>
			</tr>
			<tr>
				<td class="td1" align="right" valign="middle">�����ַ��</td>
				<td class="td2">
				<input type="text" name="bEmail" id="bEmail" size="40" value="<%=mvarbEmail%>">
				</td>
			</tr>
			<tr>
				<td class="td1" align="right" valign="middle">��ϵ�绰��</td>
				<td class="td2">
				<input type="text" name="bPhone" id="bPhone" size="40" value="<%=mvarbPhone%>">
				</td>
			</tr>
			<tr>
				<td class="td1" align="right" valign="middle">��˾��ַ��</td>
				<td class="td2">
				<input type="text" name="bAddress" id="bAddress" size="40" value="<%=mvarbAddress%>">
				</td>
			</tr>
			<tr>
				<td align="right" class="td1" valign="top">���Թؼ��ֹ��ˣ�</td>
				<td class="td2">
				<textarea name="bReplacewords" cols="60" rows="5" id="bReplacewords"><%=mvarbReplacewords%></textarea>
				</td>
			</tr>			
		</table>
		<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
			<tr>
				<td width="150" align="right" height="40"></td>
				<td><input type="hidden" name="go" id="go" value="updatemaster">
				<input type="submit" class="button" name="submit" value="ȷ���ύ">&nbsp;&nbsp;&nbsp;&nbsp;
				<input type="reset" class="button" value="������д" name="Button">
				</td>
			</tr>
		</table>
		</td>
		<td valign="top" width="140px">
		<img src="../<%=mvarbPhoto%>" width="140px" height="120px" onerror="this.src='../images/none.gif'">
		</td>
	</tr>
</table>
<%
end sub
%>