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
<!--#include file="../inc/md5.asp"-->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<meta http-equiv="Expires" content="0">
<link href="images/css.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
body {font-size: 12px; color: #000; font-family: ����}
td {font-size: 12px; color: #000; font-family: ����;line-height:130%}

.t1 {font:12px ����;color=000000} 
.t2 {font:12px ����;color:ffffff} 
.t3 {font:12px ����;color:336699} 
.t4 {font:12px ����;color:ff0000;} 
.bt1 {font:14px ����;color=000000} 
.bt2 {font:14px ����;color:ffffff} 
.bt3 {font:14px ����;color:336699} 
.bt4 {font:bold 16px ����;color:maroon} 

.td1 {font-size:12px;text-align:right;background-color:#F5F5F5;color:#000000}
.td2 {font-size:12px;text-align:left;background-color:#ffffff;color:#000000;}
.td3 {font-size:12px;text-align:left;background-color:#ffffff;color:#000000;}

A:link {color: #000077}
A:visited {color: #000077}
A:hover {color: #ff0000}
-->
</style>
<script language="javascript">
	function check()
	{
		var obj = document.editform;

		if (obj.bName.value == '')
		{
			alert("�������û��ʺţ�");
			obj.bName.focus();
			return false;
		}

		if (obj.bPassword2.value != obj.bPassword.value)
		{
			alert("������������벻һ�£�");
			obj.bPassword.value = '';
			obj.bPassword2.value = '';
			obj.bPassword.focus();
			return false;
		}

		return true;
	}
</script>
</head>

<body topmargin="5" leftmargin="5" bgcolor="#ffffff">

<%
select case request("go")	
	case "saveedit"
	    call UpdateItem()
	case else
		call EditItem()
end select
call CloseConn()
%>
</body>
</html>
<%
private sub UpdateItem()
	dim ed_bName,ed_bPassword,ed_oldpassword
	ed_bName=request("bName")
	ed_bPassword=request("bPassword")
	ed_oldpassword=request("oldpassword")

	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Admin" 
	rs.open sql,conn,1,3
	if not rs.eof then	
		if trim(rs("bName")&"")<>ed_bName then
			rs("bName")=ed_bName
			rs("bLoginTime")=now
			rs("bLoginIP")=""
			rs("bLoginCount")=0
		end if
		if trim(ed_bPassword)<>trim(ed_oldpassword) then
			rs("bPassword")=md5(ed_bPassword)
		end if
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('����Ա���óɹ������ס�ʺź����룡');window.location.href='admin_password.asp';</script>"
		response.end
	end if
end sub 
%>
<%
private sub EditItem()  
    dim mvarbName,mvarbPassword
    
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Admin"
	rs.open sql,conn,1,1
	if not rs.eof then
		mvarbName=trim(rs("bName")&"")
	    mvarbPassword=trim(rs("bPassword")&"")	 
	end if
	if rs.state<>0 then rs.close
	set rs=nothing
%>
<table width="100%" align="center" border="0" cellspacing="0" cellpadding="2">
	<form autocomplete="off" name="editform" id="editform" method="post" onsubmit="return check();" action="admin_password.asp?go=saveedit">
		<tr>
			<td align="left" valign="top">
			<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
				<tr valign="bottom">
					<td style="font-size:12px;">��ǰλ�ã� 
					<font color="DarkSlateGray" style="font-size:12px"><b>�û�����</b></font>
					<a href="index.asp?go=body">���ع�����ҳ</a> </td>
					<td></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td><hr size="1"></td>
				</tr>
				<tr bgcolor="#898989">
					<td height="24"><font class="t2">&nbsp;�� �� Ա �� ��</font><br>
					</td>
				</tr>
				<tr>
					<td height="10"></td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">
				<tr>
					<td class="td1">�� �� �� �ţ�</td>
					<td class="td2">
					<input type="text" name="bName" id="bName" value="<%=mvarbName%>" size="20">
					<font color="#FF0000">*</font></td>
				</tr>
				<tr>
					<td class="td1">�� �룺</td>
					<td class="td2">
					<input type="password" name="bPassword" id="bPassword" value="<%=mvarbPassword%>" size="20">
					<font color="#FF0000">*</font> ���ִ�Сд</td>
				</tr>
				<tr>
					<td class="td1">ȷ �� �� �룺</td>
					<td class="td2">
					<input type="password" name="bPassword2" id="bPassword2" value="<%=mvarbPassword%>" size="20">
					<font color="#FF0000">*</font> ���ٴ���������(����������һ��)</td>
				</tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
				<tr>
					<td width="150" align="right" height="40"></td>
					<td>
					<input type="hidden" id="oldpassword" name="oldpassword" value="<%=mvarbPassword%>">
					<input type="submit" class="button" name="submit" value="ȷ���ύ">&nbsp;&nbsp;&nbsp;&nbsp;
					<input type="reset" class="button" value="������д" name="Button">
					</td>
				</tr>
			</table>
			</td>
		</tr>
	</form>
</table>
<script language="Javascript">
	editform.bName.focus()
</script>
<%
end sub
%>