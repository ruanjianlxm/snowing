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
function checkAll(bCheck)
{
    var arr = document.all.listid;
    for(var i=0; i<arr.length; i++)
    {
        if(!arr[i].disabled == true)
            arr[i].checked = bCheck;
    }
}
function checkform(form)
	{
		var flag=true;
		if(form("bReply").value==""){alert("�ظ����ݲ���Ϊ��!");form("bReply").focus();return false;}		
		return flag;
	}

</script>

</head>

<body topmargin="5" leftmargin="5" bgcolor="#ffffff">

<%
select case request("go")	
	case "edit"
	    call EditItem()
	case "saveedit"
	    call UpdateItem()
	case "delete"
	    call DeleteItem()
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
private sub UpdateItem()
    dim ed_bReply
    
	ed_bReply=request("bReply")
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Book where bId=" & request("id") 
	rs.open sql,conn,1,3
	if not rs.eof then	
		rs("bReply")=ed_bReply	
		rs("bIsReply")=1		
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('���Իظ��ɹ��������������');window.location.href='admin_book.asp';</script>"
		response.end
	end if
end sub
%>
<%
private sub DeleteItem()
    dim de_id
	de_id=clng(Request("id"))
	sql="delete from Ay_Book where bId=" & de_id
	conn.execute sql
	response.write "<script language=javascript>alert('����ɾ���ɹ�!');window.location.href='admin_book.asp';</script>"
end sub
%>
<%
private sub BatchDelete()
    dim mm_ndelid
	mm_ndelid = request.Form("listid")
	if mm_ndelid  = "" then
		response.write "<script language=javascript>alert('û���κ�ѡ��!');window.location.href='admin_book.asp';</script>"
		response.end
	end if
	sql="delete from Ay_Book where bId in (" & Replace(mm_ndelid, "'", "''") & ")"
	conn.execute sql
	response.write "<script language=javascript>alert('����ɾ���ɹ�!');window.location.href='admin_book.asp';</script>"
end sub
%>
<%
private sub ListItem()
%>
<form autocomplete="off" name="form1" id="form1" method="post" action>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>�����ڵ�λ�ã�<font color="DarkSlateGray" style="font-size:12px"><b>���Է���</b></font>&nbsp;<a href="index.asp?go=body">[����]</a>&nbsp;<a href="admin_book.asp">[ˢ���б�]</a> </td>
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
		<td class="t2" align="center">�ο�����</td>
		<td class="t2" align="center">��Դ��ַ</td>
		<td class="t2" align="center">��������</td>
		<td class="t2" align="center">�ظ�����</td>				
		<td class="t2" align="center" width="40px">�ظ�</td>				
		<td class="t2" align="center" width="40px">���</td>
		<td class="t2" align="center">����ʱ��</td>
		<td class="t2" align="center">����</td>
	</tr>
	<%
	dim curpage
	
	sql="select a.* from Ay_Book a order by a.bId desc"
	Set rs=Server.CreateObject("Adodb.Recordset")
	Set mypage=new xdownpage
	mypage.getconn=conn
	mypage.getsql=sql
	mypage.pagesize=20
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
				Response.Write("<tr height='23' bgcolor=fefefe>")
			else
				Response.Write("<tr height='23' bgcolor=efefef>")
			end if	
	%>			
		<td width="10" height="20" align="center">
		<input name="listid" type="checkbox" id="listid" value="<%=(rs("bId"))%>">
		</td>
		<td width="40" align="center"><%=trim(cstr(i + (curpage -1) * mypage.pagesize))%></td>
		<td align="center"><%=trim(rs("bGuest") & "")%></td>
		<td align="center"><%=trim(rs("bIpAddress") & "")%></td>
		<td align="left"><%=GotTopic(trim(rs("bContent") & ""),10)%></td>	
		<td align="left"><%=GotTopic(trim(rs("bReply") & ""),10)%></td>			
		<td align="center">
		<input type="checkbox" disabled <%if rs("bIsReply")=1 then response.write "checked" end if%> /></td>				
		<td align="center">
		<input type="checkbox" disabled <%if rs("bIsPass")=1 then response.write "checked" end if%> /></td>
		<td align="center" width="150px"><%=trim(rs("bAddTime") & "")%></td>
		<td align="center" width="80px">
		<a href="admin_book.asp?go=edit&id=<%=trim(rs("bId")&"")%>">�ظ�</a>
		<a href="admin_book.asp?go=delete&id=<%=trim(rs("bId")&"")%>" onclick="javascript:return confirm('��ȷ��ɾ������ ?')">
		ɾ��</a> </td>
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
		<input type="submit" name="mydele" class="button" value="����ɾ��" onclick="javascript:this.form.action='admin_book.asp?go=batchdelete';return confirm('��ȷ��ɾ������ ?')">
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
private sub EditItem()
    dim mvarbName,mvarbIpAddress,mvarbContent,mvarbReply
       
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Book  where bId=" & request("id")
	rs.open sql,conn,1,1
	if not rs.eof then
	    mvarbName=trim(rs("bGuest")&"")
	    mvarbIpAddress=trim(rs("bIpAddress")&"")
	    mvarbContent=trim(rs("bContent")&"")	
	    mvarbReply=trim(rs("bReply")&"")
	end if
	if rs.state<>0 then rs.close
	set rs=nothing
%>
<form autocomplete="off" name="editform" id="editform" method="post" onsubmit="return checkform(editform)" action="admin_book.asp?go=saveedit&id=<%=request("id")%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>�����ڵ�λ�ã�<font color="DarkSlateGray" style="font-size:12px"><b>���Է���</b></font>&nbsp;->&nbsp;�ظ�����&nbsp;&nbsp;��&nbsp;<a href="admin_book.asp">�����б�</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td><hr size="1"></td></tr>
	<tr bgcolor="#898989"><td height="23"><font class="t2">&nbsp;��ϸ����</font></td></tr>
	<tr><td height="10"></td></tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">	
	<tr height="22">
		<td class="td1" align="right" valign="middle">�� �� �� �ƣ�</td>
		<td class="td2"><%=mvarbName%></td>
	</tr> 
	<tr height="22">
		<td class="td1" align="right" valign="middle">�� Դ �� ַ��</td>
		<td class="td2"><%=mvarbIpAddress%></td>
	</tr> 
	<tr height="22">
		<td class="td1" align="right" valign="middle">�� �� �� �ݣ�</td>
		<td class="td2">
		<textarea name="bContent" cols="60" rows="5" id="bContent" disabled><%=mvarbContent%></textarea>
		</td>
	</tr> 							
	<tr>
		<td align="right" class="td1" valign="top">վ �� �� ����</td>
		<td class="td2">			
		<textarea name="bReply" cols="60" rows="5" id="bReply"><%=mvarbReply%></textarea>					
		</td>
	</tr>
	
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td width="150" align="right" height="40"></td>
		<td>
		<input type="submit" class="button" name="submit" value="ȷ���ύ">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="reset" class="button" value="������д" name="Button">
		</td>
	</tr>				
</table>
</form> 
<script language="Javascript">
	editform.bReply.focus()
</script>
<%
end sub
%>