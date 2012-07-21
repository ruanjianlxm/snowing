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
<!--#include file="conn.asp"-->
<!--#include file="cls_page.asp" -->
<html>
<head>
<title>网站资料</title>
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
		if(form("bTitle").value==""){alert("标题不能为空!");form("bTitle").focus();return false;}
		if(form("bOrder").value==""){alert("排序只能输入数字!");form("bOrder").focus();return false;}
		return flag;
	}

</script>

</head>

<body topmargin="5" leftmargin="5" bgcolor="#ffffff">

<%
select case request("go")
	case "add"
		call AddItem()
	case "saveadd"
	    call SaveItem()
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
private sub SaveItem()
    dim ad_bTitle,ad_bOrder,ad_bContent
       
	ad_bTitle=request("bTitle")
	ad_bOrder=request("bOrder")
	ad_bContent=request("bContent")
	
	set rs=server.createobject("adodb.recordset")
	sql="select top 1 * from Ay_Notice where bTitle='" & ad_bTitle & "'"	
	rs.open sql,conn,1,3	
	rs.addnew	
	rs("bTitle")=ad_bTitle
	rs("bOrder")=ad_bOrder
	rs("bContent")=ad_bContent
	rs("bUrl")="../about/?" & rs("bId") & ".html"
	rs("bAddTime")=now
	rs("bAddUser")=session("admin")		
	rs.update
	if rs.state<>0 then rs.close
	set rs=nothing
	response.write "<script>alert('公告添加成功，请继续操作！');window.location.href='admin_notice.asp';</script>"
	response.end
end sub
%>
<%
private sub UpdateItem()
    dim ed_bTitle,ed_bOrder,ed_bContent
    
	ed_bTitle=request("bTitle")
	ed_bOrder=request("bOrder")
	ed_bContent=request("bContent")
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Notice where bId=" & request("id") 
	rs.open sql,conn,1,3
	if not rs.eof then	
		rs("bTitle")=ed_bTitle
		rs("bOrder")=ed_bOrder
		rs("bContent")=ed_bContent
		rs("bUrl")="../about/?" & rs("bId") & ".html"
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('公告更新成功，请继续操作！');window.location.href='admin_notice.asp';</script>"
		response.end
	end if
end sub
%>
<%
private sub DeleteItem()
    dim de_id
	de_id=clng(Request("id"))
	sql="delete from Ay_Notice where bId=" & de_id
	conn.execute sql
	response.write "<script language=javascript>alert('公告删除成功!');window.location.href='admin_notice.asp';</script>"
end sub
%>
<%
private sub BatchDelete()
    dim mm_ndelid
	mm_ndelid = request.Form("listid")
	if mm_ndelid  = "" then
		response.write "<script language=javascript>alert('没有任何选择!');window.location.href='admin_notice.asp';</script>"
		response.end
	end if
	sql="delete from Ay_Notice where bId in (" & Replace(mm_ndelid, "'", "''") & ")"
	conn.execute sql
	response.write "<script language=javascript>alert('批量删除成功!');window.location.href='admin_notice.asp';</script>"
end sub
%>
<%
private sub ListItem()
%>
<form autocomplete="off" name="form1" id="form1" method="post" action>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>公告管理</b></font>&nbsp;<a href="index.asp?go=body">[返回]</a>&nbsp;<a href="admin_notice.asp">[刷新列表]</a>&nbsp;<a href="admin_notice.asp?go=add">[增加]</a> </td>
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
		<td class="t2" width="40">序号</td>
		<td class="t2" align="center">标题</td>	
		<td class="t2" width="50" align="center">排序</td>
		<td class="t2" align="center">操作</td>
	</tr>
	<%
	dim curpage
	
	sql="select a.* from Ay_Notice a order by a.bId desc"
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
		response.write("<td align='center' colspan=15>找不到任何记录！</td>")
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
		<td align="left"><%=trim(rs("bTitle") & "")%></td>	
		<td align="center"><%=trim(rs("bOrder") & "")%></td>
		<td align="center" width="80px">
		<a href="admin_notice.asp?go=edit&id=<%=trim(rs("bId")&"")%>">修改</a>
		<a href="admin_notice.asp?go=delete&id=<%=trim(rs("bId")&"")%>" onClick="javascript:return confirm('请确认删除操作 ?')">
		删除</a> </td>
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
		<input type="submit" name="mydele" class="button" value="批量删除" onClick="javascript:this.form.action='admin_notice.asp?go=batchdelete';return confirm('请确认删除操作 ?')">
		<input type="checkbox" name="all" id="all" onClick="checkAll(this.checked)"><label for="all">全选</label>
		</td>
		<td align="right" height="30"><%mypage.showpage()%></td>
	</tr>
</table>	
</form>
<%
end sub
%> <%
private sub AddItem()
%>
<form autocomplete="off" name="addform" id="addform" method="post" onSubmit="return checkform(addform)" action="admin_notice.asp?go=saveadd">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>公告管理</b></font>&nbsp;->&nbsp;新增公告&nbsp;&nbsp;←&nbsp;<a href="admin_notice.asp">返回列表</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td><hr size="1"></td>
	</tr>
	<tr bgcolor="#898989">
		<td height="22"><font class="t2">&nbsp;详细资料</font><br>
		</td>
	</tr>
	<tr>
		<td height="10"></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">		
 
	<tr>
		<td class="td1" align="right" valign="middle">内容：</td>
		<td class="td2">
			<input type="text" name="bTitle" id="bTitle" size="50">
		</td>
	</tr> 
	<tr>
		<td class="td1" align="right">排序：</td>
		<td class="td2">
		<input type="text" value="0" name="bOrder" id="bOrder" size="10">&nbsp;<font color="ff4500">*</font>
		</td>
	</tr>
	
	
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td width="150" align="right" height="40"></td>
		<td>
		<input type="submit" class="button" name="submit" value="确认提交">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="reset" class="button" value="重新填写" name="Button">
		</td>
	</tr>				
</table>
</form> 
<script language="Javascript">
	addform.bTitle.focus()
</script>
<%
end sub
%>
<%
private sub EditItem()
    dim mvarbTitle,mvarbOrder,mvarbContent
    
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Notice  where bId=" & request("id") & " order by bId"
	rs.open sql,conn,1,1
	if not rs.eof then
	    mvarbTitle=trim(rs("bTitle")&"")
	    mvarbOrder=trim(rs("bOrder")&"")
	    mvarbContent=trim(rs("bContent")&"")
	end if
	if rs.state<>0 then rs.close
	set rs=nothing
%>
<form autocomplete="off" name="editform" id="editform" method="post" onSubmit="return checkform(editform)" action="admin_notice.asp?go=saveedit&id=<%=request("id")%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>公告管理</b></font>&nbsp;->&nbsp;编辑公告&nbsp;&nbsp;←&nbsp;<a href="admin_notice.asp">返回列表</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td><hr size="1"></td>
	</tr>
	<tr bgcolor="#898989">
		<td height="22"><font class="t2">&nbsp;详细资料</font><br>
		</td>
	</tr>
	<tr>
		<td height="10"></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">	
	<tr>
		<td class="td1" align="right" valign="middle">标题：</td>
		<td class="td2">
			<input type="text" name="bTitle" id="bTitle" size="50" value="<%=mvarbTitle%>">
		</td>
	</tr>
	<tr>
		<td class="td1" align="right">排序：</td>
		<td class="td2">
		<input type="text" value="<%=mvarbOrder%>" name="bOrder" id="bOrder" size="10"><font color="ff4500">*</font>
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">内容：</td>
		<td class="td2" colspan=3>			
			<textarea name="bContent" id="bContent" style="display:none"><%=Server.HtmlEncode(mvarbContent)%></textarea>
			<iframe ID="eWebEditor1" src="../editor/ewebeditor.htm?id=bContent&style=mini" frameborder="0" scrolling="no" width="100%" HEIGHT="300"></iframe>					
		</td>
	</tr>
	
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td width="150" align="right" height="40"></td>
		<td>
		<input type="submit" class="button" name="submit" value="确认提交">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="reset" class="button" value="重新填写" name="Button">
		</td>
	</tr>				
</table>
</form> 
<script language="Javascript">
	editform.bTitle.focus()
</script>
<%
end sub
%>