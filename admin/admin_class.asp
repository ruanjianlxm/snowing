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

<%
'//添加数据库时分类下拉显示
Public Function AddClassBox()
Dim tsql,rss
Response.Write("<select class='cssselect'  name='bParent' id='bParent'>")
tsql="select * from Ay_Class where bParent=0 Order By bOrder"
set rss=Server.CreateObject("Adodb.recordset")
	rss.open tsql,conn,1,1 
If rss.eof And rss.bof Then
Response.Write("<option value='0'>无（作为一级分类）</option>")
Else
	Response.Write("<option value='0'>无（作为一级分类）</option>")
	Do while not(rss.eof)
		Response.Write("<option value='"&rss("bId")&"'>")
		Response.Write(rss("bName")&"</option>")   
	rss.movenext
	Loop
end if
	rss.close
set rss=Nothing
Response.Write("</select>")
End Function
'//编辑数据库时分类下拉显示
Public Function EditClassBox(para_classid)
Dim tsql,rss
Response.Write("<select class='cssselect' name='bParent' id='bParent'>")
tsql="select * from Ay_Class where bParent=0 order by bOrder"
set rss=Server.CreateObject("Adodb.recordset")
	rss.open tsql,conn,1,1 
If rss.eof And rss.bof Then
Response.Write("<option value='0'>无（作为一级分类）</option>")
Else
	Response.Write("<option value='0'")
	if para_classid="0" then
		Response.Write " selected "
	end if
	Response.Write (">无（作为一级分类）</option>")
	do while not(rss.eof)
		Response.Write("<option value='"&rss("bid")&"'")
		    IF para_classid=Cint(Rss("bId")) Then 
		Response.Write(" selected ")
		    End IF
		Response.Write(">")	

		Response.Write(rss("bName")&"</option>")   
		rss.movenext
		Loop
end if
	rss.close
set rss=nothing
Response.Write("</select>")
End Function
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<meta http-equiv="Expires" content="0">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language="Javascript">

function checkform(form)
{			
	var flag=true;
	if(form("bName").value==""){alert("请输入分类名称!");form("bName").focus();return false;}
	if(form("bOrder").value==""){alert("排序只能输入数字!");form("bOrder").focus();return false;}	
	return flag;
}

</script>
</head>

<body topmargin="5" leftmargin="5" bgcolor="#ffffff">
<%
select case request("go")
	Case "add"
		Call AddItem()
	Case "edit"
		Call EditItem()
	case "saveadd"
		call SaveAdd()
	case "saveedit"
		call SaveEdit()
	case "delete"
		call Delete()
	Case Else
		ListItem()
end Select
Call CloseConn()
%>
</body>
</html>

<%
Private Sub ListItem()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>文章分类</b></font>&nbsp;<a href="index.asp?go=body">[返回]</a>&nbsp;<a href="admin_class.asp">[刷新列表]</a>&nbsp;<a href="admin_class.asp?go=add">[增加]</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td><hr size="1"></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="1">
	<tr bgcolor="527c72" style="color:#ffffff" align="center" height="22">				
		<td class="t2" width="40">序号</td>				
		<td class="t2" align="center">分类名称</td>		
		<td class="t2" align="center" width="100">导航位置</td>		
		<td class="t2" align="center" width="60">首页显示</td>
		<td class="t2" width="50" align="center">排序</td>		
		<td class="t2" align="center">备注描述</td>
		<td class="t2" align="center" width="80">操作</td>
	</tr>
	<%
	set rs=server.createobject("adodb.recordset")
	sql="select a.* from Ay_Class a where a.bParent=0 order by a.bId"
	rs.open sql,conn,1,1
	dim i
	i=1
	do while not rs.EOF			
	%>			
	<tr bgcolor="#efefef" height="22">		
		<td align="center" height="22"><%=trim(cstr(i))%></td>				
		<td align="left" valign="middle">&nbsp;<IMG SRC="images/close.gif" WIDTH='9' HEIGHT='16' align="absmiddle">&nbsp;<b><%=trim(rs("bName")&"")%></b>&nbsp;(<%=trim(rs("bId")&"")%>)</td>				
		<td align="center">
		<%
		if trim(rs("bPosition") & "")="0" or trim(rs("bPosition") & "")="" then
			response.write ""
		end if
		if trim(rs("bPosition") & "")="1" then
			response.write "页面头部"
		end if
		if trim(rs("bPosition") & "")="2" then
			response.write "页面尾部"
		end if
		%>
		</td>		
		<td align="center"><input type="checkbox" disabled <%if trim(rs("bShowIndex")&"")="1" then response.write "checked" end if%> /></td>
		<td align="center"><%=trim(rs("bOrder")&"")%></td>
		<td align="left"><%=trim(rs("bRemark")&"")%></td>
		<td align="center"><a href="admin_class.asp?go=edit&id=<%=rs("bId")%>">编辑</a>&nbsp;
		<a href="admin_class.asp?go=delete&id=<%=rs("bId")%>">删除</a>
		</td>
	</tr>
	<%	
		set rss=server.createobject("adodb.recordset")
		sql="select a.* from Ay_Class a where a.bParent=" & trim(rs("bId")&"") & " order by a.bId"
		if rss.state<>0 then rss.Close
		rss.open sql,conn,1,1
		
		do while not rss.eof
			i=i+1
	%>
		<tr bgcolor="#fefefe">
			<td align="center" height="22"><%=trim(cstr(i))%></td>					
			<td align="left">&nbsp;<%=String(2,"　")%><IMG SRC="images/open.gif" WIDTH='9' HEIGHT='16'align="absmiddle">&nbsp;<%=trim(rss("bName")&"")%>&nbsp;(<%=trim(rss("bId")&"")%>)</td>			
			<td align="center">
			<%
			if trim(rss("bPosition") & "")="0" or trim(rss("bPosition") & "")="" then
				response.write ""
			end if
			if trim(rss("bPosition") & "")="1" then
				response.write "页面头部"
			end if
			if trim(rss("bPosition") & "")="2" then
				response.write "页面尾部"
			end if
			%>
			</td>			
			<td align="center"><input type="checkbox" disabled <%if trim(rss("bShowIndex")&"")="1" then response.write "checked" end if%> /></td>
			<td align="center"><%=trim(rss("bOrder")&"")%></td>
			<td align="left"><%=trim(rss("bRemark")&"")%></td>
			<td align="center"><a href="admin_class.asp?go=edit&id=<%=rss("bId")%>">编辑</a>&nbsp;
			<a href="admin_class.asp?go=delete&id=<%=rss("bId")%>">删除</a>
			</td>
		</tr>
	<%
			
			rss.movenext
		loop
		if rss.state<>0 then rss.Close	
		set rss=nothing	
	i=i+1
	rs.MoveNext
loop
if rs.state<>0 then rs.Close				
set rs=nothing		
%>
</table>
<%
End Sub 
%>
<%
private sub SaveAdd()
	dim ad_bName,ad_bRemark,ad_bOrder,ad_bParent,ad_bPosition
	Dim ad_bShowIndex
	ad_bName=request("bName")
	ad_bRemark=request("bRemark")
	ad_bOrder=request("bOrder")
	ad_bParent=request("bParent")
	ad_bPosition=request("bPosition")
	ad_bShowIndex=request("bShowIndex")

	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Class where bParent=" & ad_bParent & " and bName='" & ad_bName & "'"	
	rs.open sql,conn,1,3
	if not rs.eof then
		response.write "<script>alert('该分类已经存在，请重新输入！');window.location.href='admin_class.asp?go=add';</script>"
		response.end
	else
		rs.addnew	
		rs("bName")=ad_bName
		rs("bRemark")=ad_bRemark
		rs("bOrder")=ad_bOrder
		rs("bParent")=ad_bParent
		rs("bPosition")=ad_bPosition	
		rs("bShowIndex")=CLng(ad_bShowIndex+0)
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('分类添加成功，请继续操作！');window.location.href='admin_class.asp';</script>"
		response.end
	end if

end sub
%>
<%
private sub SaveEdit()
	dim ad_bName,ad_bRemark,ad_bOrder,ad_bParent,ad_bPosition
	Dim ad_bShowIndex

	ad_bName=request("bName")
	ad_bRemark=request("bRemark")
	ad_bOrder=request("bOrder")
	ad_bParent=request("bParent")
	ad_bPosition=request("bPosition")	
	ad_bShowIndex=request("bShowIndex")

	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Class where bId=" & request("id")
	rs.open sql,conn,1,3
	if not rs.eof then		
		rs("bName")=ad_bName
		rs("bRemark")=ad_bRemark
		rs("bOrder")=ad_bOrder
		rs("bParent")=ad_bParent
		rs("bPosition")=ad_bPosition		
		rs("bShowIndex")=CLng(ad_bShowIndex+0)
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('分类修改成功，请继续操作！');window.location.href='admin_class.asp';</script>"
		response.end
	end if

end sub
%>
<%
private sub Delete()
	sql="delete from Ay_Class where bId=" & request("id")
	conn.execute sql
	response.write "<script>alert('分类删除成功，请继续操作！');window.location.href='admin_class.asp';</script>"
	response.end
end sub
%>
<%
Private Sub AddItem()
%>
<form autocomplete="off" name="addform" id="addform" method="post" onsubmit="return checkform(addform)" action="admin_class.asp?go=saveadd">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>文章分类</b></font>&nbsp;->&nbsp;新增分类&nbsp;&nbsp;←&nbsp;<a href="admin_class.asp">返回列表</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td><hr size="1"></td></tr>
	<tr bgcolor="#898989"><td height="23"><font class="t2">&nbsp;详细资料</font></td></tr>
	<tr><td height="10"></td></tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
	<tr>
		<td class="td1" align="right" height="5"></td>
		<td class="td2"></td>
	</tr>				
	<tr>
		<td  class="td1" align="right">分类名称：</td>
		<td class="td2">
		<input type="text" value name="bName" id="bName" size="20">&nbsp;<font color="ff4500">*</font>
		</td>
	</tr>
	<tr>
		<td  class="td1" align="right">所属类别：</td>
		<td class="td2">
		<%call AddClassBox()%>
		</td>
	</tr>	
	<tr>
		<td class="td1" align="right" valign="middle">导航位置：</td>
		<td class="td2">
		   <select id="bPosition" name="bPosition">
		   	<option value="0"> </option>
		   	<option value="1">页面头部</option>
		   	<option value="2">页面尾部</option>
		   </select>&nbsp;<font color="ff4500">留空表示不作为导航菜单</font>	  
		</td>
	</tr>
	<tr>
		<td class="td1" align="right">首页显示：</td>
		<td class="td2">
		<input name="bShowIndex" id="bShowIndex" type="checkbox" value="1" checked/>
		</td>
	</tr>
	<tr>
		<td class="td1" align="right">排序：</td>
		<td class="td2">
		<input type="text" value name="bOrder" id="bOrder" size="10">&nbsp;<font color="ff4500">*</font>
		</td>
	</tr>
	<tr>
		<td class="td1" align="right">备注说明：</td>
		<td class="td2">
		<input type="text" value name="bRemark" id="bRemark" size="30">
		</td>
	</tr>
	<tr>
		<td class="td2" bgcolor="ffffff"></td>
		<td class="td2" bgcolor="ffffff" align="left"><br>
		<input type="submit" class="button" name="submit" value="确认提交"> </td>
	</tr>
</table>
</form> 
<script language="Javascript">addform.bName.focus()</script>
<%
End sub
%>
<%
Private Sub EditItem()

	dim mvarbName,mvarbRemark,mvarbOrder,mvarbParent,mvarbPosition
	Dim mvarbShowIndex

	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Class  where bId=" & request("id") & " order by bId"
	rs.open sql,conn,1,1
	if not rs.eof then
		mvarbName=trim(rs("bName")&"")
		mvarbRemark=trim(rs("bRemark")&"")
		mvarbOrder=trim(rs("bOrder")&"")
		mvarbParent=trim(rs("bParent")&"")
		mvarbPosition=trim(rs("bPosition")&"")	
		mvarbShowIndex=trim(rs("bShowIndex")&"")
	end if
	if rs.state<>0 then rs.close
	set rs=nothing
%>
<form autocomplete="off" name="editform" id="editform" method="post" onsubmit="return checkform(editform)" action="admin_class.asp?go=saveedit&id=<%=request("id")%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>文章分类</b></font>&nbsp;->&nbsp;编辑分类&nbsp;&nbsp;←&nbsp;<a href="admin_class.asp">返回列表</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td><hr size="1"></td></tr>
	<tr bgcolor="#898989"><td height="23"><font class="t2">&nbsp;详细资料</font></td></tr>
	<tr><td height="10"></td></tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="1" align="center">
	<tr>
		<td class="td1" align="right" height="5"></td>
		<td class="td2"></td>
	</tr>				
	<tr>
		<td class="td1" align="right">分类名称：</td>
		<td class="td2">
		<input type="text" value="<%=mvarbName%>" name="bName" id="bName" size="20"><font color="ff4500">*</font>
		</td>
	</tr>
	<tr>
		<td  class="td1" align="right">所属类别：</td>
		<td class="td2">					
		<%call EditClassBox(clng(mvarbParent+0))%>
		</td>
	</tr>	
	<tr>
		<td class="td1" align="right" valign="middle">菜单显示：</td>
		<td class="td2">
		   <select id="bPosition" name="bPosition">
		    <option value="0" <%if mvarbPosition="0" then response.write "selected" end if%>> </option>
		   	<option value="1" <%if mvarbPosition="1" then response.write "selected" end if%>>页面头部</option>
		   	<option value="2" <%if mvarbPosition="2" then response.write "selected" end if%>>页面尾部</option>
		   </select>&nbsp;<font color="ff4500">留空表示不作为导航菜单</font>		  
		</td>
	</tr>
	<tr>
		<td class="td1" align="right">首页显示：</td>
		<td class="td2">
		<input name="bShowIndex" id="bShowIndex" type="checkbox" value="1" <% if mvarbShowIndex="1" then response.write "checked" end if %>/>
		</td>
	</tr>
	<tr>
		<td class="td1" align="right">排序：</td>
		<td class="td2">
		<input type="text" value="<%=mvarbOrder%>" name="bOrder" id="bOrder" size="10"><font color="ff4500">*</font>
		</td>
	</tr>
	<tr>
		<td class="td1" align="right">备注说明：</td>
		<td class="td2">
		<input type="text" value="<%=mvarbRemark%>" name="bRemark" id="bRemark" maxlength="20"  size="30">
		</td>
	</tr>
	<tr>
		<td class="td2" bgcolor="ffffff"></td>
		<td class="td2" bgcolor="ffffff" align="left"><br>
		<input id="Hidden1" name="id" type="hidden" value="<%=request("id")%>">
		<input type="submit" class="button" name="submit" value="确认提交"> </td>
	</tr>
</table>
</form> 
<script language="Javascript">editform.bName.focus()</script>
<%
End Sub 
%>