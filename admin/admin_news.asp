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
<%
'//添加数据库时分类下拉显示
public Sub AddClassBox(para_rootid)
Dim tsql,rss
Set rss=Server.Createobject("adodb.recordset")
tsql="Select * from Ay_Class where  bParent=" & clng(para_rootid+0) & " Order by bOrder"
rss.open tsql,conn,1,1
While not rss.eof
	Response.Write "<option value='" & rss("bId") & "'>"
	IF rss("bParent")=0 THEN
		Response.Write "┿"
	Else
		Response.Write "　├"
	End IF
	Response.Write " " & rss("bName") & "</option>" 
	call AddClassBox(rss("bId"))
	rss.MoveNext
wend
if rss.state<>0 then rss.Close
Set rss=NoThing
End Sub

'//编辑数据库时分类下拉显示
Public Function EditClassBox(para_rootid,para_classid)
Dim tsql,rss
Set rss=Server.Createobject("adodb.recordset")
tsql="Select * from Ay_Class where bParent=" & clng(para_rootid+0) & " Order by bOrder"
rss.open tsql,conn,1,1
While not rss.eof	
	Response.Write("<option value='"&rss("bid")&"'")
	IF para_classid=Cint(Rss("bId")) Then 
		Response.Write("selected='selected'")
	End IF
	Response.Write(">")
		
	IF rss("bParent")=0 THEN
		Response.Write "┿"
	Else
		Response.Write "　├"
	End IF
	Response.Write " " & rss("bName") & "</option>" 
	call EditClassBox(rss("bId"),para_classid)
	rss.MoveNext
wend
if rss.state<>0 then rss.Close
Set rss=NoThing
End Function
%>

<html>
<head>
<title>网站资料</title>
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
		if(form("bClassID").value=="0"){alert("请选择分类!");form("bClassID").focus();return false;}	
		if(form("bTitle").value==""){alert("标题名称不能为空!");form("bTitle").focus();return false;}		
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
	case "comment"
		call CommentList()
	case "commentdelete"
		call DeleteComment()
	case "batchcomm"
		call BatchDeleteComm()
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
	dim ad_bClassID,ad_bTitle,ad_bWriter,ad_bCopyRight	
	dim ad_bContent,ad_bPic,ad_bIsTop,ad_bIsBest,ad_bIsPass,ad_bIsReply	
	
	ad_bClassID=request("bClassID")
	ad_bTitle=request("bTitle")		
	ad_bWriter=request("bWriter")	
	ad_bCopyRight=request("bCopyRight")	
	ad_bContent=request("bContent")	
	ad_bPic=request("bPic")	
	if trim(ad_bPic&"")="" then
		ad_bPic=GetEditorImg(ad_bContent)
	end if
	ad_bIsTop=request("bIsTop")
	ad_bIsBest=request("bIsBest")
	ad_bIsPass=request("bIsPass")
	ad_bIsReply=request("bIsReply")	
	
	set rs=server.createobject("adodb.recordset")
	sql="select top 1 * from Ay_Content"
	rs.open sql,conn,1,3	
	rs.addnew	
	rs("bClassID")=ad_bClassID		
	rs("bTitle")=ad_bTitle	
	rs("bWriter")=ad_bWriter	
	rs("bCopyRight")=ad_bCopyRight	
	rs("bContent")=ad_bContent	
	rs("bPic")=ad_bPic	
	rs("bIsTop")=clng(ad_bIsTop+0)
	rs("bIsBest")=clng(ad_bIsBest+0)
	rs("bIsPass")=clng(ad_bIsPass+0)
	rs("bIsReply")=clng(ad_bIsReply+0)	
	
	rs("bAddTime")=now
	rs("bAddUser")=session("admin")		
	rs.update
	if rs.state<>0 then rs.close
	set rs=nothing
	response.write "<script>alert('文章添加成功，请继续操作！');window.location.href='admin_news.asp';</script>"
	response.end
	
end sub
%>
<%
private sub UpdateItem()
    dim ed_bClassID,ed_bTitle,ed_bWriter,ed_bCopyRight	
	dim ed_bContent,ed_bPic,ed_bIsTop,ed_bIsBest,ed_bIsPass,ed_bIsReply
		
	ed_bClassID=request("bClassID")		
	ed_bTitle=request("bTitle")		
	ed_bWriter=request("bWriter")	
	ed_bCopyRight=request("bCopyRight")	
	ed_bContent=request("bContent")	
	ed_bPic=request("bPic")	
	if trim(ed_bPic&"")="" then
		ed_bPic=GetEditorImg(ed_bContent)
	end if
	ed_bIsTop=request("bIsTop")
	ed_bIsBest=request("bIsBest")
	ed_bIsPass=request("bIsPass")
	ed_bIsReply=request("bIsReply")	
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Content where bId=" & request("id") 
	rs.open sql,conn,1,3
	if not rs.eof then	
		rs("bClassID")=ed_bClassID		
		rs("bTitle")=ed_bTitle	
		rs("bWriter")=ed_bWriter	
		rs("bCopyRight")=ed_bCopyRight	
		rs("bContent")=ed_bContent	
		rs("bPic")=ed_bPic	
		rs("bIsTop")=clng(ed_bIsTop+0)
		rs("bIsBest")=clng(ed_bIsBest+0)
		rs("bIsPass")=clng(ed_bIsPass+0)
		rs("bIsReply")=clng(ed_bIsReply+0)	
		
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('文章更新成功，请继续操作！');window.location.href='admin_news.asp';</script>"
		response.end
	end if
end sub
%>
<%
private sub DeleteItem()
    dim de_id
	de_id=clng(Request("id"))	
	sql="delete from Ay_Content where bId=" & de_id
	conn.execute sql
	response.write "<script language=javascript>alert('文章删除成功!');window.location.href='admin_news.asp';</script>"
end Sub
%>
<%
private sub BatchDelete()
    dim mm_ndelid
	mm_ndelid = request.Form("listid")
	if trim(mm_ndelid)  = "" then
		response.write "<script language=javascript>alert('没有任何选择!');window.location.href='admin_news.asp';</script>"
		response.end
	end if
	sql="delete from Ay_Content where bId in (" & Replace(mm_ndelid, "'", "''") & ")"
	
	conn.execute sql
	response.write "<script language=javascript>alert('批量删除成功!');window.location.href='admin_news.asp';</script>"
end sub
%>
<%
private sub DeleteComment()
    dim de_id
	de_id=clng(Request("id"))	
	sql="delete from Ay_Comment where bId=" & de_id
	conn.execute sql
	response.write "<script language=javascript>alert('评论删除成功!');window.location.href='admin_news.asp?go=comment';</script>"
end Sub
%>
<%
private sub BatchDeleteComm()
    dim mm_ndelid
	mm_ndelid = request.Form("listid")
	if trim(mm_ndelid)  = "" then
		response.write "<script language=javascript>alert('没有任何选择!');window.location.href='admin_news.asp?go=comment';</script>"
		response.end
	end if
	sql="delete from Ay_Comment where bId in (" & Replace(mm_ndelid, "'", "''") & ")"
	
	conn.execute sql
	response.write "<script language=javascript>alert('批量删除成功!');window.location.href='admin_news.asp?go=comment';</script>"
end sub
%>

<%
private sub ListItem()
%>
<form autocomplete="off" name="form1" id="form1" method="post" action>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>文章管理</b></font>&nbsp;<a href="index.asp?go=body">[返回]</a>&nbsp;<a href="admin_news.asp">[刷新列表]</a>&nbsp;<a href="admin_news.asp?go=add">[添加]</a> </td>
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
		<td class="t2" align="left">&nbsp;标题</td>
		<td class="t2" align="center">类别</td>			
		<td class="t2" align="center">评论数</td>		
		<td class="t2" align="center">人气</td>	
		<td class="t2" align="center" width="40px">置顶</td>
		<td class="t2" align="center" width="40px">推荐</td>
		<td class="t2" align="center" width="40px">评论</td>
		<td class="t2" align="center" width="40px">审核</td>
		<td class="t2" align="center">操作</td>
	</tr>
	<%
	dim curpage
	
	sql="select *  from Ay_Content_v a order by a.bId desc"
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
		response.write("<td align='center' colspan=15>找不到任何记录！</td>")
		response.write("</tr>")
	end if
	for i=1 to mypage.pagesize
		if rs.eof or rs.bof then exit for		
			if (i mod 2)=0 then
				Response.Write("<tr height='23' bgcolor='#fefefe' ")
			else
				Response.Write("<tr height='23' bgcolor='#efefef' ")
			end if	
			if trim(rs("bIsPass")&"")<>"1" then
				response.write "style='color:#ff0000;'"
			end if
			response.write ">"
	%>			
		<td width="10" height="20" align="center">
		<input name="listid" type="checkbox" id="listid" value="<%=(rs("bId"))%>">
		</td>
		<td width="40" align="center"><%=trim(cstr(i + (curpage -1) * mypage.pagesize))%></td>	
		
		<td align="left"><%=trim(rs("bTitle") & "")%></td>
		<td align="center"><%=trim(rs("bClassName") & "")%></td>		
		<td align="center"><%=trim(rs("bCommentCount") & "")%></td>		
		<td align="center"><%=trim(rs("bClick") & "")%></td>	
		<td align="center">
		<input type="checkbox" disabled <%if rs("bIsTop")=1 then response.write "checked" end if%> /></td>
		<td align="center">
		<input type="checkbox" disabled <%if rs("bisbest")=1 then response.write "checked" end if%> /></td>
		<td align="center">
		<input type="checkbox" disabled <%if rs("bIsReply")=1 then response.write "checked" end if%> /></td>
		<td align="center">
		<input type="checkbox" disabled <%if rs("bispass")=1 then response.write "checked" end if%> /></td>
		<td align="center" width="100px">
		<a href="admin_news.asp?go=comment&id=<%=trim(rs("bId")&"")%>">评论</a>
		<a href="admin_news.asp?go=edit&id=<%=trim(rs("bId")&"")%>">修改</a>
		<a href="admin_news.asp?go=delete&id=<%=trim(rs("bId")&"")%>" onClick="javascript:return confirm('请确认删除操作 ?')">
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
		<input type="submit" name="mydele" class="button" value="批量删除" onClick="javascript:this.form.action='admin_news.asp?go=batchdelete';return confirm('请确认删除操作 ?')">
		<input type="checkbox" name="all" id="all" onClick="checkAll(this.checked)"><label for="all">全选</label>
		</td>
		<td align="right" height="30"><%mypage.showpage()%></td>
	</tr>
</table>
</form>
<%
end sub
%> 

<%
private sub CommentList()
%>
<form autocomplete="off" name="comform" id="comform" method="post" action>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>评论留言</b></font>&nbsp;<a href="index.asp?go=body">[返回]</a>&nbsp;<a href="admin_news.asp?go=comment">[刷新列表]</a></td>
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
		<td class="t2" align="left">标题</td>
		<td class="t2" align="left">&nbsp;评论留言</td>
		<td class="t2" align="center">评论者</td>	
		<td class="t2" align="center">来源地址</td>			
		<td class="t2" align="center">操作</td>
	</tr>
	<%
	dim curpage
	
	sql="select *  from Ay_Comment_v a where 1=1 "
	if request("id")<>"" then
		sql=sql & " and bArtID=" & request("id")
	end if
	sql=sql & " order by a.bId desc,a.bOrder"
	Set rs=Server.CreateObject("Adodb.Recordset")
	Set mypage=new xdownpage
	mypage.getconn=conn
	mypage.getsql=sql
	mypage.pagesize=12
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
		<td align="left"><%=left(trim(rs("bContent") & ""),30)%></td>
		<td align="center"><%=trim(rs("bAddUser") & "")%></td>	
		<td align="center"><%=trim(rs("bIPAddress") & "")%></td>
		<td align="center" width="80px">		
		<a href="admin_news.asp?go=commentdelete&id=<%=trim(rs("bId")&"")%>" onClick="javascript:return confirm('请确认删除操作 ?')">
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
		<input type="submit" name="mydele" class="button" value="批量删除" onClick="javascript:this.form.action='admin_news.asp?go=batchcomm';return confirm('请确认删除操作 ?')">
		<input type="checkbox" name="all" id="all" onClick="checkAll(this.checked)"><label for="all">全选</label>
		</td>
		<td align="right" height="30"><%mypage.showpage()%></td>
	</tr>
</table>
</form>
<%
end sub
%> 

<%
private sub AddItem()
%>
<form autocomplete="off" name="addform" id="addform" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>文章管理</b></font>&nbsp;->&nbsp;新增文章&nbsp;&nbsp;←&nbsp;<a href="admin_news.asp">返回列表</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td><hr size="1"></td></tr>
	<tr bgcolor="#898989"><td height="23"><font class="t2">&nbsp;详细资料</font></td></tr>
	<tr><td height="10"></td></tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">		 					
	<tr>
		<td align="right" class="td1" valign="middle">所属分类：</td>
		<td class="td2">
		<select name="bClassID" id="bClassID">
		<option value="0">--选择分类--</option>
		<%call AddClassBox(0)%>
		</select>&nbsp;<font color="#ff0000">*</font>
		</td>			
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle">标题：</td>
		<td class="td2">
		    <input type="text" name="bTitle" id="bTitle" size="60">&nbsp;<font color="#ff0000">*</font>
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">内容图片：</td>
		<td class="td2">
		<input type="text" class="input" id="bPic" name="bPic"  style="width:250px;"  />
		<input type="checkbox" onClick="display('upload');" id="box"/><label for='box'>上传图片</label>&nbsp;<font color ="#ff0000">如果这里置空，则系统自动提取内容中的第一张图片</font>
		<br>
		<div id="upload" style="display:none;" class="td2">			
		<iframe src="upload.asp?go=pic" frameborder='0' style='height:22px;width:100%;' scrolling='no'></iframe>
		</div>
		</td>
	</tr>
	<tr>
		<td width="15%" class="td1" align="right">作者：</td>
		<td width="85%" class="td2">
		<input type="text" name="bWriter" id="bWriter" size="20" value="" maxlength="50">&nbsp;
		<button class="button" onClick="bWriter.value='<%=session("UserName")%>'"><%=session("UserName")%></button>&nbsp;
		<button class="button" onClick="bWriter.value='未知'">未知</button>
		</td>
 	</tr> 
  	<tr>
		<td align="right" class="td1">来源：</td>
		<td class="td2">
		<input type="text" id="bCopyRight" name="bCopyRight" size="20" value="" maxlength="50">&nbsp;
		<button class="button" onClick="bCopyRight.value='本站'">本站</button>&nbsp;
		<button class="button" onClick="bCopyRight.value='未知'">未知</button>&nbsp;
		<button class="button" onClick="bCopyRight.value='互联网'">互联网</button>
		</td>
  	</tr> 
	<tr>
		<td align="right" class="td1">置顶推荐：</td>
		<td class="td2">
		<input name="bIsTop" id="bIsTop" type="checkbox" value="1" checked/><label for="bIsTop">置顶</label>&nbsp;
		<input name="bIsBest" id="bIsBest" type="checkbox" value="1" checked/><label for="bIsBest">推荐</label>&nbsp;
		<input name="bIsReply" id="bIsReply" type="checkbox" value="1"/><label for="bIsReply">评论</label>&nbsp;
		<input name="bIsPass" id="bIsPass" type="checkbox" value="1" checked/><label for="bIsPass">审核通过</label>
		</td>
	</tr>	
	<tr>
		<td align="right" class="td1" valign="top">内容：</td>
		<td class="td2" colspan=3>			
			<textarea name="bContent" id="bContent" style="display:none"></textarea>
			<iframe ID="eWebEditor1" src="../editor/ewebeditor.asp?id=bContent&style=s_coolblue" frameborder="0" scrolling="no" width="100%" HEIGHT="500"></iframe>					
		</td>
	</tr>
 	
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td width="150" align="right" height="40"></td>
		<td>
		<input type="button" class="button" name="submit1" value="确认提交" onClick="if(checkform(addform)){this.form.action='admin_news.asp?go=saveadd';this.form.submit();}">&nbsp;&nbsp;&nbsp;&nbsp;
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
	dim mvarbClassID,mvarbTitle,mvarbPic,mvarbWriter,mvarbCopyRight
	dim mvarbIsTop,mvarbIsBest,mvarbIsPass,mvarbIsReply
	dim mvarbContent
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Content  where bId=" & request("id") & " order by bId"
	rs.open sql,conn,1,1
	if not rs.eof Then	
	    mvarbClassID=trim(rs("bClassID")&"")		   
	    mvarbTitle=trim(rs("bTitle")&"")	    
	    mvarbPic=trim(rs("bPic")&"")
	    mvarbWriter=trim(rs("bWriter")&"")
	    mvarbCopyRight=trim(rs("bCopyRight")&"") 
	    mvarbContent=trim(rs("bContent")&"")	    
	    mvarbIsTop=trim(rs("bIsTop")&"")
		mvarbIsBest=trim(rs("bIsBest")&"")
	    mvarbIsPass=trim(rs("bIsPass")&"")
	    mvarbIsReply=trim(rs("bIsReply")&"")		
	end if
	if rs.state<>0 then rs.close
	set rs=nothing
%>
<form autocomplete="off" name="editform" id="editform" method="post">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>文章管理</b></font>&nbsp;->&nbsp;编辑文章&nbsp;&nbsp;←&nbsp;<a href="admin_news.asp">返回列表</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td><hr size="1"></td></tr>
	<tr bgcolor="#898989"><td height="23"><font class="t2">&nbsp;详细资料</font></td></tr>
	<tr><td height="10"></td></tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">		 					
	<tr>
		<td align="right" class="td1" valign="middle">所属分类：</td>
		<td class="td2">
		<select name="bClassID" id="bClassID">
		<option value="0">--选择分类--</option>
		<%call EditClassBox(0,CLng(mvarbClassID+0))%>
		</select>&nbsp;<font color="#ff0000">*</font>
		</td>			
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle">标题：</td>
		<td class="td2">
		    <input type="text" name="bTitle" id="bTitle" value="<%=mvarbTitle%>" size="60">&nbsp;<font color="#ff0000">*</font>
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">内容图片：</td>
		<td class="td2">
		<input type="text" class="input" id="bPic" name="bPic" value="<%=mvarbPic%>"  style="width:250px;"  />
		<input type="checkbox" onClick="display('upload');" id="box"/><label for='box'>上传图片</label>&nbsp;<font color ="#ff0000">如果这里置空，则系统自动提取内容中的第一张图片</font>
		<br>
		<div id="upload" style="display:none;" class="td2">			
		<iframe src="upload.asp?go=pic" frameborder='0' style='height:22px;width:100%;' scrolling='no'></iframe>
		</div>
		</td>
	</tr>	
	<tr>
		<td width="15%" class="td1" align="right">作者：</td>
		<td width="85%" class="td2">
		<input type="text" name="bWriter" id="bWriter" size="20" value="<%=mvarbWriter%>" maxlength="50">&nbsp;
		<button class="button" onClick="bWriter.value='<%=session("UserName")%>'"><%=session("UserName")%></button>&nbsp;
		<button class="button" onClick="bWriter.value='未知'">未知</button>
		</td>
 	</tr> 
  	<tr>
		<td align="right" class="td1">来源：</td>
		<td class="td2">
		<input type="text" id="bCopyRight" name="bCopyRight" size="20" value="<%=mvarbCopyRight%>" maxlength="50">&nbsp;
		<button class="button" onClick="bCopyRight.value='本站'">本站</button>&nbsp;
		<button class="button" onClick="bCopyRight.value='未知'">未知</button>&nbsp;
		<button class="button" onClick="bCopyRight.value='互联网'">互联网</button>
		</td>
  	</tr> 
	<tr>
		<td align="right" class="td1">置顶推荐：</td>
		<td class="td2">
		<input name="bIsTop" id="bIsTop" type="checkbox" value="1" <% if mvarbIsTop="1" then response.write "checked" end if %>/><label for="bIsTop">置顶</label>&nbsp;
		<input name="bIsBest" id="bIsBest" type="checkbox" value="1" <% if mvarbIsBest="1" then response.write "checked" end if %>/><label for="bIsBest">推荐</label>&nbsp;
		<input name="bIsReply" id="bIsReply" type="checkbox" value="1" <% if mvarbIsReply="1" then response.write "checked" end if %>/><label for="bIsReply">评论</label>&nbsp;
		<input name="bIsPass" id="bIsPass" type="checkbox" value="1" <% if mvarbIsPass="1" then response.write "checked" end if %>/><label for="bIsPass">审核通过</label>
		</td>
	</tr>	
	<tr>
		<td align="right" class="td1" valign="top">内容：</td>
		<td class="td2">			
			<textarea name="bContent" id="bContent" style="display:none"><%=Server.HtmlEncode(mvarbContent)%></textarea>
			<iframe ID="eWebEditor1" src="../editor/ewebeditor.asp?id=bContent&style=s_coolblue" frameborder="0" scrolling="no" width="100%" HEIGHT="500"></iframe>					
		</td>
	</tr>
 	
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td width="150" align="right" height="40"></td>
		<td>
		<input type="button" class="button" name="submit1" value="确认提交" onClick="if(checkform(editform)){this.form.action='admin_news.asp?go=saveedit&id=<%=request("id")%>';this.form.submit();}">&nbsp;&nbsp;&nbsp;&nbsp;
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