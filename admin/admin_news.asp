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
<%
'//������ݿ�ʱ����������ʾ
public Sub AddClassBox(para_rootid)
Dim tsql,rss
Set rss=Server.Createobject("adodb.recordset")
tsql="Select * from Ay_Class where  bParent=" & clng(para_rootid+0) & " Order by bOrder"
rss.open tsql,conn,1,1
While not rss.eof
	Response.Write "<option value='" & rss("bId") & "'>"
	IF rss("bParent")=0 THEN
		Response.Write "��"
	Else
		Response.Write "����"
	End IF
	Response.Write " " & rss("bName") & "</option>" 
	call AddClassBox(rss("bId"))
	rss.MoveNext
wend
if rss.state<>0 then rss.Close
Set rss=NoThing
End Sub

'//�༭���ݿ�ʱ����������ʾ
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
		Response.Write "��"
	Else
		Response.Write "����"
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
	response.write "<script>alert('������ӳɹ��������������');window.location.href='admin_news.asp';</script>"
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
		response.write "<script>alert('���¸��³ɹ��������������');window.location.href='admin_news.asp';</script>"
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
	response.write "<script language=javascript>alert('����ɾ���ɹ�!');window.location.href='admin_news.asp';</script>"
end Sub
%>
<%
private sub BatchDelete()
    dim mm_ndelid
	mm_ndelid = request.Form("listid")
	if trim(mm_ndelid)  = "" then
		response.write "<script language=javascript>alert('û���κ�ѡ��!');window.location.href='admin_news.asp';</script>"
		response.end
	end if
	sql="delete from Ay_Content where bId in (" & Replace(mm_ndelid, "'", "''") & ")"
	
	conn.execute sql
	response.write "<script language=javascript>alert('����ɾ���ɹ�!');window.location.href='admin_news.asp';</script>"
end sub
%>
<%
private sub DeleteComment()
    dim de_id
	de_id=clng(Request("id"))	
	sql="delete from Ay_Comment where bId=" & de_id
	conn.execute sql
	response.write "<script language=javascript>alert('����ɾ���ɹ�!');window.location.href='admin_news.asp?go=comment';</script>"
end Sub
%>
<%
private sub BatchDeleteComm()
    dim mm_ndelid
	mm_ndelid = request.Form("listid")
	if trim(mm_ndelid)  = "" then
		response.write "<script language=javascript>alert('û���κ�ѡ��!');window.location.href='admin_news.asp?go=comment';</script>"
		response.end
	end if
	sql="delete from Ay_Comment where bId in (" & Replace(mm_ndelid, "'", "''") & ")"
	
	conn.execute sql
	response.write "<script language=javascript>alert('����ɾ���ɹ�!');window.location.href='admin_news.asp?go=comment';</script>"
end sub
%>

<%
private sub ListItem()
%>
<form autocomplete="off" name="form1" id="form1" method="post" action>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>�����ڵ�λ�ã�<font color="DarkSlateGray" style="font-size:12px"><b>���¹���</b></font>&nbsp;<a href="index.asp?go=body">[����]</a>&nbsp;<a href="admin_news.asp">[ˢ���б�]</a>&nbsp;<a href="admin_news.asp?go=add">[���]</a> </td>
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
		<td class="t2" align="left">&nbsp;����</td>
		<td class="t2" align="center">���</td>			
		<td class="t2" align="center">������</td>		
		<td class="t2" align="center">����</td>	
		<td class="t2" align="center" width="40px">�ö�</td>
		<td class="t2" align="center" width="40px">�Ƽ�</td>
		<td class="t2" align="center" width="40px">����</td>
		<td class="t2" align="center" width="40px">���</td>
		<td class="t2" align="center">����</td>
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
		<a href="admin_news.asp?go=comment&id=<%=trim(rs("bId")&"")%>">����</a>
		<a href="admin_news.asp?go=edit&id=<%=trim(rs("bId")&"")%>">�޸�</a>
		<a href="admin_news.asp?go=delete&id=<%=trim(rs("bId")&"")%>" onClick="javascript:return confirm('��ȷ��ɾ������ ?')">
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
		<input type="submit" name="mydele" class="button" value="����ɾ��" onClick="javascript:this.form.action='admin_news.asp?go=batchdelete';return confirm('��ȷ��ɾ������ ?')">
		<input type="checkbox" name="all" id="all" onClick="checkAll(this.checked)"><label for="all">ȫѡ</label>
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
		<td>�����ڵ�λ�ã�<font color="DarkSlateGray" style="font-size:12px"><b>��������</b></font>&nbsp;<a href="index.asp?go=body">[����]</a>&nbsp;<a href="admin_news.asp?go=comment">[ˢ���б�]</a></td>
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
		<td class="t2" align="left">����</td>
		<td class="t2" align="left">&nbsp;��������</td>
		<td class="t2" align="center">������</td>	
		<td class="t2" align="center">��Դ��ַ</td>			
		<td class="t2" align="center">����</td>
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
		<td align="left"><%=trim(rs("bTitle") & "")%></td>
		<td align="left"><%=left(trim(rs("bContent") & ""),30)%></td>
		<td align="center"><%=trim(rs("bAddUser") & "")%></td>	
		<td align="center"><%=trim(rs("bIPAddress") & "")%></td>
		<td align="center" width="80px">		
		<a href="admin_news.asp?go=commentdelete&id=<%=trim(rs("bId")&"")%>" onClick="javascript:return confirm('��ȷ��ɾ������ ?')">
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
		<input type="submit" name="mydele" class="button" value="����ɾ��" onClick="javascript:this.form.action='admin_news.asp?go=batchcomm';return confirm('��ȷ��ɾ������ ?')">
		<input type="checkbox" name="all" id="all" onClick="checkAll(this.checked)"><label for="all">ȫѡ</label>
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
		<td>�����ڵ�λ�ã�<font color="DarkSlateGray" style="font-size:12px"><b>���¹���</b></font>&nbsp;->&nbsp;��������&nbsp;&nbsp;��&nbsp;<a href="admin_news.asp">�����б�</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td><hr size="1"></td></tr>
	<tr bgcolor="#898989"><td height="23"><font class="t2">&nbsp;��ϸ����</font></td></tr>
	<tr><td height="10"></td></tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">		 					
	<tr>
		<td align="right" class="td1" valign="middle">�������ࣺ</td>
		<td class="td2">
		<select name="bClassID" id="bClassID">
		<option value="0">--ѡ�����--</option>
		<%call AddClassBox(0)%>
		</select>&nbsp;<font color="#ff0000">*</font>
		</td>			
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle">���⣺</td>
		<td class="td2">
		    <input type="text" name="bTitle" id="bTitle" size="60">&nbsp;<font color="#ff0000">*</font>
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">����ͼƬ��</td>
		<td class="td2">
		<input type="text" class="input" id="bPic" name="bPic"  style="width:250px;"  />
		<input type="checkbox" onClick="display('upload');" id="box"/><label for='box'>�ϴ�ͼƬ</label>&nbsp;<font color ="#ff0000">��������ÿգ���ϵͳ�Զ���ȡ�����еĵ�һ��ͼƬ</font>
		<br>
		<div id="upload" style="display:none;" class="td2">			
		<iframe src="upload.asp?go=pic" frameborder='0' style='height:22px;width:100%;' scrolling='no'></iframe>
		</div>
		</td>
	</tr>
	<tr>
		<td width="15%" class="td1" align="right">���ߣ�</td>
		<td width="85%" class="td2">
		<input type="text" name="bWriter" id="bWriter" size="20" value="" maxlength="50">&nbsp;
		<button class="button" onClick="bWriter.value='<%=session("UserName")%>'"><%=session("UserName")%></button>&nbsp;
		<button class="button" onClick="bWriter.value='δ֪'">δ֪</button>
		</td>
 	</tr> 
  	<tr>
		<td align="right" class="td1">��Դ��</td>
		<td class="td2">
		<input type="text" id="bCopyRight" name="bCopyRight" size="20" value="" maxlength="50">&nbsp;
		<button class="button" onClick="bCopyRight.value='��վ'">��վ</button>&nbsp;
		<button class="button" onClick="bCopyRight.value='δ֪'">δ֪</button>&nbsp;
		<button class="button" onClick="bCopyRight.value='������'">������</button>
		</td>
  	</tr> 
	<tr>
		<td align="right" class="td1">�ö��Ƽ���</td>
		<td class="td2">
		<input name="bIsTop" id="bIsTop" type="checkbox" value="1" checked/><label for="bIsTop">�ö�</label>&nbsp;
		<input name="bIsBest" id="bIsBest" type="checkbox" value="1" checked/><label for="bIsBest">�Ƽ�</label>&nbsp;
		<input name="bIsReply" id="bIsReply" type="checkbox" value="1"/><label for="bIsReply">����</label>&nbsp;
		<input name="bIsPass" id="bIsPass" type="checkbox" value="1" checked/><label for="bIsPass">���ͨ��</label>
		</td>
	</tr>	
	<tr>
		<td align="right" class="td1" valign="top">���ݣ�</td>
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
		<input type="button" class="button" name="submit1" value="ȷ���ύ" onClick="if(checkform(addform)){this.form.action='admin_news.asp?go=saveadd';this.form.submit();}">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="reset" class="button" value="������д" name="Button">
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
		<td>�����ڵ�λ�ã�<font color="DarkSlateGray" style="font-size:12px"><b>���¹���</b></font>&nbsp;->&nbsp;�༭����&nbsp;&nbsp;��&nbsp;<a href="admin_news.asp">�����б�</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td><hr size="1"></td></tr>
	<tr bgcolor="#898989"><td height="23"><font class="t2">&nbsp;��ϸ����</font></td></tr>
	<tr><td height="10"></td></tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">		 					
	<tr>
		<td align="right" class="td1" valign="middle">�������ࣺ</td>
		<td class="td2">
		<select name="bClassID" id="bClassID">
		<option value="0">--ѡ�����--</option>
		<%call EditClassBox(0,CLng(mvarbClassID+0))%>
		</select>&nbsp;<font color="#ff0000">*</font>
		</td>			
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle">���⣺</td>
		<td class="td2">
		    <input type="text" name="bTitle" id="bTitle" value="<%=mvarbTitle%>" size="60">&nbsp;<font color="#ff0000">*</font>
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">����ͼƬ��</td>
		<td class="td2">
		<input type="text" class="input" id="bPic" name="bPic" value="<%=mvarbPic%>"  style="width:250px;"  />
		<input type="checkbox" onClick="display('upload');" id="box"/><label for='box'>�ϴ�ͼƬ</label>&nbsp;<font color ="#ff0000">��������ÿգ���ϵͳ�Զ���ȡ�����еĵ�һ��ͼƬ</font>
		<br>
		<div id="upload" style="display:none;" class="td2">			
		<iframe src="upload.asp?go=pic" frameborder='0' style='height:22px;width:100%;' scrolling='no'></iframe>
		</div>
		</td>
	</tr>	
	<tr>
		<td width="15%" class="td1" align="right">���ߣ�</td>
		<td width="85%" class="td2">
		<input type="text" name="bWriter" id="bWriter" size="20" value="<%=mvarbWriter%>" maxlength="50">&nbsp;
		<button class="button" onClick="bWriter.value='<%=session("UserName")%>'"><%=session("UserName")%></button>&nbsp;
		<button class="button" onClick="bWriter.value='δ֪'">δ֪</button>
		</td>
 	</tr> 
  	<tr>
		<td align="right" class="td1">��Դ��</td>
		<td class="td2">
		<input type="text" id="bCopyRight" name="bCopyRight" size="20" value="<%=mvarbCopyRight%>" maxlength="50">&nbsp;
		<button class="button" onClick="bCopyRight.value='��վ'">��վ</button>&nbsp;
		<button class="button" onClick="bCopyRight.value='δ֪'">δ֪</button>&nbsp;
		<button class="button" onClick="bCopyRight.value='������'">������</button>
		</td>
  	</tr> 
	<tr>
		<td align="right" class="td1">�ö��Ƽ���</td>
		<td class="td2">
		<input name="bIsTop" id="bIsTop" type="checkbox" value="1" <% if mvarbIsTop="1" then response.write "checked" end if %>/><label for="bIsTop">�ö�</label>&nbsp;
		<input name="bIsBest" id="bIsBest" type="checkbox" value="1" <% if mvarbIsBest="1" then response.write "checked" end if %>/><label for="bIsBest">�Ƽ�</label>&nbsp;
		<input name="bIsReply" id="bIsReply" type="checkbox" value="1" <% if mvarbIsReply="1" then response.write "checked" end if %>/><label for="bIsReply">����</label>&nbsp;
		<input name="bIsPass" id="bIsPass" type="checkbox" value="1" <% if mvarbIsPass="1" then response.write "checked" end if %>/><label for="bIsPass">���ͨ��</label>
		</td>
	</tr>	
	<tr>
		<td align="right" class="td1" valign="top">���ݣ�</td>
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
		<input type="button" class="button" name="submit1" value="ȷ���ύ" onClick="if(checkform(editform)){this.form.action='admin_news.asp?go=saveedit&id=<%=request("id")%>';this.form.submit();}">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="reset" class="button" value="������д" name="Button">
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