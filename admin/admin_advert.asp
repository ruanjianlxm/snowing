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
		if(form("bTitle").value==""){alert("广告名称不能为空!");form("bTitle").focus();return false;}	
		if(form("bType").value=="0"){alert("请选择广告类型!");form("bType").focus();return false;}	
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
	Case "buildjs"
		Call BuildJS()
	case else
		call ListItem()
end select
call CloseConn()
%>

</body>

</html>
<%
private sub SaveItem()
    dim ad_bTitle,ad_bType,ad_bPic,ad_bScript
    dim ad_bUrl,ad_bOpenMode,ad_bPicWidth,ad_bPicHeight,ad_bRemark
    
	ad_bTitle=request("bTitle")
	ad_bType=request("bType")
	ad_bPic=request("bPic")
	ad_bScript=request("bScript")
	ad_bUrl=request("bUrl")
	ad_bOpenMode=request("bOpenMode")
	ad_bPicWidth=request("bPicWidth")
	ad_bPicHeight=request("bPicHeight")
	ad_bRemark=request("bRemark")
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Advert where bTitle='" & ad_bTitle & "'"	
	rs.open sql,conn,1,3
	if not rs.eof then
		response.write "<script>alert('该广告标题已经存在，请重新输入！');window.location.href='admin_advert.asp?go=add';</script>"
		response.end
	else
		rs.addnew	
		rs("bTitle")=ad_bTitle
		rs("bType")=ad_bType
		rs("bPic")=ad_bPic
		rs("bScript")=HTMLEncode(ad_bScript)
		rs("bUrl")=ad_bUrl
		rs("bOpenMode")=ad_bOpenMode
		rs("bPicWidth")=ad_bPicWidth
		rs("bPicHeight")=ad_bPicHeight
		rs("bRemark")=ad_bRemark
		rs("bAddTime")=now
		rs("bAddUser")=session("admin")		
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('广告添加成功，请继续操作！');window.location.href='admin_advert.asp';</script>"
		response.end
	end if
end sub
%> <%
private sub UpdateItem()
    dim ed_bTitle,ed_bType,ed_bPic,ed_bScript
    dim ed_bUrl,ed_bOpenMode,ed_bPicWidth,ed_bPicHeight,ed_bRemark
    
	ed_bTitle=request("bTitle")
	ed_bType=request("bType")
	ed_bPic=request("bPic")
	ed_bScript=request("bScript")
	ed_bUrl=request("bUrl")
	ed_bOpenMode=request("bOpenMode")
	ed_bPicWidth=request("bPicWidth")
	ed_bPicHeight=request("bPicHeight")
	ed_bRemark=request("bRemark")
	
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Advert where bId=" & request("id") 
	rs.open sql,conn,1,3
	if not rs.eof then	
		rs("bTitle")=ed_bTitle
		rs("bType")=ed_bType
		rs("bPic")=ed_bPic
		rs("bScript")=HTMLEncode(ed_bScript)
		rs("bUrl")=ed_bUrl
		rs("bOpenMode")=ed_bOpenMode
		rs("bPicWidth")=ed_bPicWidth
		rs("bPicHeight")=ed_bPicHeight
		rs("bRemark")=ed_bRemark
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('广告更新成功，请继续操作！');window.location.href='admin_advert.asp';</script>"
		response.end
	end if
end sub
%>
<% 
Private Sub BuildJS()
	Dim goaler
	goaler=""
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Advert where bId=" & request("id") 
	rs.open sql,conn,1,3
	if rs("bType")="1" then
		goaler = goaler + "<a href="""& rs("bUrl")&""" target="""& rs("bOpenMode")&"""><img src=""../"& rs("bPic")&""" width="""& rs("bPicWidth")&""" height="""& rs("bPicHeight")&"""  title="""& rs("bTitle")&"""></a>"  
	elseif rs("bType")="2" then
		goaler = goaler + "<embed src=""../"&rs("bPic")&""" quality=""height"" type=""application/x-shockwave-flash""  width="""& rs("bPicWidth")&""" height="""& rs("bPicHeight")&""" ></embed>" 
	else
		goaler = goaler + ""& HTMLCode(rs("bScript"))&"" 
	end if
'生成JS文件
	goaler = "" + goaler + ""
	goaler = "document.write('" & goaler & "')"
	FolderPath = Server.MapPath("../")
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Set fout = fso.CreateTextFile(FolderPath&"\upload\ad"& rs("bId")&".js")
	fout.WriteLine goaler
	'关闭连接
	fout.close
	set fout = nothing
	if rs.state<>0 then rs.close
	set rs=nothing
	Response.Write "<script>alert('广告JS已经生成!');window.location.href='admin_advert.asp';</script>" 
End sub
%>
<%
private sub DeleteItem()
    dim de_id
	de_id=clng(Request("id"))
	sql="delete from Ay_Advert where bId=" & de_id
	conn.execute sql
	response.write "<script language=javascript>alert('广告删除成功!');window.location.href='admin_advert.asp';</script>"
end sub
%> <%
private sub BatchDelete()
    dim mm_ndelid
	mm_ndelid = request.Form("listid")
	if mm_ndelid  = "" then
		response.write "<script language=javascript>alert('没有任何选择!');window.location.href='admin_advert.asp';</script>"
		response.end
	end if
	sql="delete from Ay_Advert where bId in (" & Replace(mm_ndelid, "'", "''") & ")"
	conn.execute sql
	response.write "<script language=javascript>alert('批量删除成功!');window.location.href='admin_advert.asp';</script>"
end sub
%> <%
private sub ListItem()
%>
<form autocomplete="off" name="form1" id="form1" method="post" action>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>广告管理</b></font>
		<a href="index.asp?go=body">[返回]</a> <a href="admin_advert.asp">[刷新列表]</a>
		<a href="admin_advert.asp?go=add">[增加]</a> </td>
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
		<td class="t2" align="center">广告名称</td>
		<td class="t2" align="center">类型</td>		
		<td class="t2" align="center">调用地址</td>
		<td class="t2" align="center" width="60px">点击</td>		
		<td class="t2" align="left">备注</td>
		<td class="t2" align="center">操作</td>
	</tr>
	<%
	dim curpage
	
	sql="select a.* from Ay_Advert a order by a.bId desc"
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
	<tr>
		<td width="10" height="20" align="center">
		<input name="listid" type="checkbox" id="listid" value="<%=(rs("bId"))%>">
		</td>
		<td width="40" align="center"><%=trim(cstr(i + (curpage -1) * mypage.pagesize))%></td>
		<td align="center"><%=trim(rs("bTitle") & "")%></td>
		<td align="center"><%
		Select Case trim(rs("bType") & "")
			Case "1"
				response.write "图片广告"
			Case "2"
				response.write "FLASH动画"
			Case "3"
				response.write "文字广告"
		End select
		%></td>		
		<td align="left"><%response.write "&lt;script src=""../upload/ad" & rs("bId") & ".js""&gt;&lt;/script&gt;"%></td>
		<td align="center"><%=trim(rs("bClick") & "")%></td>
		<td align="left"><%=trim(rs("bRemark") & "")%></td>
		
		<td align="center" width="100px">
		<a href="admin_advert.asp?go=buildjs&id=<%=trim(rs("bId")&"")%>">生成JS</a>		
		<a href="admin_advert.asp?go=edit&id=<%=trim(rs("bId")&"")%>">修改</a>
		<a href="admin_advert.asp?go=delete&id=<%=trim(rs("bId")&"")%>" onclick="javascript:return confirm('请确认删除操作 ?')">
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
		<input type="submit" name="mydele" class="button" value="批量删除" onclick="javascript:this.form.action='admin_advert.asp?go=batchdelete';return confirm('请确认删除操作 ?')">
		<input type="checkbox" name="all" id="all" onclick="checkAll(this.checked)"><label for="all">全选</label>
		</td>
		<td align="right" height="30"><%mypage.showpage()%></td>
	</tr>
</table>
<%
end sub
%> <%
private sub AddItem()
%>
<form autocomplete="off" name="addform" id="addform" method="post" onsubmit="return checkform(addform)" action="admin_advert.asp?go=saveadd">
	<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
		<tr valign="bottom">
			<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>广告管理</b></font> 
			-&gt; 新增广告&nbsp; ← <a href="admin_advert.asp">返回列表</a> </td>
			<td></td>
		</tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td><hr size="1"></td>
		</tr>
		<tr bgcolor="#898989">
			<td height="23"><font class="t2">&nbsp;详细资料</font></td>
		</tr>
		<tr>
			<td height="10"></td>
		</tr>
	</table>
	<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">
		<tr>
			<td class="td1" align="right" valign="middle">广告名称：</td>
			<td class="td2">
			<input type="text" name="bTitle" id="bTitle" size="50"> </td>
		</tr>
		<tr>
			<td class="td1" align="right" valign="middle">广告类型：</td>
			<td class="td2">
			<select name="bType" id="bType" onchange="if(this.selectedIndex==2){document.getElementById('trscript').style.display = '';document.getElementById('trpic1').style.display = 'none';document.getElementById('trpic2').style.display = 'none';document.getElementById('trpic3').style.display = 'none';}else{document.getElementById('trscript').style.display = 'none';document.getElementById('trpic1').style.display = '';document.getElementById('trpic2').style.display = '';document.getElementById('trpic3').style.display = '';}">
			<option value="1">
			图片广告</option>
			<option value="2">
			FLASH动画</option>
			<option value="3">
			文字广告</option>
			</select>
			</td>
		</tr>		
		<tr id="trscript" style="display:none;">
			<td class="td1" align="right" valign="middle">文字内容：</td>
			<td class="td2">
			<textarea name="bScript" cols="60" rows="5" id="bScript"></textarea></td>
		</tr>		
		<tr id="trpic1">
			<td align="right" class="td1" valign="top">内 容 图 片：</td>
			<td class="td2" colspan="3">
			<input type="text" class="input" id="bPic" name="bPic" style="width:250px;" />
			<input type="checkbox" onclick="display('upload');" id="box" /><label for="box">图片地址</label>
			<font color="#ff0000">格式要求:jpg,gif,swf</font> <br>
			<div id="upload" style="display:none;" class="td2">
				<iframe src="upload.asp?go=pic" frameborder="0" style="height:22px;width:100%;" scrolling="no">
				</iframe></div>
			</td>
		</tr>
		<tr id="trpic2">
			<td class="td1" align="right" valign="middle">链接地址：</td>
			<td class="td2">
			<input name="bUrl" id="bUrl" type="text" value="http://" size="40" /> <label>打开方式： 
			<select name="bOpenMode" id="bOpenMode">
			<option value="_blank">_blank</option>
			<option value="_parent">_parent</option>
			<option value="_self">_self</option>
			<option value="_top">_top</option>
			</select> </label>
			</td>
		</tr>
		<tr id="trpic3">
			<td class="td1" align="right" valign="middle">广告尺寸：</td>
			<td class="td2">
			<input name="bPicWidth" id="bPicWidth" type="text" size="10" />
			<font color="#FF0000">*</font> 宽 
			<input name="bPicHeight" id="bPicHeight" type="text" size="10" /> <font color="#FF0000">
			*</font> 高</td>
		</tr>
		</div>
		<tr>
			<td class="td1" align="right" valign="middle">备注：</td>
			<td class="td2">
			<input type="text" name="bRemark" id="bRemark" size="50"> </td>
		</tr>		
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
		<tr>
			<td width="150" align="right" height="40"></td>
			<td><input type="submit" class="button" name="submit" value="确认提交">&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="reset" class="button" value="重新填写" name="Button"> </td>
		</tr>
	</table>
</form>
<script language="Javascript">
	addform.bTitle.focus()
</script>
<%
end sub
%> <%
private sub EditItem()
    dim mvarbTitle,mvarbType,mvarbPic,mvarbScript
    dim mvarbUrl,mvarbOpenMode,mvarbPicWidth,mvarbPicHeight,mvarbRemark
    
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Advert  where bId=" & request("id") & " order by bId"
	rs.open sql,conn,1,1
	if not rs.eof then
	    mvarbTitle=trim(rs("bTitle")&"")
	    mvarbType=trim(rs("bType")&"")
	    mvarbPic=trim(rs("bPic")&"")
	    mvarbScript=trim(rs("bScript")&"")
	    mvarbUrl=trim(rs("bUrl")&"")
	    mvarbOpenMode=trim(rs("bOpenMode")&"")
	    mvarbPicWidth=trim(rs("bPicWidth")&"")
		mvarbPicHeight=trim(rs("bPicHeight")&"")
		mvarbRemark=trim(rs("bRemark")&"")
	end if
	if rs.state<>0 then rs.close
	set rs=nothing
%>
<form autocomplete="off" name="editform" id="editform" method="post" onsubmit="return checkform(editform)" action="admin_advert.asp?go=saveedit&id=<%=request("id")%>">
	<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
		<tr valign="bottom">
			<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>广告管理</b></font> 
			-&gt; 编辑广告&nbsp; ← <a href="admin_advert.asp">返回列表</a> </td>
			<td></td>
		</tr>
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td><hr size="1"></td>
		</tr>
		<tr bgcolor="#898989">
			<td height="23"><font class="t2">&nbsp;详细资料</font></td>
		</tr>
		<tr>
			<td height="10"></td>
		</tr>
	</table>
	<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">
		<tr>
			<td class="td1" align="right" valign="middle">广告名称：</td>
			<td class="td2">
			<input type="text" name="bTitle" id="bTitle" size="50" value="<%=mvarbTitle%>"> </td>
		</tr>
		<tr>
			<td class="td1" align="right" valign="middle">广告类型：</td>
			<td class="td2">
			<select id="select1" disabled>
			<option value="1" <%If mvarbType="1" Then response.write "selected"%>>
			图片广告</option>
			<option value="2" <%If mvarbType="2" Then response.write "selected"%>>
			FLASH动画</option>
			<option value="3" <%If mvarbType="3" Then response.write "selected"%>>
			文字广告</option>
			</select>
			<input type="hidden" name="bType" id="bType" value="<%=mvarbType%>">
			</td>
		</tr>		
		<tr id="trscript" style="<%If mvarbType<>"3" Then response.write "display:none" Else response.write "display:" End if%>">
			<td class="td1" align="right" valign="middle">文字内容：</td>
			<td class="td2">
			<textarea name="bScript" cols="60" rows="5" id="bScript"><%=HTMLCode(mvarbScript)%></textarea></td>
		</tr>		
		<tr id="trpic1" style="<%If mvarbType<>"3" Then response.write "display:" Else response.write "display:none" End if%>">
			<td align="right" class="td1" valign="top">内 容 图 片：</td>
			<td class="td2" colspan="3">
			<input type="text" class="input" id="bPic" name="bPic" value="<%=mvarbPic%>" style="width:250px;" />
			<input type="checkbox" onclick="display('upload');" id="box" /><label for="box">图片地址</label>
			<font color="#ff0000">格式要求:jpg,gif,swf</font> <br>
			<div id="upload" style="display:none;" class="td2">
				<iframe src="upload.asp?go=pic" frameborder="0" style="height:22px;width:100%;" scrolling="no">
				</iframe></div>
			</td>
		</tr>
		<tr id="trpic2" style="<%If mvarbType<>"3" Then response.write "display:" Else response.write "display:none" End if%>">
			<td class="td1" align="right" valign="middle">链接地址：</td>
			<td class="td2">
			<input name="bUrl" id="bUrl" type="text" value="<%=mvarbUrl%>" size="40" /> <label>打开方式： 
			<select name="bOpenMode" id="bOpenMode">
			<option value="_blank" <%If mvarbOpenMode="_blank" Then response.write "_blank"%>>_blank</option>
			<option value="_parent" <%If mvarbOpenMode="_parent" Then response.write "_parent"%>>_parent</option>
			<option value="_self" <%If mvarbOpenMode="_self" Then response.write "_self"%>>_self</option>
			<option value="_top" <%If mvarbOpenMode="_top" Then response.write "_top"%>>_top</option>
			</select> </label>
			</td>
		</tr>
		<tr id="trpic3" style="<%If mvarbType<>"3" Then response.write "display:" Else response.write "display:none" End if%>">
			<td class="td1" align="right" valign="middle">广告尺寸：</td>
			<td class="td2">
			<input name="bPicWidth" id="bPicWidth" type="text" value="<%=mvarbPicWidth%>" size="10" />
			<font color="#FF0000">*</font> 宽 
			<input name="bPicHeight" id="bPicHeight" type="text" value="<%=mvarbPicHeight%>" size="10" /> <font color="#FF0000">
			*</font> 高</td>
		</tr>
		</div>
		<tr>
			<td class="td1" align="right" valign="middle">备注：</td>
			<td class="td2">
			<input type="text" name="bRemark" id="bRemark" value="<%=mvarbRemark%>" size="50"> </td>
		</tr>		
	</table>
	<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
		<tr>
			<td width="150" align="right" height="40"></td>
			<td><input type="submit" class="button" name="submit" value="确认提交">&nbsp;&nbsp;&nbsp;&nbsp;
			<input type="reset" class="button" value="重新填写" name="Button"> </td>
		</tr>
	</table>
</form>
<script language="Javascript">
	editform.bTitle.focus()
</script>
<%
end sub
%>