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
		if(form("bClassID").value=="0"){alert("请选择分类!");form("bClassID").focus();return false;}	
		if(form("bTitle").value==""){alert("标题名称不能为空!");form("bTitle").focus();return false;}		
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
		response.write "<script>alert('SQL防注入配置成功，请继续操作！');window.location.href='admin_sql.asp?go=edit';</script>"
		response.end
	end if
end sub
%>
<%
private sub BatchDelete()
    dim mm_ndelid
	mm_ndelid = request.Form("listid")
	if trim(mm_ndelid)  = "" then
		response.write "<script language=javascript>alert('没有任何选择!');window.location.href='admin_sql.asp';</script>"
		response.end
	end if
	sql="delete from Ay_SqlIn where bId in (" & Replace(mm_ndelid, "'", "''") & ")"
	
	conn.execute sql
	response.write "<script language=javascript>alert('批量删除成功!');window.location.href='admin_sql.asp';</script>"
end sub
%>

<%
private sub ListItem()
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<form autocomplete="off" name="form1" id="form1" method="post" action>
	<tr valign="bottom">
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>SQL注入管理</b></font>&nbsp;<a href="index.asp?go=body">[返回]</a>&nbsp;<a href="admin_sql.asp">[刷新列表]</a></td>
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
		<td class="t2" align="center">攻击ＩＰ</td>
        <td class="t2" align="center">当前状态</td>
        <td class="t2" align="center">是否锁定</td>
        <td class="t2" align="center">操作页面</td>
        <td class="t2" align="center">操作时间</td>
        <td class="t2" align="center">提交方式</td>
        <td class="t2" align="center">提交参数</td>
        <td class="t2" align="center">提交数据</td>

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
			response.write ">"
	%>			
		<td width="10" height="20" align="center">
		<input name="listid" type="checkbox" id="listid" value="<%=(rs("bId"))%>">
		</td>
		<td width="40" align="center"><%=trim(cstr(i + (curpage -1) * mypage.pagesize))%></td>	
		
		<td align="center"><%=trim(rs("bIPAddress") & "")%></td>
		<td align="center">
			<%	if rs("bIsKill")=1 then 
					response.write "<font color='red'>已锁定</font>"
				else
					response.write "<font color='green'>已解锁</font>"
				end if
			%>
		</td>		
		<td align="center">
			<%	if rs("bIsKill")=1 then 
					response.write "<a href=admin_sql.asp?go=unlock&id="&rs("bId")&" style=""color:#FF0000"">解锁IP</a>"
				else
					response.write "<a href=admin_sql.asp?go=lock&id="&rs("bId")&" style=""color:#006600"">锁定IP</a>"
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
		<input type="submit" name="mydele" class="button" value="批量删除" onclick="javascript:this.form.action='admin_sql.asp?go=batchdelete';return confirm('请确认删除操作 ?')">
		<input type="checkbox" name="all" id="all" onclick="checkAll(this.checked)"><label for="all">全选</label>
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
		<td>你现在的位置：<font color="DarkSlateGray" style="font-size:12px"><b>SQL防注入设置</b></font>&nbsp;->&nbsp;编辑配置&nbsp;&nbsp;←&nbsp;<a href="admin_sql.asp?go=edit">返回列表</a> </td>
		<td></td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr><td><hr size="1"></td></tr>
	<tr bgcolor="#898989"><td height="23"><font class="t2">&nbsp;详细资料</font></td></tr>
	<tr><td height="10"></td></tr>
</table>
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">	
<form autocomplete="off" name="editform" id="editform" method="post">	 					
	<tr>
		<td align="right" class="td1" valign="middle">需要过滤的关键字：：</td>
		<td class="td2">
		<input name="bFilterKeys" type="text" value="<%=mvarbFilterKeys%>" id="bFilterKeys" style=" " size="50">
                  用&quot;|&quot;分开
		</td>			
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle">是否记录入侵者信息：</td>
		<td class="td2">
		    <select name="bIsWriteLog" id="bIsWriteLog">
              <option value="1" <%if mvarbIsWriteLog="1" Then response.write "selected"%>>是</option>
              <option value="0" <%if mvarbIsWriteLog="0" Then response.write "selected"%>>否</option>
          </select>
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">是否启用锁定IP：</td>
		<td class="td2">
		<select name="bIsKill" id="bIsKill">
          <option value="1" <%if mvarbIsKill="1" Then response.write "selected"%>>是</option>
          <option value="0" <%if mvarbIsKill="0" Then response.write "selected"%>>否</option>
      </select>
        </td>
	</tr>	
	<tr>
		<td width="15%" class="td1" align="right">是否启用安全页面：</td>
		<td width="85%" class="td2">
		<select name="bIsSafeOpen" id="bIsSafeOpen">
                      <option value="1" <%if mvarbIsSafeOpen="1" Then response.write "selected"%>>是</option>
                      <option value="0" <%if mvarbIsSafeOpen="0" Then response.write "selected"%>>否</option>
                    </select>
                  慎用这个功能，除非你对确认此页面无需过滤，并确定对安全没影响！
		</td>
 	</tr> 
  	<tr>
		<td align="right" class="td1">您认为安全的页面：</td>
		<td class="td2">
		<input name="bSafePage" type="text" value="<%=mvarbSafePage%>" id="bSafePage" style=" " size="50">
                  用&quot;|&quot;分开
		</td>
  	</tr> 
	<tr>
		<td align="right" class="td1">出错后的处理方式：</td>
		<td class="td2">
		<select name="bErrorHandle" id="bErrorHandle">
          <option value="1" <%if mvarbErrorHandle="1" Then response.write "selected"%>>直接关闭网页</option>
          <option value="2" <%if mvarbErrorHandle="2" Then response.write "selected"%>>警告后关闭</option>
          <option value="3" <%if mvarbErrorHandle="3" Then response.write "selected"%>>跳转到指定页面</option>
          <option value="4" <%if mvarbErrorHandle="4" Then response.write "selected"%>>警告后跳转</option>
      </select>
		</td>
	</tr>	
	<tr>
		<td align="right" class="td1" valign="top">出错后跳转Url：</td>
		<td class="td2">			
			<input name="bAlertUrl" type="text" value="<%=mvarbAlertUrl%>" id="bAlertUrl"  size="30">
		</td>
	</tr>
 	<tr>
		<td align="right" class="td1" valign="top">警告提示信息：</td>
		<td class="td2">			
			<textarea name="bAlertInfo" cols="45" rows="4" id="bAlertInfo"><%=mvarbAlertInfo%></textarea>                  
                  \n\n换行
		</td>
	</tr>
	<tr>
		<td align="right" class="td1" valign="top">阻止访问提示信息：</td>
		<td class="td2">			
			<textarea name="bKillInfo" cols="45" rows="4" id="bKillInfo"><%=mvarbKillInfo%></textarea>
                  \n\n换行
		</td>
	</tr>

</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr>
		<td width="150" align="right" height="40"></td>
		<td>
		<input type="button" class="button" name="submit1" value="确认提交" onclick="this.form.action='admin_sql.asp?go=saveedit';this.form.submit();">&nbsp;&nbsp;&nbsp;&nbsp;
		<input type="reset" class="button" value="重新填写" name="Button">
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