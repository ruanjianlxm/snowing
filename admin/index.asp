<%@language=vbscript codepage=936 %>
<%
option explicit
'ǿ����������·��ʷ���������ҳ�棬�����Ǵӻ����ȡҳ��
Response.Buffer = True 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1 
Response.Expires = 0 
Response.CacheControl = "no-cache" 
%>
<!--#include file="conn.asp"-->
<!--#include file="../inc/md5.asp"-->

<%
		
select case request("go")    
    case "top"
        call Admin_Top()  
    case "left"
        call Admin_Left()  
	Case "buttom"
		Call Admin_Buttom()
	Case "body"
	    if session("admin")<>"" then
		    Call Admin_Body()
		else
		    call Admin_Login()
		end if
	case "login"
	    call Check_Login()
	case "logout"
	    session("admin")=""
		Session.Abandon
	    Response.Redirect "index.asp"
    case else
        call Admin_Main()
end Select

Call CloseConn()
%>
<%
private sub Check_Login()

	dim FoundErr,ErrMsg
	
	dim username,password
	username=replace(trim(request("username")),"'","")
	password=replace(trim(Request("password")),"'","")
	
	password=md5(password)
	set rs=server.createobject("adodb.recordset")
	sql="select * from Ay_Admin where  bName='"&username&"'"
	rs.open sql,conn,1,3
	if rs.bof and rs.eof then
		FoundErr=True
		ErrMsg="�û�������ȷ����"
	else
		if password<>rs("bPassword") then
			FoundErr=True
			ErrMsg="������󣡣�"
		else
      session("LoginTime")=trim(rs("bLoginTime") & "")
      session("LoginCount")=trim(rs("bLoginCount") & "")			
			rs("bLoginIP")=Request.ServerVariables("REMOTE_ADDR")
			rs("bLoginTime")=now()
			rs("bLoginCount")=rs("bLoginCount")+1
			rs.update			
			session("admin")=rs("bName")			
			rs.close
			
			sql="select * from Ay_System"
			rs.open sql,conn,1,1
			if not rs.eof then
				session("UserName")=trim(rs("bUserName")&"")
			end if
			rs.close
			set rs=nothing
			call CloseConn()
			Response.Redirect "index.asp"
		end if
	end if
	rs.close
	set rs=nothing
	call CloseConn()
	if FoundErr=True then	    
		response.write "<script>alert('" & ErrMsg & "');top.location.href='index.asp';</script>"
		response.end
	end if
	
end sub
%>
<%
private sub Admin_Main()
%>
<html>
<head>
<title>��վ��������</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<meta http-equiv="Expires" content="0">
</head>
<frameset rows="40,*,25" col="*" framespacing="0">
	<frame name="title" src="index.asp?go=top" target="main" scrolling="no" noresize>
	<frameset cols="190,*" framespacing="0" >
		<frame name="tree" src="index.asp?go=left" target="main" frameborder="no">
		<frame name="main" src="index.asp?go=body" frameborder="NO">
	</frameset>
	<frame name="buttom" scrolling="No" noresize target="main" src="index.asp?go=buttom">
	<noframes>
		<body bgcolor="#FFFFFF">
			<p>����ҳʹ���˿�ܣ��������������֧�ֿ�ܡ�</p>
		</body>
	</noframes>
</frameset>

</html>
<%
end sub
%>
<%
private sub Admin_Buttom()
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>New Page 1</title>
<style>
BODY{font-family:verdana,arial,helvetica;margin:0;}
td {font-family:Tahoma,Verdana, Arial;font-size:11px;}
A:link, A:active,A:visited{color: #FFFFFF;text-decoration: none;padding-left:6px;padding-right:6px;}
A:hover{color: #FF3300;text-decoration: none;padding-left:6px;padding-right:6px;}
.STYLE1 {color: #CCCCCC}
</style>
</head>
<BODY TOPMARGIN="0"  LEFTMARGIN="0" MARGINHEIGHT="0" MARGINWIDTH="0" BGCOLOR="#31615A" TEXT="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td height="25" width="228" nowrap valign="middle" class="STYLE1">��ǰ�û���<%=session("admin")%></td>
		
	  <td  height="25" align="right" nowrap valign="middle"> 
	  <span class="STYLE1">
	  <script>document.write(unescape('%u7A0B%u5E8F%u5236%u4F5C%u3001%u9875%u9762%u8BBE%u8BA1%uFF1A%u9ED1%u9F99%u6C5F%u79D1%u6280%u5B66%u9662%u4FE1%u606F%u7F51%u7EDC%u4E2D%u5FC3%20%3Ca%20href%3D%22http%3A//inc.usth.net.cn/%22%20target%3D%22_blank%22%3Ehttp%3A//inc.usth.net.cn/%3C/a%3E%20'))</script>
	  </span>
	  </td>
	  <td width="10px"></td>
	</tr>
</table>
</BODY> 

</html>
<%
end sub
%>
<%
private sub Admin_Top()
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>New Page 1</title>
<style>
BODY{font-family:verdana,arial,helvetica;margin:0;}
td {font-family:Tahoma,Verdana, Arial;font-size:11px;}
A:link, A:active,A:visited{color: #FFFFFF;text-decoration: none;padding-left:6px;padding-right:6px;}
A:hover{color: #FF3300;text-decoration: none;padding-left:6px;padding-right:6px;}
</style>
</head>
<BODY TOPMARGIN="0"  LEFTMARGIN="0" MARGINHEIGHT="0" MARGINWIDTH="0" BGCOLOR="#31615A" TEXT="#000000" style="border-bottom:solid 2px #000000;">
<table border="0"  id="tbBody" cellspacing=0 cellpadding=0 width="100%" height="100%">
	<tr>		
		<td valign="middle"><font size="4" color="#ffffff"><strong style="padding:10px;">�������Ƽ�ѧԺ - ������ʡ����ú�绯���̼����з�ƽ̨ - ���¹�������</strong></font></td>
		<td valign="bottom">
		    <table  border="0" cellspacing="0" cellpadding="0" align="right">
	            <tr height="24" >
		            <td>��</td>
		            <td width="60" align="center" valign="middle">
		            <a href="index.asp?go=body" target="main">������ҳ</a>
		            </td>
		            <td width="60" align="center" valign="middle">
		            <a href="admin_password.asp" target="main">�޸�����</a>
		            </td>
		            <td width="60" align="center" valign="middle">
		            <a href="../../../index.asp" target="_blank">��վ��ҳ</a>
		            </td>
		            <td  width="60" align="center" valign="middle">
		            <a href="index.asp?go=logout" target="_top">�˳�ϵͳ</a>
		            </td>
	            </tr>	            
            </table>
		</td>
	</tr>	
</table>
</body>
</html>
<%
end sub
%>
<%
private sub Admin_Left()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<meta http-equiv="Expires" content="0">
<script type="text/javascript" src="dtree.js"></script>
<style type="text/css" >
body{margin:0px;}
.icotree {padding:6px;font-family: Verdana, Geneva, Arial, Helvetica, sans-serif;font-size: 12px;color: #666;white-space: nowrap;}
.icotree img {border: 0px;vertical-align: middle;}
.icotree a {color: #333;text-decoration: none;}
.icotree a.node, .icotree a.nodeSel {white-space: nowrap;padding: 1px 2px 1px 2px;}
.icotree a.node:hover, .icotree a.nodeSel:hover {color: #333;text-decoration: underline;}
.icotree a.nodeSel {background-color: #c0d2ec;}
.icotree .clip {overflow: hidden;}
</style>
</head>
<body  TOPMARGIN="0"  LEFTMARGIN="0" MARGINHEIGHT="0" MARGINWIDTH="0"  TEXT="#000000" style="border-right:2px solid #000000;">
<div class="icotree">
<script type="text/javascript">
	<!--
	d = new dTree('d');
	d.config.folderLinks=false;
	d.config.useCookies=false;
	d.config.target = "main";
	d.add(0,-1,'�ҵĿ������');
	d.add(1,0,'��������','');
	d.add(11,1,'��վ����','admin_system.asp');	
	d.add(12,1,'�������','admin_notice.asp');
	d.add(13,1,'SQLע��','admin_sql.asp');
	d.add(15,1,'ע������','admin_sql.asp?go=edit');
	d.add(3,0,'���¹���','');
	d.add(31,3,'�������','admin_news.asp?go=add');	
	d.add(32,3,'��������','admin_news.asp');
	d.add(33,3,'���·���','admin_class.asp');	
	d.add(5,0,'������','');
	d.add(51,5,'��ӹ��','admin_advert.asp?go=add');	
	d.add(52,5,'������','admin_advert.asp');		
	d.add(6,0,'��������','');	
	d.add(61,6,'���ݱ���','admin_data.asp?action=BackupData');
	d.add(62,6,'���ݻָ�','admin_data.asp?action=RestoreData');
	d.add(63,6,'ѹ�����ݿ�','admin_data.asp?action=CompressData');	
	d.add(64,6,'��������','admin_link.asp');
	d.add(65,6,'�ؼ��ֹ���','admin_keyword.asp');
	d.add(66,6,'���Թ���','admin_book.asp');
	document.write(d);
	//-->
</script>
</div>
</body>
</html>
<%
end sub
%>
<%
Private Sub Admin_Body()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<meta http-equiv="Expires" content="0">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="dtree.js"></script>
</head>
<body topmargin="5" leftmargin="5" bgcolor="#ffffff">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="24">
	<tr valign="bottom">
		<td>�����ڵ�λ�ã�<b><%=Request.ServerVariables("Http_HOST")%> -&gt; ��������</b> </td>
		<td></td>
	</tr>
</table>
<table width="100%" align="center" border="0" cellspacing="2" cellpadding="0">
	<tr>
		<td rowspan="3" width="120" align="center"><img src="images/admin_p.jpg" width="200" height="50"></td>
		<td rowspan="3" width="100">��</td>
		<td style="color:#191970;" height="30"><%=Now()%></td>
	</tr>
	<tr>
		<td style="font-family:����;font-size:20px;line-height:30px;" height="50">������ʡ����ú�绯���̼����з�ƽ̨</td>
	</tr>
	<tr>
		<td height="30">��ӭ����<span style="font-weight:bold;color:#00008B;padding:5px;"><%=Request.ServerVariables("HTTP_HOST")%></span>��������</td>
	</tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
	<tr height="12" bgcolor="#EEEEEE">
		<td></td>
	</tr>
	<tr height="25" bgcolor="#31615A">
		<td  style="color:#FFFFFF;padding-left:10px;" valign="middle">���������Ϣ
		</td>
	</tr>
</table> 
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">	
	<tr>
		<td class="td1" align="right" valign="middle" height="22" nowrap>�ٷ����棺</td>
		<td class="td2" style="color:#800000" id="msg">
		<script>checkupdate("http://paducn.com/update.asp","msg")</script>
		</td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle" height="22" nowrap>��ǰ�汾��</td>
		<td class="td2" style="color:#800000">V2.0 Build20100723
		</td>
	</tr>	
	<tr>
		<td class="td1" align="right" valign="middle" height="22" nowrap>��ǰIP��ַ��</td>
		<td class="td2" style="color:#800000"><%=Request.ServerVariables("REMOTE_ADDR")%></td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle" height="22" nowrap>���ߴ�����</td>
		<td class="td2" style="color:#800000"><%=session("LoginCount")%></td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle" height="22" nowrap>����ʱ�䣺</td>
		<td class="td2" style="color:#800000"><%=session("LoginTime")%></td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle" height="22" nowrap>��ַ��</td>
		<td class="td2" style="color:#800000"><%=Request.ServerVariables("HTTP_HOST")%></td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle" height="22" nowrap>������汾��</td>
		<td class="td2" style="color:#800000"><%=Request.ServerVariables("HTTP_USER_AGENT")%></td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle" height="22" nowrap>WEB��������</td>
		<td class="td2" style="color:#800000"><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
	</tr>
	<tr>
		<td class="td1" align="right" valign="middle" height="22" nowrap>��ݹ��ڣ�</td>
		<td class="td2" style="color:#800000"><%=Session.timeout%> ����</td>
	</tr>
</table> 
<table width="100%" border="0" cellspacing="1" cellpadding="2" align="center" bgcolor="#e6e6e6">
	<tr>
		<td colspan=2><b><font color="#ff0000" size="2">��ȫ��ʾ</font></b></td>
	</tr>
	<tr>
		<td rowspan="2" class="td2"></td>
		<td class="td2">�붨�ڸ��������Ա�֤���ʰ�ȫ������Ӧ����6λ�������Ϊ�������������ĸ���б����ŵ���ϣ�����[A43Q&ma1#b6]
		</td>
	</tr>
	<tr>
		
	</tr>
</table>
</body>
</html>
<%
End Sub 
%>

<%
private sub Admin_Login()
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Cache-Control" content="no-cache">
<meta http-equiv="Expires" content="0">
<link href="images/css.css" rel="stylesheet" type="text/css">
<script language=javascript>
function SetFocus()
{
if (document.Login.username.value=="")
	document.Login.username.focus();
else
	document.Login.username.select();
}
function CheckForm()
{
	if(document.Login.username.value=="")
	{
		alert("�������û�����");
		document.Login.username.focus();
		return false;
	}
	//if(document.Login.Password.value == "")
	//{
	//	alert("���������룡");
	//	document.Login.Password.focus();
	//	return false;
	//}
	
}
</script>
</head>
<body onLoad="SetFocus();" bgcolor="#FFFFFF" >
<p></p>
<br><br>
<form autocomplete="off" name="Login" id="Login" action="index.asp?go=login" method="post" target="_top" onSubmit="return CheckForm();">
    <table width="250" border="0"  cellspacing="8" cellpadding="0" align="center">
          <tr> 
            <td align="right">�û����ƣ�</td>
            <td><input name="username"  type="text"  id="username" size="22"></td>
          </tr>
          <tr> 
            <td align="right">�û����룺</td>
            <td><input name="Password"  type="password"  size="22"></td>
          </tr>
          
          <tr> 
            <td colspan="2" height="30"> <div align="center"> 
                <input   type="submit" name="Submit" class="button" value="ȷ����¼">
                &nbsp;&nbsp;
                <input name="reset" type="reset" class="button" id="reset" value="������д">
                <br>
              </div></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
   </form> 
</body> 
</html> 
<%
end sub
%>