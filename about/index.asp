<!--#include file="../conn.asp"-->

<%
dim m_querystring,cid,id
m_querystring=split(replace(Request.ServerVariables("QUERY_STRING"),".html","") & "-","-")
cid=m_querystring(0)
id=m_querystring(1)
%>
<%
dim mvarbTitle,mvarbContent
set rs=server.CreateObject("adodb.recordset")
msql="select  * from Ay_Notice where bId=" & cid
rs.open msql,conn,1,1																
if  not rs.eof then	
	mvarbTitle=trim(rs("bTitle")&"")
	mvarbContent=trim(rs("bContent")&"")
end if
if rs.state<>0 then rs.close
set rs=nothing
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<title><%=mvarbTitle%> - <%=ay_sitetitle%></title>
<meta name="keywords" content="<%=ay_keywords%>" />
<meta name="description" content="<%=ay_description%>" />
<link href="../images/css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div class="head">
	
	<div class="banner">
		
	</div>	
</div>
<div id="menu_out">
	<div id="menu_in">
		<div id="menu">
			<%call HeadNavigation()%>
		</div>
	</div>
</div>
<div class="navdh">
	<strong>当前位置：&nbsp;系统公告</strong>
</div>
<div class="main">
	<div class="channel_left">
		<div class="border mt8">
			<h1><%=mvarbTitle%></h1>			
			<div class="common">
				<p>
				<script src="upload/ad2.js" language="javascript"></script>
				</p>				
				<p>
				<%=mvarbContent%>
				<p>
				<script src="upload/ad2.js" language="javascript"></script>
				</p>
			</div>	
		</div>
	</div>
	
	<div style="clear: both;">
	</div>
</div>
<div class="main">
	<div class="footer">	
	<%Call Footer()%>
	</div>
</div>
</body>

</html>
