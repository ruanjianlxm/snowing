<!--#include file="../conn.asp"-->
<!--#include file="../inc/class_page.asp" -->
<%
dim mvarKeywords,mvarSearchtype
mvarKeywords=request("keyword")
%>
<%
dim m_querystring,cid,id
m_querystring=split(Split(replace(Request.ServerVariables("QUERY_STRING"),".html","") & "_","_")(0) & "-" ,"-")
cid=m_querystring(0)
id=m_querystring(1)
cpage=Split(replace(Request.ServerVariables("QUERY_STRING"),".html","") & "_","_")(1)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<title>����-<%=mvarKeywords%>-<%=ay_sitename%></title>
<meta name="keywords" content="<%=ay_keywords%>" />
<meta name="description" content="<%=ay_description%>" />
<link href="../images/css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div class="head">
	<div class="logo">
		<img src="../images/logo.gif" width="262" height="60" /></div>
	<div class="banner">
		<script src="../upload/ad1.js"></script>
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
	<strong>��ǰλ�ã�
	<a href="../index.asp">��վ��ҳ</a> &gt;&nbsp;�������&nbsp;&nbsp;<font color="#ff0000"><%=mvarKeywords%></font>
	</strong>
</div>
<div class="index_gd">
	<div class="gd_left">
	</div>
	<%call HotNews()%>
	<div class="gd_right">
	<%call SearchForm()%>
	</div>
	<div style="clear: both;">
	</div>
</div>
<div class="main">
	<div class="channel_left">
		<div class="border mt8">
			<%
			dim curpage					
			sql="select a.* from Ay_Content_v a where 1=1 "
			sql=sql & " and (a.bTitle like '%" & mvarKeywords & "%' or a.bContent like '%" & mvarKeywords & "%')"
			sql=sql & " order by a.bId desc,bAddTime desc"
			Set rs=Server.CreateObject("Adodb.Recordset")
			Set mypage=new xdownpage
			mypage.getconn=conn
			mypage.getsql=sql
			mypage.pagesize=20
			set rs=mypage.getrs()	
			if cpage<>"" then
				curpage=clng(cpage+0)
			else
				curpage=1
			end if	
			%>
			<div class="title2">
				<strong>�ҵ���ؼ�¼Լ<%=mypage.RecordCount()%>ƪ</strong>
			</div>
			<ul class="list_list">
			<%	
			if rs.eof and rs.bof then			
				response.write("<p>�Ҳ����κμ�¼��</p>")				
			end if
			for i=1 to mypage.pagesize
				if rs.eof  then exit for										
			%>
				<li <%If (i Mod 5)=0 Then response.write "class=""borbom""" End if%>><span><%=FormatDate(trim(rs("bAddTime") & "") ,"1")%></span>
				<a href="../show/?<%=rs("bClassID")%>-<%=rs("bId")%>.html" title="<%=trim(rs("bTitle") & "")%>"><%=Search(trim(rs("bTitle") & ""),mvarKeywords)%></a>
				</li>
			<%	
				rs.MoveNext 
			next		
			if rs.State <>0 then rs.Close 
			set rs=nothing
			%>
			</ul>
			<ul class="pagelist">
				<%call mypage.showpage()%>
			</ul>
		</div>
		<div style="clear: both;">
		</div>
	</div>
	<div class="channel_right mt8">
		<div class="border">
			���λ��250*250 </div>
		<div class="border mt8">
			<div class="title">
				<strong>�Ƽ�ͼ��</strong></div>
			<ul class="pic_text">
				<%call Pic_Text()%>
			</ul>
		</div>
		<div class="border mt8">
			<div class="title">
				<strong>�Ƽ�����</strong></div>
			<ul class="text_list">
				<%call Text_List()%>
			</ul>
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
