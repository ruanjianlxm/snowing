<!--#include file="../conn.asp"-->
<!--#include file="../inc/class_page.asp" -->
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
<title><%=ay_sitename%></title>
<meta name="keywords" content="<%=ay_keywords%>" />
<meta name="description" content="<%=ay_description%>" />
<link href="../images/css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div class="head"> </div>
<div id="menu_out">
  <div id="menu_in">
    <div id="menu">
      <%call HeadNavigation()%>
    </div>
  </div>
</div>
<div class="navdh"> <strong>��ǰλ�ã�&nbsp;<a href="../index.asp">��վ��ҳ</a> &gt;<a href="../list/?42.html">��������</a> &gt;�б� </strong> </div>
<div class="main">
  <div class="channel_left">
    <div class="border mt8">
      <div class="title"> <strong> ���� </strong> </div>
      <ul class="list_list">
        <LI><A title=�л����񹲺͹���ȫ������ 
  href="../list/?43.html">����</A></LI>
   <LI><A title=�л����񹲺͹���ȫ������ 
  href="../list/?44.html">��������</A></LI>
   <LI><A title=�л����񹲺͹���ȫ������ 
  href="../list/?45.html">���Ź���</A></LI>
   <LI><A title=�л����񹲺͹���ȫ������ 
  href="../list/?46.html">���߽��</A></LI>
   <LI><A title=�л����񹲺͹���ȫ������ 
  href="../list/?47.html">���ұ�׼</A></LI>
   <LI><A title=�л����񹲺͹���ȫ������ 
  href="../list/?48.html">��ҵ��׼</A></LI>
      </ul>
    </div>
    <div style="clear: both;"> </div>
  </div>
  <div style="clear: both;"> </div>
</div>
<div class="main">
  <div class="footer"> Copyright &copy; 2008-2015 �������Ƽ�ѧԺ-������ʡ����ú�绯���̼����з�ƽ̨. All rights reserved. </div>
</div>
</body>
</html>
