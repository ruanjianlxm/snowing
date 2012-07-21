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
<div class="navdh"> <strong>当前位置：&nbsp;<a href="../index.asp">网站首页</a> &gt;<a href="../list/?42.html">资料中心</a> &gt;列表 </strong> </div>
<div class="main">
  <div class="channel_left">
    <div class="border mt8">
      <div class="title"> <strong> 分类 </strong> </div>
      <ul class="list_list">
        <LI><A title=中华人民共和国安全生产法 
  href="../list/?43.html">法律</A></LI>
   <LI><A title=中华人民共和国安全生产法 
  href="../list/?44.html">行政法规</A></LI>
   <LI><A title=中华人民共和国安全生产法 
  href="../list/?45.html">部门规章</A></LI>
   <LI><A title=中华人民共和国安全生产法 
  href="../list/?46.html">政策解读</A></LI>
   <LI><A title=中华人民共和国安全生产法 
  href="../list/?47.html">国家标准</A></LI>
   <LI><A title=中华人民共和国安全生产法 
  href="../list/?48.html">行业标准</A></LI>
      </ul>
    </div>
    <div style="clear: both;"> </div>
  </div>
  <div style="clear: both;"> </div>
</div>
<div class="main">
  <div class="footer"> Copyright &copy; 2008-2015 黑龙江科技学院-黑龙江省东部煤电化工程技术研发平台. All rights reserved. </div>
</div>
</body>
</html>
