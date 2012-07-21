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
    <title>
        <%=ay_sitename%>
    </title>
    <meta name="keywords" content="<%=ay_keywords%>" />
    <meta name="description" content="<%=ay_description%>" />
    <link href="../images/css.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <div class="head">
   
    </div>
    <div id="menu_out">
        <div id="menu_in">
            <div id="menu">
                <%call HeadNavigation()%>
            </div>
        </div>
    </div>
    <div class="navdh">
        <strong>当前位置：&nbsp;<a href="../index.asp">网站首页</a> &gt;&nbsp;<%call ChannelNav(cid)%> 列表 </strong>
    </div>
    <div class="main">
      <div class="channel_left">
            <div class="border mt8">
                <div class="title">
                    <strong>
                        <%
				set rs=server.createobject("adodb.recordset")				
				if cid<>"" then
					sql="select * from Ay_Class where bId=" & cid& ""
					rs.open sql,conn,1,1
					if not rs.eof then	
						if trim(rs("bParent") & "")<>"0" then
							response.write trim(rs("bName") & "")
						else
							response.write "全部分类"
						end if
					end if
					if rs.state<>0 then rs.close
					set rs=nothing
				else
					response.write "全部分类"
				end if
                        %>
                    </strong>
                </div>
                <ul class="list_list">
                    <%
			dim curpage					
			sql="select a.* from Ay_Content_v a where 1=1 "
			if cid <>"" then
				sql=sql & " and (a.bClassID=" & cid & " or a.bParentID=" & cid & ")"
			end if
			sql=sql & " order by a.bAddTime desc,a.bId desc"
			Set rs=Server.CreateObject("Adodb.Recordset")
			Set mypage=new xdownpage
			mypage.getconn=conn
			mypage.getsql=sql
			mypage.pagesize=15
			set rs=mypage.getrs()	
			if cpage<>"" then
				curpage=clng(cpage+0)
			else
				curpage=1
			end if		
			if rs.eof and rs.bof then			
				response.write("<p>找不到任何记录！</p>")				
			end if
			for i=1 to mypage.pagesize
				if rs.eof  then exit for										
                    %>
                    <li <%if (i mod 5)=0 then response.write "class=""borbom""" end if%>><span>
                        <%=FormatDate(trim(rs("bAddTime") & "") ,"1")%>
                    </span><a href="../show/?<%=rs("bClassID")%>-<%=rs("bId")%>.html" title="<%=trim(rs("bTitle") & "")%>">
                        <%=trim(rs("bTitle") & "")%><span style="float: none; color:#F00;"><%if (i < 6) then response.write " NEW" end if%></span>
                    </a></li>
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
      <div style="clear: both;">
      </div>
    </div>
    <div class="main">
        <div class="footer">
          Copyright &copy; 2008-2015 黑龙江科技学院-黑龙江省东部煤电化工程技术研发平台. All rights reserved.
        </div>
    </div>
</body>
</html>
