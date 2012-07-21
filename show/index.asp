<!--#include file="../conn.asp"-->
<%
dim m_querystring,cid,id
m_querystring=split(replace(Request.ServerVariables("QUERY_STRING"),".html","") & "-","-")
cid=m_querystring(0)
id=m_querystring(1)
%>
<%
dim mvarbClick,mvarbId,mvarbTitle,mvarbWriter,mvarbCopyRight,mvarbContent,mvarbAddTime
set rs=server.CreateObject("adodb.recordset")
msql="select  * from Ay_Content_v where bId=" & id
rs.open msql,conn,1,3																
if  not rs.eof then	
	mvarbId=trim(rs("bId") & "")
	mvarbTitle=Juncode(trim(rs("bTitle")&""))
	mvarbAddTime=FormatDate(trim(rs("bAddTime")&""),"2")
	mvarbWriter=trim(rs("bWriter")&"")
	mvarbCopyRight=trim(rs("bCopyRight")&"")
	mvarbContent=Juncode(trim(rs("bContent")&""))
	mvarbClick=trim(rs("bClick")&"")
	mvarbClassID=trim(rs("bClassID") & "")
	rs("bClick")=rs("bClick")+1
	rs.update
end if
if rs.state<>0 then rs.close
set rs=nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
    <title>
        <%=mvarbTitle%>
        -
        <%=ay_sitetitle%>
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
        <strong>当前位置：&nbsp;<%call ChannelNav(mvarbClassID)%>内容</strong>
    </div>
    <div class="main">
      <div class="channel_left">
            <div class="border">
                <h1>
                    <%=mvarbTitle%>
                </h1>
                <div class="info">
                    <%=mvarbWriter%>
                    &nbsp;&nbsp; 发布时间：<%=mvarbAddTime%>
                    来源：<%=mvarbCopyRight%></div>
                <div class="common">
                    <p>

                       

                    </p>
                    <p>
                        <%=mvarbContent%>
          <p>

                         

                        </p>
                </div>
                <div class="per_nex">
                    <p>
                        上一篇：<%call ShowPrev(mvarbClassID,mvarbId)%></p>
                    <p>
                        下一篇：<%call ShowNext(mvarbClassID,mvarbId)%></p>
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
