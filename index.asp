<!--#include file="conn.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<!-- Powered by UCAIS - Linfcstmr -->
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<title><%=ay_sitename%></title>
<meta name="keywords" content="<%=ay_keywords%>" />
<meta name="description" content="<%=ay_description%>" />
<link href="images/css.css" rel="stylesheet" type="text/css" />
</head>
<body>
<div id="content">
<div class="head">
  
</div>
<div id="menu_out">
  <div id="menu_in">
    <div id="menu">
      <%call HeadNavigation()%>
    </div>
  </div>
</div>
<div class="main">
  <div class="flash border">
    <%Call FlashNews()%>
  </div>
  <div class="top_news border"> <a style="cursor:hand;" href="../list/?26.html"><b class="news"></b></a>
    <%call Top_News()%>
  </div>
  <div class="notice border">
    <%call Notice()%>
  </div>
  <div style="clear: both;"> </div>
</div>
<div class="mainStage">
  <div class="dataCenter">
    <div class="bb" style=" margin: -10px 0 0 0px;"><br />
 
    </div>
    
    
     <div class="qyfc border">
    <div class="title"> <span></span><strong><a href="../list/?49.html">��ҵ���</a></strong></div>
    <div style=" margin: 4px;"><%Call QiyeFlash()%></div>
  </div>
    
    
    
    
  </div>
  
  <div class="ad1">
    <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=9,0,28,0" width="732" height="70">
      <param name="movie" value="images/ad1.swf" />
      <param name="quality" value="high" />
      <param name="wmode" value="opaque" />
      <embed src="images/ad1.swf" quality="high" wmode="opaque" pluginspage="http://www.adobe.com/shockwave/download/download.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="732" height="70"></embed>
    </object>
  </div>
 
  <div class="mainBlock">
    <%
set rs=server.createobject("adodb.recordset")
msql="select * from Ay_Class where bShowIndex=1 order by bParent,bOrder"
rs.open msql,conn,1,1
dim clsindex
clsindex=1
do while not rs.eof
	select case (clsindex mod 2)
		case 1
%>
    <div class="main_left border">
      <%call Index_Class(rs("bId"))%>
    </div>
    <%
		case 0
%>
    <div class="main_right border ">
      <%call Index_Class(rs("bId"))%>
    </div>
    <div style="clear: both;"> </div>
  </div>
  <div class="mainBlock">
    <%
		end select 
	rs.movenext
	clsindex=clsindex +1
loop
if rs.state<>0 then rs.close
set rs=nothing
%>
  </div>
  
  <div class="linkStage border">
    <div class="title"> <span></span><strong>��������</strong></div>
    <%call LinkList()%>
  </div> 
</div>
<div class="main">
	<div class="border">
		<div class="title">
			<span></span><strong>��������</strong></div>
            <table width="978" bgcolor="#EBF709">
                    <tr>
                      <td width="150"><a href="../list/?43.html">&nbsp;����</a></td>
                      <td width="150"><a href="../list/?44.html">&nbsp;��������</a></td>
                      <td width="150"><a href="../list/?45.html">&nbsp;���Ź���</a></td>
                      <td width="150"><a href="../list/?46.html">&nbsp;���߽��</a></td>
                      <td width="150"><a href="../list/?47.html">&nbsp;���ұ�׼</a></td>
                      <td width="150"><a href="../list/?48.html">&nbsp;��ҵ��׼</a></td>
                  
                    </tr>
                    <tr>
                      
   
                    </tr>
                  </table>
		
	</div>
</div>

<div class="main" style="margin-top:;">
  <div class="gear border">
    <div class="wzzb-bj">
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="90%" align="center"><table align="center" width="95%" border="0">
                    <tr>
                      <td><a href="../list/?28.html">&nbsp;�����豸</a></td>
                      <td><a href="../list/?29.html">&nbsp;ͨ�š��ź�װ��</a></td>
                      <td><a href="../list/?30.html">&nbsp;�����豸</a></td>
                      <td><a href="../list/?32.html">&nbsp;��׻��߼�����</a></td>
                      <td><a href="../list/?33.html">&nbsp;��������</a></td>
                      <td><a href="../list/?34.html">&nbsp;�����������豸</a></td>
                    </tr>
                    <tr>
                      <td><a href="../list/?36.html">&nbsp;��������</a></td>
                      <td><a href="../list/?37.html">&nbsp;ͨ�硢����װ��</a></td>
                      <td><a href="../list/?38.html">&nbsp;֧���豸</a></td>
                      <td><a href="../list/?39.html">&nbsp;��ȼ�������Ʒ</a></td>
                      <td><a href="../list/?40.html">&nbsp;�ɾ��е</a></td>
                      <td><a href="../list/?41.html">&nbsp;���������</a></td>
                    </tr>
                  </table>
                  &nbsp;</td>
              </tr>
            </table></td>
        </tr>
        <tr>
          <td height="5"></td>
        </tr>
        <tr>
        
        </tr>
      </table>
      
        <div style=" width:958px;">
            
            <div style="float:left;width:945px;overflow:hiddenr;">
              <marquee direction="left" scrollamount="4" scrolldelay="1" onmouseover="this.stop()" onmouseout="this.start()" border="0" width="100%" height="140px"><%call HotPic()%></marquee>
              <div style="clear: both;"> </div>
            </div>
          </div>
    </div>
  </div>
  </td>
  </tr>
  </table>
</div>
</div>
<div class="main">
  <div class="footer">&nbsp;<br />
    Copyright &copy; 2008-2015 �������Ƽ�ѧԺ-������ʡ����ú�绯���̼����з�ƽ̨. All rights reserved.&nbsp;&nbsp;&nbsp;<a href="admin/index.asp">�����½</a></div><br />
    <div align="center">
    <% 
dim count 
Set fs=CreateObject("scripting.filesystemobject") 
Set hs=fs.opentextfile(server.Mappath("count.txt")) 
count=hs.readline 

if session("iscount")="" then 
session("iscount")="iscount" 
count=count+1 
end if 

response.write "���ǵ�" & count & "λ�����ߣ�" 
Set hs=fs.createtextfile(server.Mappath("count.txt")) 
hs.writeline(count) 
hs.close 
set fs=nothing 

%> </div>
</div>
</div>

</body>
</html>
