<%
public sub HeadNavigation()
%>
<ul id="nav">
	<li><a href="../" class="nav_on"><span>&nbsp;&nbsp;&nbsp;首页&nbsp;&nbsp;&nbsp;</span></a></li>
	<%
	set rss=server.CreateObject("adodb.recordset")
	msql="select * from Ay_Class where bPosition=1 order by bOrder asc "					
	rss.open msql,conn	
	do while not rss.eof 	
		response.write "<li class=""menu_line""></li>"	
		response.write "<li><a href='../list/?" & trim(rss("bId")&"") & ".html'"      
		response.write "><span>&nbsp;&nbsp;&nbsp;"&rss("bName")&"&nbsp;&nbsp;&nbsp;</span></a></li>"	 
		rss.movenext
	loop
	if rss.state<>0 then rss.close
	set rss=nothing
	%>
	
	
	<li>
	
	</li><li class="menu_line"></li>
    <li><a href="../book" class="nav_on"><span>&nbsp;&nbsp;&nbsp;&nbsp;论坛&nbsp;&nbsp;&nbsp;&nbsp;</span></a></li><li class="menu_line"></li>
    <li><a href="../list?18.html" class="nav_on"><span>&nbsp;&nbsp;联系我们</span></a></li>
</ul>
<%
end sub
%>
<%
Public Sub Footer()
%>

<div class="footer">&nbsp;<br />
    Copyright &copy; 2008-2015 黑龙江科技学院-<%=ay_sitename%> All rights reserved.&nbsp;&nbsp;&nbsp;<a href="../admin/index.asp">管理登陆</a></div></p>


<%
Call CloseConn()
End Sub 
%>

<%
Public Sub FlashNews()
dim mvarpics,mvarlinks,mvartexts			
set rss=server.CreateObject("adodb.recordset")
msql="select top 8 * from Ay_Content where bClassID=26 and bPic<>'' order by bId desc"
rss.open msql,conn															
do while not rss.eof
	mvarpics= mvarpics & Trim(rss("bPic")&"") & "|"
	mvarlinks= mvarlinks & "show/?" & Trim(rss("bClassID")&"") & "-" & Trim(rss("bId")&"") & ".html" & "|"
	mvartexts= mvartexts & GotTopic(Trim(rss("bTitle")&""),32) & "|"							
	rss.movenext
loop
if mvarpics<>"" then mvarpics=left(mvarpics,len(mvarpics)-1)
if mvarlinks<>"" then mvarlinks=left(mvarlinks,len(mvarlinks)-1)
if mvartexts<>"" then mvartexts=left(mvartexts,len(mvartexts)-1)
If rss.state<>0 Then rss.close
Set rss=nothing
%>
<script language="javascript">
	var swf_width=240;
	var swf_height=201;
	var configtg='0xffffff|0|0x4AC6BB|5|0xffffff|0x0EA094|0x000033|3|2|1|_blank';	
	var files="<%=mvarpics%>";
	var links="<%=mvarlinks%>";
	var texts="<%=mvartexts%>";
	document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
	document.write('<param name="movie" value="images/bcastr3.swf"><param name="quality" value="high">');
	document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
	document.write('<param name="FlashVars" value="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'">');
	document.write('<embed src="images/bcastr3.swf" wmode="opaque" FlashVars="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'&menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>');
</script>
<%
End Sub 
%>
<%
public sub Top_News()
set rss=server.CreateObject("adodb.recordset")
msql="select top 1 * from Ay_Content where bClassID=26 order by bAddTime desc,bId "					
rss.open msql,conn	
if not rss.eof  then
%>
<div class="top_title">
	<a href="show/?<%=trim(rss("bClassID") & "")%>-<%=trim(rss("bId") & "")%>.html" target="_blank"><%=GotTopic(trim(rss("bTitle") & ""),40)%></a>
</div>
<div class="top_con">
<%=GotTopic(RemoveHTML(trim(rss("bContent") & "")),126)%>
</div>
<%
end if
if rss.state<>0 then rss.close
set rss=nothing
%>
<ul class="top_list">
	<%
	set rss=server.CreateObject("adodb.recordset")
	msql="select top 4 * from Ay_Content where bClassID=26 order by bAddTime desc,bId "					
	rss.open msql,conn	
	do while not rss.eof 
	%>
	<li><span>[<%=FormatDate(trim(rss("bAddTime") & ""),13)%>]</span>
	<a href="show/?<%=trim(rss("bClassID") & "")%>-<%=trim(rss("bId") & "")%>.html" target="_blank"><%=GotTopic(trim(rss("bTitle") & ""),40)%></a></li>		
	<%
		rss.movenext
	loop
	if rss.state<>0 then rss.close
	set rss=nothing
	%>	
</ul>
<%
end sub
%>
<%
public sub Notice()
%>
<div style="padding-top:50px;">
	<marquee direction="up" scrollamount="2" scrolldelay="5" onmouseover="this.stop()" onmouseout="this.start()" border="0" height="230px" width="100%">
	<ul class="none_list">
		<%
		set rss=server.CreateObject("adodb.recordset")
		msql="select top 10 * from Ay_Notice order by bOrder,bAddTime desc,bId "					
		rss.open msql,conn	
		do while not rss.eof 
		%>
		<li>
		<a href="<%=trim(rss("bUrl") & "")%>" target="_blank"><%=GotTopic(trim(rss("bTitle") & ""),28)%></a>
		</li>		
		<%
			rss.movenext
		loop
		if rss.state<>0 then rss.close
		set rss=nothing
		%>
	</ul></marquee>
</div>
<%
end sub
%>
<%
public sub HotNews()
%>
<div class="gd_center">
	<ul class="gd_list" id="marqueebox0">
		<%
		set rss=server.CreateObject("adodb.recordset")
		msql="select top 20 * from Ay_Content  order by bClick desc,bAddTime desc,bId "					
		rss.open msql,conn	
		do while not rss.eof 
		%>
		<li>
		<a href="../show/?<%=trim(rss("bClassID") & "")%>-<%=trim(rss("bId") & "")%>.html" target="_blank"><%=GotTopic(trim(rss("bTitle") & ""),26)%></a>
		</li>		
		<%
			rss.movenext
		loop
		if rss.state<>0 then rss.close
		set rss=nothing
		%>
	</ul>
	<script language="javascript" type="text/javascript" src="images/gd.js"></script>
</div>
<%
end sub
%>
<%
public sub SearchForm()
%>
<form action="../search/" method="post" name="frmsearch" id="frmsearch">
	<input type="hidden" name="kwtype" value="0" />
	<input name="keyword" type="text" class="input" id="keyword" value="输入您要查询的关键字" onblur="if(this.value==''){this.value='输入您要查询的关键字';}this.style.color='#666';" onclick="if(this.value=='输入您要查询的关键字'){this.value='';}this.style.color='#666';" />
	<input name="button" type="image" src="../images/sub.gif" class="btn" id="button" value="搜索" onclick="if(frmsearch.keyword.value=='' || frmsearch.keyword.value=='输入您要查询的关键字'){alert('请输入搜索关键词');frmsearch.keyword.value='';frmsearch.keyword.focus();frmsearch.keyword.style.color='#666';return false;}this.form.submit();"
 />
	<span><b>热门搜索
	<%
	set rss=server.CreateObject("adodb.recordset")
	msql="select top 4 * from Ay_Search order by bClick desc,bAddTime desc"					
	rss.open msql,conn
	Do While Not rss.eof
	%>
	<a href="../search/?keyword=<%=Trim(rss("bKeywords")&"")%>"><%=Trim(rss("bKeywords")&"")%></a> 	
	<%
		rss.movenext
	Loop
	If rss.state<>0 Then rss.close
	Set rss=Nothing
	%>
	</b></span>
</form>
<%
end sub
%>
<%
public sub Index_Class(cls_id)
	set rss=server.CreateObject("adodb.recordset")
	msql="select  * from Ay_Class where bId=" & cls_id					
	rss.open msql,conn
	if not rss.eof then
%>
<div class="title">
	<span><a href="../list/?<%=trim(rss("bId") & "")%>.html">more</a> </span>
	<strong><a href="../list/?<%=trim(rss("bId") & "")%>.html"><%=trim(rss("bName") & "")%></a></strong>
</div>
<%
	end if
	if rss.state<>0 then rss.close
	set rss=nothing
%>
<div style="clear:both;"></div>
<ul class="text_list">
	<%
	
	set rss=server.CreateObject("adodb.recordset")
	msql="select top 8 * from Ay_Content_v where bClassID=" & cls_id & " or bParentID=" & cls_id & " order by bAddTime desc,bId "					
	rss.open msql,conn	
	do while not rss.eof 
	%>
	<li>
	<a href="show/?<%=trim(rss("bClassID") & "")%>-<%=trim(rss("bId") & "")%>.html" target="_blank"><%if (counter>9) then response.write " <span><img src='../images/new.jpg' /></span>" end if%><%=GotTopic(trim(rss("bTitle") & ""),38)%></a>
	</li>		
	<%
		rss.movenext
		
	loop
	if rss.state<>0 then rss.close
	set rss=nothing
	%>

</ul>
<%
end sub
%>
<%
public sub HotPic()
set rss=server.CreateObject("adodb.recordset")
msql="select top 12 * from Ay_Content_v where bClassID=41  and bPic<>''  or bClassID=40 and bPic<>''  or bClassID=39 and bPic<>''  or bClassID=38 and bPic<>''  or bClassID=37 and bPic<>''  or bClassID=36 and bPic<>''  or bClassID=34 and bPic<>''  or bClassID=33 and bPic<>''  or bClassID=32 and bPic<>''  or bClassID=30 and bPic<>''  or bClassID=29 and bPic<>''  or bClassID=28 and bPic<>''  order by bClick desc,bAddTime desc,bId "					
rss.open msql,conn	
do while not rss.eof 
%>
<div class="hot_pic">
<a href="../show/?<%=trim(rss("bClassID") & "")%>-<%=trim(rss("bId") & "")%>.html" target="_blank">
<img src="<%=trim(rss("bPic") & "")%>" border="0" width="134" height="104" alt="<%=trim(rss("bTitle") & "")%>"></a>
<span><a href="../show/?<%=trim(rss("bClassID") & "")%>-<%=trim(rss("bId") & "")%>.html" target="_blank"><%=GotTopic(trim(rss("bTitle") & ""),18)%></a></span>
</div>			
<%
	rss.movenext
loop
if rss.state<>0 then rss.close
set rss=nothing
end sub
%>
<%
public sub LinkList()
%>
<ul class="link">
	<%
	set rss=server.CreateObject("adodb.recordset")
	msql="select top 8 * from Ay_Link order by bAddTime desc,bId "					
	rss.open msql,conn	
	do while not rss.eof 
	%>
	<li style="float:left; width:200px; overflow:hidden;">
	<a href="<%=trim(rss("bUrl") & "")%>" target="_blank"><%=GotTopic(trim(rss("bName") & ""),28)%></a>	
	</li>	
	<%
		rss.movenext
	loop
	if rss.state<>0 then rss.close
	set rss=nothing
	%>
</ul>
<%
end sub
%>
<%
public sub NavClass(cid)
%>
<ul class="lgx_list">
	<%
	dim pid
	set rss=server.CreateObject("adodb.recordset")
	msql="select a.* from Ay_Class a where a.bParent in (select iif(t.bParent=0,t.bId,t.bParent) as bId from Ay_Class t where t.bId=" & cid & ")"
	msql=msql & " order by bParent,bOrder"
	rss.open msql,conn		
	if not rss.eof then
		if pid="" then pid=trim(rss("bParent") & "")
	end if
	%>
	<li>
	<a href="../list/?<%=pid%>.html">全部分类</a>
	</li>
	<%
	do while not rss.eof 
		
	%>
	<li>
	<a href="../list/?<%=trim(rss("bId") & "")%>.html"><%=trim(rss("bName") & "")%></a>
	</li>
	<%
		rss.movenext
	loop
	if rss.state<>0 then rss.close
	set rss=nothing
	%>
</ul>
<%
end sub
%>
<%
public sub ChannelNav(cid)
	if cid<>"" then
		set rss=server.CreateObject("adodb.recordset")	
		msql="select * from ("	
		msql=msql & "SELECT a.* FROM Ay_Class AS a where a.bId=" & cid 
		msql=msql & " union "
		msql=msql & " (select b.* from Ay_Class AS a INNER JOIN Ay_Class AS b ON a.bParent = b.bId where a.bId=" & cid & ")"
		msql=msql & " ) order by bParent,bOrder"
		rss.open msql,conn		
		do while not rss.eof
%>
		<a href="../list/?<%=trim(rss("bId")&"")%>.html"><%=trim(rss("bName")&"")%></a>&gt;
<%
			rss.movenext
		loop
		if rss.state<>0 then rss.close
		set rss=nothing	
	end if
end sub
%>
<%
public sub Pic_Text()
set rss=server.CreateObject("adodb.recordset")
msql="select top 2 * from Ay_Content_v where bPic<>'' and bIsBest=1 order by bClick desc,bAddTime desc,bId "					
rss.open msql,conn	
do while not rss.eof 
%>
<li>
<span class="co">
<a href="../show/?<%=trim(rss("bClassID") & "")%>-<%=trim(rss("bId") & "")%>.html" target="_blank"><%=GotTopic(trim(rss("bTitle") & ""),22)%></a><br />
<%=GotTopic(RemoveHTML(trim(rss("bContent") & "")),60)%>
</span>
<a href="../show/?<%=trim(rss("bClassID") & "")%>-<%=trim(rss("bId") & "")%>.html" target="_blank"><img src="../<%=trim(rss("bPic") & "")%>" border="0" width="70" height="80" alt="<%=trim(rss("bTitle") & "")%>"></a>
</li>
<%
	rss.movenext
loop
if rss.state<>0 then rss.close
set rss=nothing
end sub
%>
<%
public sub Text_List()
	set rss=server.CreateObject("adodb.recordset")
	msql="select top 8 * from Ay_Content_v where bIsBest=1 order by bAddTime desc,bId "					
	rss.open msql,conn	
	do while not rss.eof 
	%>
	<li>
	<a href="../show/?<%=trim(rss("bClassID") & "")%>-<%=trim(rss("bId") & "")%>.html" target="_blank"><%=GotTopic(trim(rss("bTitle") & ""),32)%></a>
	</li>		
	<%
		rss.movenext
	loop
	if rss.state<>0 then rss.close
	set rss=nothing
end sub
%><%
Sub ShowPrev(ClassID,Article_id)
	set rss=server.CreateObject("adodb.recordset")
    msql = "Select top 1 * from Ay_Content Where bId < "& Article_ID &" And bClassID = "& ClassID &" order by bId desc,bAddTime"
    rss.open msql,conn
    If Not rss.Eof Then
        Response.Write "<a href=""../show/?"& trim(rss("bClassID")&"") & "-" & trim(rss("bId")&"") & ".html"">"&trim(rss("bTitle") & "")&"</a>" 
    Else
        Response.Write "已经是最前一篇了"
    End If
    if rss.state<>0 then rss.close
    Set rss = Nothing
End Sub

Sub ShowNext(ClassID,Article_id)
    set rss=server.CreateObject("adodb.recordset")
    msql = "Select top 1 * from Ay_Content Where bId > "& Article_ID &" And bClassID = "& ClassID &" order by bId asc,bAddTime"
    rss.open msql,conn
    If Not rss.Eof Then
        Response.Write "<a href=""../show/?" & trim(rss("bClassID")&"") & "-" & trim(rss("bId")&"") & ".html"">"&trim(rss("bTitle") & "")&"</a>" 
    Else
        Response.Write "已经是最后一篇了"
    End If
    if rss.state<>0 then rss.close
    Set rss = Nothing
End Sub
%>


<%
Public Sub QiyeFlash()
dim mvarpics,mvarlinks,mvartexts			
set rss=server.CreateObject("adodb.recordset")
msql="select top 8 * from Ay_Content where bClassID=49 and bPic<>'' order by bId desc"
rss.open msql,conn															
do while not rss.eof
	mvarpics= mvarpics & Trim(rss("bPic")&"") & "|"
	mvarlinks= mvarlinks & "show/?" & Trim(rss("bClassID")&"") & "-" & Trim(rss("bId")&"") & ".html" & "|"
	mvartexts= mvartexts & GotTopic(Trim(rss("bTitle")&""),32) & "|"							
	rss.movenext
loop
if mvarpics<>"" then mvarpics=left(mvarpics,len(mvarpics)-1)
if mvarlinks<>"" then mvarlinks=left(mvarlinks,len(mvarlinks)-1)
if mvartexts<>"" then mvartexts=left(mvartexts,len(mvartexts)-1)
If rss.state<>0 Then rss.close
Set rss=nothing
%>
<script language="javascript">
	var swf_width=230;
	var swf_height=200;
	var configtg='0xffffff|0|0x4AC6BB|5|0xffffff|0x0EA094|0x000033|3|2|1|_blank';	
	var files="<%=mvarpics%>";
	var links="<%=mvarlinks%>";
	var texts="<%=mvartexts%>";
	document.write('<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0" width="'+ swf_width +'" height="'+ swf_height +'">');
	document.write('<param name="movie" value="images/bcastr3.swf"><param name="quality" value="high">');
	document.write('<param name="menu" value="false"><param name=wmode value="opaque">');
	document.write('<param name="FlashVars" value="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'">');
	document.write('<embed src="images/bcastr3.swf" wmode="opaque" FlashVars="bcastr_file='+files+'&bcastr_link='+links+'&bcastr_title='+texts+'&bcastr_config='+configtg+'&menu="false" quality="high" width="'+ swf_width +'" height="'+ swf_height +'" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />'); document.write('</object>');
</script>
<%
End Sub 
%>