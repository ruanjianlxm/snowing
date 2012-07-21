<%
dim rs,sql
on error resume next
dim conn
dim connstr
dim db
db="data/#$%~~data.mdb"      '数据库文件的位置
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
conn.Open connstr
if err.number<>0 then
		err.clear
		db="../data/#$%~~data.mdb"
		connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
		conn.Open connstr
end if
sub CloseConn()
	conn.close
	set conn=nothing
end sub
%>
<!-- #include file="inc/function.asp" -->
<!-- #include file="inc/public.asp" -->
<!-- #include file="inc/sql.asp" -->
<%
dim ay_sitename
dim ay_sitetitle
dim ay_sitedomain
dim ay_miibeian
dim ay_telphone
dim ay_email
dim ay_address
dim ay_author
dim ay_keywords
dim ay_description
dim ay_replacewords
dim ay_maxonline
dim ay_needpass
msql="select * from Ay_System"
Set rss=Server.CreateObject("Adodb.Recordset")
rss.Open msql, Conn
if not rss.eof then
	
	ay_sitename=trim(rss("bName") & "")
	ay_sitetitle=trim(rss("bTitle") & "")
	ay_sitedomain=trim(rss("bUrl") & "")
	
	ay_miibeian=trim(rss("bMiibeian") & "")
	ay_telphone=trim(rss("bPhone") & "")
	ay_email=trim(rss("bEmail") & "")
	ay_address=trim(rss("bAddress") & "")
	ay_author=trim(rss("bAuthor") & "")
	ay_keywords=trim(rss("bKeywords") & "")
	ay_description=trim(rss("bDescriptions") & "")
	ay_replacewords=trim(rss("bReplacewords") & "")
	ay_maxonline=trim(rss("bMaxOnline")&"")
	ay_needpass=trim(rss("bNeedPass")&"")
end if
if ay_maxonline="" then ay_maxonline="0"

if rss.state<>0 then rss.close
set rss=nothing
%>


