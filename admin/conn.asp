<%
dim rs,sql

dim conn
dim connstr
dim db
db="../data/#$%~~data.mdb"      '���ݿ��ļ���λ��
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(db)
conn.Open connstr

sub CloseConn()
	conn.close
	set conn=nothing
end sub
%><!-- #include file="../inc/function.asp" -->


