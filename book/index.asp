<!--#include file="../conn.asp"-->
<!--#include file="../inc/class_page.asp" -->
<%
dim m_querystring,cid,id
m_querystring=split(Split(replace(Request.ServerVariables("QUERY_STRING"),".html","") & "_","_")(0) & "-" ,"-")
cid=m_querystring(0)
id=m_querystring(1)
cpage=Split(replace(Request.ServerVariables("QUERY_STRING"),".html","") & "_","_")(1)
%>
<%
Select Case request("action")
	Case "save"
	    if trim(Session("safecode")) <> trim(Request("Code")) then 
            ErrorMessage = "��������ȷ����֤��" 
            response.write(" <script>alert('"&ErrorMessage&"');location.href='../book' </script>") 
            response.end 
        end if 
		Dim ad_bGuest,ad_bIpAddress,ad_bContent,ad_bReply
		ad_bGuest=request("bGuest")	
		ad_bIpAddress=GetIPAddress()
		ad_bContent=request("bContent")	
		ad_bReply="�ȴ��ظ�"		
		set rs=server.createobject("adodb.recordset")
		sql="select top 1 * from Ay_Book "	
		rs.open sql,conn,1,3	
		rs.addnew
		rs("bGuest")=ad_bGuest	
		rs("bIpAddress")=ad_bIpAddress
		rs("bContent")=ad_bContent		
		rs("bIsPass")=1		
		rs("bReply")=ad_bReply
		rs("bAddTime")=now			
		rs.update
		if rs.state<>0 then rs.close
		set rs=nothing
		response.write "<script>alert('������ӳɹ�����ȴ�����Ա�ظ���');window.location.href='../book';</script>"
		response.end	
End select
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

    <script language="javascript">
function check()
{
if (document.chatform.bGuest.value=="")
{
	alert("���������Ĵ�����");
	document.chatform.bGuest.focus();
	return false;
}
if (document.chatform.bContent.value=="")
{
	alert("�������������ݣ�");
	document.chatform.bContent.focus();
	return false;
}
if(document.chatform.Code.value=="")
{
    alert('��֤��Ϊ�գ�');
    return false;
}  
}
    </script>

</head>
<body>
    <div class="head"></div>
    <div id="menu_out">
        <div id="menu_in">
            <div id="menu">
                <%call HeadNavigation()%>
            </div>
        </div>
    </div>
    <div class="navdh">
        <strong>��ǰλ�ã�&nbsp;<a href="../">��վ��ҳ</a> &gt;&nbsp;��̳ </strong>
    </div>
    <div class="main">
      <div class="channel_left">
            <div class="border">
                <div class="title">
                    <strong>��̳</strong>
                </div>
                <ul class="list_list">
                    <%
			dim curpage
			
			sql="select a.* from Ay_Book a order by a.bId desc,bAddTime desc"
			Set rs=Server.CreateObject("Adodb.Recordset")
			Set mypage=new xdownpage
			mypage.getconn=conn
			mypage.getsql=sql
			mypage.pagesize=4
			set rs=mypage.getrs()	
			if cpage<>"" then
				curpage=clng(cpage+0)
			else
				curpage=1
			end if					
			for i=1 to mypage.pagesize
				if rs.eof  then exit for										
                    %>
                    <div id="book">
                        <div class="gtitle">
                            <span>ʱ�䣺<%=FormatDate(rs("bAddTime"),"1")%>
<%=Trim(rs("bIpAddress")&"")%></span> <font color="#ff0000">
                                    <%=Trim(rs("bGuest")&"")%>
                                </font>���ԣ�<font color="#3b78af"><%=Trim(rs("bTitle")&"")%></font>
                        </div>
                        <div class="gcontent">
                            <%=replace(Trim(rs("bContent")&""),vbcrlf,"<br>")%>
                        </div>
                        <div class="greply">
                            <%=replace(Trim(rs("bReply")&""),vbcrlf,"<br>")%>
                        </div>
                    </div>
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
            <div class="border mt8">
                <div class="title">&nbsp;��̳ | ����ò����������ԣ�</div>
                <table width="100%" border="0" cellspacing="2" cellpadding="2">
                    <form name="chatform" id="chatform" onsubmit="return check();" action="../book/?action=save"
                        method="post">
                        <tr>
                            <td width="60" align="center">
                                ������</td>
                            <td>
                                <input type="text" name="bGuest" value>&nbsp;&nbsp; ��֤�룺<input type="text" name="Code"
                                    id="Code" size="5"><span style="background: #FFFFFF; padding: 3px;"><img id="CodeImage"
                                        src="../inc/code.asp" align="absmiddle"></span></td>
                        </tr>
                        <tr>
                            <td align="center">
                                ���ݣ�</td>
                            <td>
                                <textarea name="bContent" cols="60" rows="6"></textarea></td>
                        </tr>
                        <tr>
                            <td>
                            </td>
                            <td>
                                <input type="submit" value="�ύ����"></td>
                        </tr>
                    </form>
                </table>
            </div>
            <div style="clear: both;">
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
