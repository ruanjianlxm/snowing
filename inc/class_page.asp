<%
'===================================================================
'XDOWNPAGE   ASP�汾
'�汾   1.00
'Code by  zykj2000
'Email:   zykj_2000@163.net
'BBS:   http://bbs.513soft.net
'������������ʹ�á��޸ģ�ϣ���ҵĳ�����Ϊ���Ĺ�����������
'���뱣��������Ϣ
'
'�����ص�
'��������Ҫ�Ƕ����ݷ�ҳ�Ĳ��ֽ����˷�װ����������ʾ������ȫ���û��Զ��壬
'֧��URL�������
'
'ʹ��˵��
'�������˵��
'PapgeSize      �����ҳÿһҳ�ļ�¼��
'GetRS       ���ؾ�����ҳ��Recordset������ֻ��
'GetConn      �õ����ݿ�����
'GetSQL       �õ���ѯ���
'���򷽷�˵��
'ShowPage      ��ʾ��ҳ������,Ψһ�Ĺ��÷���
'
'���޸ĳ�α��̬��ҳ��by ���ػ� 2010-06-28
'===================================================================

Const Btn_First=" ��ҳ "  '�����һҳ��ť��ʾ��ʽ
Const Btn_Prev="ǰһҳ "  '����ǰһҳ��ť��ʾ��ʽ
Const Btn_Next=" ��һҳ "  '������һҳ��ť��ʾ��ʽ
Const Btn_Last=" ���һҳ "  '�������һҳ��ť��ʾ��ʽ
Const XD_Align="Center"     '�����ҳ��Ϣ���뷽ʽ
Const XD_Width="100%"     '�����ҳ��Ϣ���С
Const Max_page=10

Class Xdownpage
Private XD_PageCount,XD_Conn,XD_Rs,XD_SQL,XD_PageSize,Str_errors,int_curpage,str_URL,int_totalPage,int_totalRecord,XD_sURL


'=================================================================
'PageSize ����
'����ÿһҳ�ķ�ҳ��С
'=================================================================
Public Property Let PageSize(int_PageSize)
 If IsNumeric(Int_Pagesize) Then
  XD_PageSize=CLng(int_PageSize)
 Else
  str_error=str_error & "PageSize�Ĳ�������ȷ"
  ShowError()
 End If
End Property
Public Property Get PageSize
 If XD_PageSize="" or (not(IsNumeric(XD_PageSize))) Then
  PageSize=10     
 Else
  PageSize=XD_PageSize
 End If
End Property

Public Property Get RecordCount
 RecordCount=int_totalRecord
End Property

'=================================================================
'GetRS ����
'���ط�ҳ��ļ�¼��
'=================================================================
Public Property Get GetRs()
 Set XD_Rs=Server.createobject("adodb.recordset")
 XD_Rs.PageSize=PageSize
 XD_Rs.Open XD_SQL,XD_Conn,1,1
 int_totalRecord=XD_Rs.RecordCount 
 If not(XD_Rs.eof and XD_RS.BOF) Then
  If int_curpage>XD_RS.PageCount Then
   int_curpage=XD_RS.PageCount
  End If
  XD_Rs.AbsolutePage=int_curpage
 End If
 Set GetRs=XD_RS
End Property

'================================================================
'GetConn  �õ����ݿ�����
'
'================================================================ 
Public Property Let GetConn(obj_Conn)
 Set XD_Conn=obj_Conn
End Property

'================================================================
'GetSQL   �õ���ѯ���
'
'================================================================
Public Property Let GetSQL(str_sql)
 XD_SQL=str_sql
End Property

 

'==================================================================
'Class_Initialize ��ĳ�ʼ��
'��ʼ����ǰҳ��ֵ
'
'================================================================== 
Private Sub Class_Initialize
 '========================
 '�趨һЩ�������a��ֵ
 '========================
 XD_PageSize=10  '�趨��ҳ��Ĭ��ֵΪ10
 '========================
 '��ȡ��ǰ���ֵ
 '========================
 If GetPage="" Then
  int_curpage=1
 ElseIf not(IsNumeric(GetPage)) Then
  int_curpage=1
 ElseIf CInt(Trim(GetPage))<1 Then
  int_curpage=1
 Else
  Int_curpage=CInt(Trim(GetPage))
 End If
  
End Sub

'====================================================================
'ShowPage  ������ҳ������
'����ҳ��ǰһҳ����һҳ��ĩҳ���������ֵ���
'
'====================================================================
Public Sub ShowPage()
 Dim str_tmp
 XD_sURL = GetUrl()
 If int_totalRecord<=0 Then
  str_error=str_error & "�ܼ�¼��Ϊ�㣬����������"
  Call ShowError()
  exit sub
 End If
 If int_totalRecord="" then
     int_TotalPage=1
     exit sub
 Else
  If int_totalRecord mod PageSize =0 Then
   int_TotalPage = CLng(int_TotalRecord / XD_PageSize * -1)*-1
  Else
   int_TotalPage = Fix(int_TotalRecord / XD_PageSize )+1
  End If
 End If
 
 If Int_curpage>int_Totalpage Then
  int_curpage=int_TotalPage
 End If
 
 '==================================================================
 '��ʾ��ҳ��Ϣ������ģ������Լ�Ҫ���������λ��
 '==================================================================
 response.write ""
 str_tmp=ShowFirstPrv
 response.write str_tmp
 str_tmp=showNumBtn
 response.write str_tmp
 str_tmp=ShowNextLast
 response.write str_tmp
 str_tmp=ShowPageInfo
 response.write str_tmp
 
 response.write ""
End Sub

'====================================================================
'ShowFirstPrv  ��ʾ��ҳ��ǰһҳ
'
'
'====================================================================
Private Function ShowFirstPrv()
 Dim Str_tmp,int_prvpage
 If int_curpage=1 Then
  str_tmp=Btn_First&" "&Btn_Prev
 Else
  int_prvpage=int_curpage-1
  str_tmp="<a href="""&XD_sURL & "1" & ".html"">" & Btn_First&"</a> <a href=""" & XD_sURL & CStr(int_prvpage) & ".html"">" & Btn_Prev&"</a>"
 End If
 ShowFirstPrv=str_tmp
End Function

'====================================================================
'ShowNextLast  ��һҳ��ĩҳ
'
'
'====================================================================
Private Function ShowNextLast()
 Dim str_tmp,int_Nextpage
 If Int_curpage>=int_totalpage Then
  str_tmp=Btn_Next & " " & Btn_Last
 Else
  Int_NextPage=int_curpage+1
  str_tmp="<a href=""" & XD_sURL & CStr(int_nextpage) & ".html"">" & Btn_Next&"</a> <a href="""& XD_sURL & CStr(int_totalpage) & ".html"">" &  Btn_Last&"</a>"
 End If
 ShowNextLast=str_tmp
End Function


'====================================================================
'ShowNumBtn  ���ֵ���
'
'
'====================================================================
Private Function showNumBtn()
 Dim i,str_tmp
 pre_page=int_curpage-1
 if pre_page<1 then pre_page=1
 next_page=int_curpage+1
 if next_page>int_totalpage then next_page=int_totalpage
 If int_totalpage>1 then
 str_tmp=str_tmp & "<a href=""" & XD_sURL & CStr(pre_page) & ".html""><<</a> "
 Else
 str_tmp=str_tmp & "<< "
 End if
 For i=1 to Max_page
 if i>int_totalpage then exit for
  if int_curpage>Max_page then
  if i+Max_page>int_totalpage then exit for
  if i+Max_page=int_curpage then
  str_tmp=str_tmp & " <b>"&i+Max_page&"</b> "
  else
  str_tmp=str_tmp & "[<a href=""" & XD_sURL & CStr(i+Max_page) & ".html"">"&i+Max_page&"</a>] "
  end if
  else
  if i=int_curpage then
  str_tmp=str_tmp & " <b>"&i&"</b> "
  else
  str_tmp=str_tmp & "[<a href=""" & XD_sURL & CStr(i) & ".html"">"&i&"</a>] "
  end if
  end if
 Next
 If int_totalpage>1 then
  str_tmp=str_tmp & "<a href=""" & XD_sURL & CStr(next_page) & ".html"">>></a> "
 Else
	str_tmp=str_tmp & ">> "
 End if
 showNumBtn=str_tmp
End Function


'====================================================================
'ShowPageInfo  ��ҳ��Ϣ
'����Ҫ�������޸�
'
'====================================================================
Private Function ShowPageInfo()
 Dim str_tmp
 str_tmp="ҳ��: <font color=red><b>"&int_curpage&"</b></font>/<font color=#333333><b>"&int_totalpage&"</b></font> ҳ �� <font color=#333333><b>"&int_totalrecord&"</b></font> ����¼ <b>"&XD_PageSize&"</b>��/ÿҳ"
 ShowPageInfo=str_tmp
End Function
'==================================================================
'GetURL  �õ���ǰ��URL
'����URL������ͬ����ȡ��ͬ�Ľ��
'
'==================================================================
Private Function GetURL()
 Dim strurl,str_url,i,j,search_str,result_url
 search_str="_"
 
 strurl=Request.ServerVariables("URL")
 Strurl=split(strurl,"/")
 i=UBound(strurl,1)
 str_url=""'strurl(i)'�õ���ǰҳ�ļ���
 
 str_params=Trim(Request.ServerVariables("QUERY_STRING"))
 str_params=Split(Replace(str_params,".html","")&"_","_")(0)

 If str_params="" Then
  result_url=str_url & "?_"
 Else
  If InstrRev(str_params,search_str)=0 Then
   result_url=str_url & "?" & str_params &"_"
  Else
   j=InstrRev(str_params,search_str)-2
   If j=-1 Then
    result_url=str_url & ""
   Else
    str_params=Left(str_params,j)
    result_url=str_url & "?" & str_params &"_"
   End If
  End If
 End If
 If result_url="" Then result_url="?"
 GetURL=result_url
End Function

Private Function GetPage()
dim m_querystring
m_querystring=split(replace(Request.ServerVariables("QUERY_STRING"),".html","")&"_" ,"_")
GetPage=m_querystring(1)

End function
'====================================================================
' ���� Terminate �¼���
'
'====================================================================
Private Sub Class_Terminate  
 if XD_RS.State <>0 then XD_RS.Close 
 Set XD_RS=nothing
End Sub
'====================================================================
'ShowError  ������ʾ
'
'
'====================================================================
Private Sub ShowError()
 If str_Error <> "" Then
  Response.Write("" & str_Error & "")
  Response.End
 End If
End Sub
End class
%>
