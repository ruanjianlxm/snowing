<%
'===================================================================
'XDOWNPAGE   ASP版本
'版本   1.00
'Code by  zykj2000
'Email:   zykj_2000@163.net
'BBS:   http://bbs.513soft.net
'本程序可以免费使用、修改，希望我的程序能为您的工作带来方便
'但请保留以上请息
'
'程序特点
'本程序主要是对数据分页的部分进行了封装，而数据显示部份完全由用户自定义，
'支持URL多个参数
'
'使用说明
'程序参数说明
'PapgeSize      定义分页每一页的记录数
'GetRS       返回经过分页的Recordset此属性只读
'GetConn      得到数据库连接
'GetSQL       得到查询语句
'程序方法说明
'ShowPage      显示分页导航条,唯一的公用方法
'
'已修改成伪静态分页类by 罗胸怀 2010-06-28
'===================================================================

Const Btn_First=" 首页 "  '定义第一页按钮显示样式
Const Btn_Prev="前一页 "  '定义前一页按钮显示样式
Const Btn_Next=" 下一页 "  '定义下一页按钮显示样式
Const Btn_Last=" 最后一页 "  '定义最后一页按钮显示样式
Const XD_Align="Center"     '定义分页信息对齐方式
Const XD_Width="100%"     '定义分页信息框大小
Const Max_page=10

Class Xdownpage
Private XD_PageCount,XD_Conn,XD_Rs,XD_SQL,XD_PageSize,Str_errors,int_curpage,str_URL,int_totalPage,int_totalRecord,XD_sURL


'=================================================================
'PageSize 属性
'设置每一页的分页大小
'=================================================================
Public Property Let PageSize(int_PageSize)
 If IsNumeric(Int_Pagesize) Then
  XD_PageSize=CLng(int_PageSize)
 Else
  str_error=str_error & "PageSize的参数不正确"
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
'GetRS 属性
'返回分页后的记录集
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
'GetConn  得到数据库连接
'
'================================================================ 
Public Property Let GetConn(obj_Conn)
 Set XD_Conn=obj_Conn
End Property

'================================================================
'GetSQL   得到查询语句
'
'================================================================
Public Property Let GetSQL(str_sql)
 XD_SQL=str_sql
End Property

 

'==================================================================
'Class_Initialize 类的初始化
'初始化当前页的值
'
'================================================================== 
Private Sub Class_Initialize
 '========================
 '设定一些参数的黙认值
 '========================
 XD_PageSize=10  '设定分页的默认值为10
 '========================
 '获取当前面的值
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
'ShowPage  创建分页导航条
'有首页、前一页、下一页、末页、还有数字导航
'
'====================================================================
Public Sub ShowPage()
 Dim str_tmp
 XD_sURL = GetUrl()
 If int_totalRecord<=0 Then
  str_error=str_error & "总记录数为零，请输入数据"
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
 '显示分页信息，各个模块根据自己要求更改显求位置
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
'ShowFirstPrv  显示首页、前一页
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
'ShowNextLast  下一页、末页
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
'ShowNumBtn  数字导航
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
'ShowPageInfo  分页信息
'更据要求自行修改
'
'====================================================================
Private Function ShowPageInfo()
 Dim str_tmp
 str_tmp="页次: <font color=red><b>"&int_curpage&"</b></font>/<font color=#333333><b>"&int_totalpage&"</b></font> 页 共 <font color=#333333><b>"&int_totalrecord&"</b></font> 条记录 <b>"&XD_PageSize&"</b>条/每页"
 ShowPageInfo=str_tmp
End Function
'==================================================================
'GetURL  得到当前的URL
'更据URL参数不同，获取不同的结果
'
'==================================================================
Private Function GetURL()
 Dim strurl,str_url,i,j,search_str,result_url
 search_str="_"
 
 strurl=Request.ServerVariables("URL")
 Strurl=split(strurl,"/")
 i=UBound(strurl,1)
 str_url=""'strurl(i)'得到当前页文件名
 
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
' 设置 Terminate 事件。
'
'====================================================================
Private Sub Class_Terminate  
 if XD_RS.State <>0 then XD_RS.Close 
 Set XD_RS=nothing
End Sub
'====================================================================
'ShowError  错误提示
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
