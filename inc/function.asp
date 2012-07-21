<%
Private Function GetIPAddress()
    Dim sIPAddress, sHTTP_X_FORWARDED_FOR
    sHTTP_X_FORWARDED_FOR = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
    If sHTTP_X_FORWARDED_FOR = "" or InStr(sHTTP_X_FORWARDED_FOR, "unknown") > 0 Then
        sIPAddress = Request.ServerVariables("REMOTE_ADDR")
    ElseIf InStr(sHTTP_X_FORWARDED_FOR, ",") > 0 Then
        sIPAddress = Mid(sHTTP_X_FORWARDED_FOR, 1, InStr(sHTTP_X_FORWARDED_FOR, ",") -1)
    ElseIf InStr(sHTTP_X_FORWARDED_FOR, ";") > 0 Then
        sIPAddress = Mid(sHTTP_X_FORWARDED_FOR, 1, InStr(sHTTP_X_FORWARDED_FOR, ";") -1)
    Else
        sIPAddress = sHTTP_X_FORWARDED_FOR
    End If
    GetIPAddress= Trim(Mid(sIPAddress, 1, 15))
End Function

'***************************************************
	'函数名：IsObjInstalled
	'作  用：检查组件是否已经安装
	'参  数：strClassString ----组件名
	'返回值：True  ----已经安装
	'       False ----没有安装
	'***************************************************
	Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False
		Err = 0
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True
		Set xTestObj = Nothing
		Err = 0
	End Function
	'================================================
 '函数名：FormatDate
 '作  用：格式化日期
 '参  数：DateAndTime   ----原日期和时间
 '        para   ----日期格式
 '返回值：格式化后的日期
 '================================================
 Public Function FormatDate(DateAndTime, para)
  On Error Resume Next
  Dim y, m, d, h, mi, s, strDateTime
  FormatDate = DateAndTime
  If Not IsNumeric(para) Then Exit Function
  If Not IsDate(DateAndTime) Then Exit Function
  y = CStr(Year(DateAndTime))
  m = CStr(Month(DateAndTime))
  If Len(m) = 1 Then m = "0" & m
  d = CStr(Day(DateAndTime))
  If Len(d) = 1 Then d = "0" & d
  h = CStr(Hour(DateAndTime))
  If Len(h) = 1 Then h = "0" & h
  mi = CStr(Minute(DateAndTime))
  If Len(mi) = 1 Then mi = "0" & mi
  s = CStr(Second(DateAndTime))
  If Len(s) = 1 Then s = "0" & s
  Select Case para
  Case "1"
   strDateTime = y & "-" & m & "-" & d & " " & h & ":" & mi & ":" & s
  Case "2"
   strDateTime = y & "-" & m & "-" & d
  Case "3"
   strDateTime = y & "/" & m & "/" & d
  Case "4"
   strDateTime = y & "年" & m & "月" & d & "日"
  Case "5"
   strDateTime = m & "-" & d & " " & h & ":" & mi
  Case "6"
   strDateTime = m & "/" & d
  Case "7"
   strDateTime = m & "月" & d & "日"
  Case "8"
   strDateTime = y & "年" & m & "月"
  Case "9"
   strDateTime = y & "-" & m
  Case "10"
   strDateTime = y & "/" & m
  Case "11"
   strDateTime = right(y,2) & "-" &m & "-" & d & " " & h & ":" & mi
  Case "12"
   strDateTime = right(y,2) & "-" &m & "-" & d
  Case "13"
   strDateTime = m & "-" & d
  Case "14"
		select case CStr(Month(DateAndTime))
			case "1"
			strDateTime="January"
			case "2"
			strDateTime="February"
			case "3"
			strDateTime="March"
			case "4"
			strDateTime="April"
			case "5"
			strDateTime="May"
			case "6"
			strDateTime="June"
			case "7"
			strDateTime="July"
			case "8"
			strDateTime="August"
			case "9"
			strDateTime="September"
			case "10"
			strDateTime="October"
			case "11"
			strDateTime="November"
			case else
			strDateTime="December"
		end Select		
	Case "15"
		strDateTime=d
	Case "16"
   		strDateTime = d & "日"

  Case Else
   strDateTime = DateAndTime
  End Select
 FormatDate = strDateTime
End Function

Function GotTopic(Str,StrLen) 
Dim l,t,c, i,LableStr,regEx,Match,Matches,focus,last_str 
if IsNull(Str) then 
GotTopic = "" 
Exit Function 
end if 
if Str = "" then 
GotTopic="" 
Exit Function 
end if 
Set regEx = New RegExp 
regEx.Pattern = "\[[^\[\]]*\]" 
regEx.IgnoreCase = True 
regEx.Global = True 
Set Matches = regEx.Execute(Str) 
For Each Match in Matches 
LableStr = LableStr & Match.Value 
Next 
Str = regEx.Replace(Str,"") 
Str=Replace(Replace(Replace(Replace(Str," "," "),"'",Chr(34)),">",">"),"<","<")
l=len(str) 
t=0 
strlen=Clng(strLen) 
for i=1 to l 
c=Abs(Asc(Mid(str,i,1))) 
if c>255 then 
t=t+2 
else 
t=t+1 
end if 
if t = strLen-2 then 
focus = i 
last_str = ".." 
end if 
if t = strLen-1 then 
focus = i 
last_str = "." 
end if 
if t>=strlen then 
GotTopic=left(str,focus)&last_str 
exit for 
else 
GotTopic=str 
end if 
next 
GotTopic = Replace(Replace(Replace(Replace(GotTopic," "," "),Chr(34),"'"),">",">"),"<","<") & LableStr 
end Function

Function IsValidEmail(email)
	IsValidEmail = True
	names = Split(email, "@")
	IF UBound(names) <> 1 THEN
	IsValidEmail = false
	Exit Function
	End IF
	For each name in names
	IF Len(name) <= 0 THEN
	IsValidEmail = false
	Exit Function
	End IF
	For i = 1 To Len(name)
	c = Lcase(Mid(name, i, 1))
	IF InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and Not IsNumeric(c) THEN
	IsValidEmail = false
	Exit Function
	End IF
	Next
	IF Left(name, 1) = "." or Right(name, 1) = "." THEN
	IsValidEmail = false
	Exit Function
	End IF
	Next
	IF InStr(names(1), ".") <= 0 THEN
	IsValidEmail = false
	Exit Function
	End IF
	i = Len(names(1)) - InStrRev(names(1), ".")
	IF i <> 2 and i <> 3 THEN
	IsValidEmail = false
	Exit Function
	End IF
	IF InStr(email, "..") > 0 THEN
	IsValidEmail = false
	End IF
End Function

Function isInteger(para)
	on error resume Next
	Dim str
	Dim l,i
	IF isNUll(para) THEN 
	isChkInteger=false
	Exit Function
	End IF
	str=cstr(para)
	IF trim(str)="" THEN
	isChkInteger=false
	Exit Function
	End IF
	l=len(str)
	For i=1 To l
	IF mid(str,i,1)>"9" or mid(str,i,1)<"0" THEN
	isChkInteger=false 
	Exit Function
	End IF
	Next
	isChkInteger=True
	IF err.number<>0 THEN err.clear
End Function
Function HTMLEncode(fString)
IF Not isnull(fString) THEN
fString = replace(fString, ">", "&gt;")
fString = replace(fString, "<", "&lt;")
fString = Replace(fString, CHR(32), "&nbsp;")
fString = Replace(fString, CHR(9), "&nbsp;")
fString = Replace(fString, CHR(34), "&quot;")
fString = Replace(fString, CHR(39), "&#39;")
fString = Replace(fString, CHR(13), "")
fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
fString = Replace(fString, CHR(10), "<BR> ")
HTMLEncode = fString
End IF
End Function
Function HTMLCode(fString)
IF Not isnull(fString) THEN
fString = replace(fString, "&gt;", ">")
fString = replace(fString, "&lt;", "<")
fString = Replace(fString,  "&nbsp;"," ")
fString = Replace(fString, "&quot;", CHR(34))
fString = Replace(fString, "&#39;", CHR(39))
fString = Replace(fString, "</P><P> ",CHR(10) & CHR(10))
fString = Replace(fString, "<BR> ", CHR(10))
HTMLCode = fString
End IF
End Function
Function RemoveHTML(strHTML)
dim tmpstring
tmpstring=strHTML
Dim objRegExp, Match, Matches
Set objRegExp = New Regexp
objRegExp.IgnoreCase = True
objRegExp.Global = True
'取闭合的<>
objRegExp.Pattern = "<.+?>"
'进行匹配
Set Matches = objRegExp.Execute(tmpstring)
' 遍历匹配集合，并替换掉匹配的项目
For Each Match in Matches
tmpstring=Replace(tmpstring,Match.Value,"")
Next
RemoveHTML=tmpstring
Set objRegExp = Nothing
End Function


Function DeleteHtml(strHtml)

Dim objRegExp, strOutput
Set objRegExp = New Regexp ' 建立正则表达式

objRegExp.IgnoreCase = True ' 设置是否区分大小写
objRegExp.Global = True '是匹配所有字符串还是只是第一个
objRegExp.Pattern = "(<[a-zA-Z].*?>)|(<[\/][a-zA-Z].*?>)" ' 设置模式引号中的是正则表达式，用来找出html标签

strOutput = objRegExp.Replace(strHtml, "") '将html标签去掉
strOutput = Replace(strOutput, "<", "<") '防止非html标签不显示
strOutput = Replace(strOutput, ">", ">") 
delHtml = strOutput

Set objRegExp = Nothing
End Function

Function Jencode(byVal iStr)
if isnull(iStr) or isEmpty(iStr) then
  Jencode=""
  Exit function
end if
dim F,i,E

  E=array("Jn0;","Jn1;","Jn2;","Jn3;","Jn4;","Jn5;","Jn6;","Jn7;","Jn8;","Jn9;","Jn10;","Jn11;","Jn12;","Jn13;","Jn14;","Jn15;","Jn16;","Jn17;","Jn18;","Jn19;","Jn20;","Jn21;","Jn22;","Jn23;","Jn24;","Jn25;")
  F=array(chrw(12468),chrw(12460),chrw(12462),chrw(12464),_
    chrw(12466),chrw(12470),chrw(12472),chrw(12474),_
    chrw(12485),chrw(12487),chrw(12489),chrw(12509),_
    chrw(12505),chrw(12503),chrw(12499),chrw(12497),_
    chrw(12532),chrw(12508),chrw(12506),chrw(12502),_
    chrw(12500),chrw(12496),chrw(12482),chrw(12480),_
    chrw(12478),chrw(12476))
  Jencode=iStr
  for i=0 to 25
   Jencode=replace(Jencode,F(i),E(i))
  next
End Function

Function Juncode(byVal iStr)
if isnull(iStr) or isEmpty(iStr) then
  Juncode=""
  Exit function
end if
dim F,i,E

E=array("Jn0;","Jn1;","Jn2;","Jn3;","Jn4;","Jn5;","Jn6;","Jn7;","Jn8;","Jn9;","Jn10;","Jn11;","Jn12;","Jn13;","Jn14;","Jn15;","Jn16;","Jn17;","Jn18;","Jn19;","Jn20;","Jn21;","Jn22;","Jn23;","Jn24;","Jn25;")
  F=array(chrw(12468),chrw(12460),chrw(12462),chrw(12464),_
    chrw(12466),chrw(12470),chrw(12472),chrw(12474),_
    chrw(12485),chrw(12487),chrw(12489),chrw(12509),_
    chrw(12505),chrw(12503),chrw(12499),chrw(12497),_
    chrw(12532),chrw(12508),chrw(12506),chrw(12502),_
    chrw(12500),chrw(12496),chrw(12482),chrw(12480),_
    chrw(12478),chrw(12476))
  Juncode=iStr
for i=0 to 25
  Juncode=replace(Juncode,E(i),F(i))
next
End Function

'通过正则表达式提取文章内容中的图片标签，实现内容与图片的分离。
Function RegExp_Execute(patrn,strng) 
Dim regEx,Match,Matches,values '建立变量 
Set regEx=New RegExp '建立正则表达式 
regEx.Pattern = Patrn '设置模式 
regEx.IgnoreCase = true '设置是否区分字符大小写 
regEx.Global = true '设置全局可用性 
Set Matches=regEx.Execute(strng) '执行搜索 
For Each Match in Matches '遍历匹配集合 
Values=values&Match.value&"," 
Next 
RegExp_Execute=values 
End Function 


''定义函数将关键字替换为红色
Function Search(strChar,strWords) 
If strChar = "" Or IsNull(strChar) Then 
Search = "" 
Exit Function  
End If 
Dim strTChar, arrTChar, tempChar, i 
strTChar = strWords
arrTChar = Split(strTChar, ",") 
tempChar = strChar 
For i = 0 To UBound(arrTChar) 
tempChar = Replace(tempChar, arrTChar(i), "<font color=#ff0000>"&strWords&"</font>") 
Next 
Search = tempChar 
End Function 
''函数结束


Function GetEditorImg(str)
dim tmp
    Set objRegExp = New Regexp
     objRegExp.IgnoreCase = True    
     objRegExp.Global = false    
     objRegExp.Pattern = "<img (.*?)src=(.[^\[^>]*)(.*?)>"
    Set Matches =objRegExp.Execute(str)
    For Each Match in Matches
         tmp=tmp & Match.Value
    Next
     GetEditorImg=GetEditorAllImg(tmp)
end Function

function GetEditorAllImg(str)
    Set objRegExp1 = New Regexp
     objRegExp1.IgnoreCase = True    
     objRegExp1.Global = True    
    objRegExp1.Pattern = "src\=.+?\.(gif|jpg|png|bmp)"
    set mm=objRegExp1.Execute(str)
    For Each Match1 in mm
         imgsrc=Match1.Value
         imgsrc=replace(imgsrc,"""","")
         imgsrc=replace(imgsrc,"src=","")
         imgsrc=replace(imgsrc,"<","")
         imgsrc=replace(imgsrc,">","")
         imgsrc=replace(imgsrc,"img","")
         imgsrc=replace(imgsrc," ","")
         GetEditorAllImg=GetEditorAllImg&imgsrc
    next
end function

%>

