<!--#include file="inc/conn.asp"--><%host=lcase(request.servervariables("HTTP_HOST"))%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<meta name="viewport" content="width=device-width,minimum-scale=1.0,maximum-scale=1.0"/>
<meta name="apple-mobile-web-app-capable" content="yes" />
<title><%=title%></title>
<script type="text/javascript" src="inc/js/Ajax_Web.js"></script>
<link href="inc/css/New_Web.css" rel="stylesheet" type="text/css" />
<body onLoad="inst();">
<div class="sub_bod"></div>
<div class="sub_top">
 <div class="title"><%=title%></div>
 <!---div class="back" id="pageback"><a href="index.asp" class="d">����</a></div---> 
 <!---<div class="menu" id="topmenus"><a href="admin/" class="d">��̨</a></div>--->
</div><div class="main">
<%
startime=timer()
duansss="���������"&duan
mamasss="������4������֤��"

'#	����1:ֱ��ͨ���趨�ģ�1-3������ѯ������ѯ
'#	��Ѽ�ʱ��ͨ:http://add.12391.net/ 
'#	��Ƶ�̳�����:http://pan.baidu.com/s/1eSoDn26 
'#	���빺��:https://item.taobao.com/item.htm?id=528692002051 
'#	�������:https://item.taobao.com/item.htm?id=520023732507 

names=""&trim(request("name"))&""
codes=""&trim(request("code"))&""
times=""&trim(request("time"))&""

newtime = DateDiff("s","1970-1-1 0:0:0",Now())
if times="" then
%><form name="queryForm" method="post" class="" action="?a=<%=newtime%>" onsubmit="return startRequest(0);">
<div class="select" id="10">
<select name="time" id="time" onBlur="startRequest(1)" />
<%
'���ú������������� 
Dim rs 
Set rs = getSortedFiles(Server.MapPath(""&UpDir&"/")) 
While Not rs.EOF
 EditFile=rs("Name") 
if right(EditFile,len(mdbtype))=mdbtype then
 ii=ii+1
 timefile=left(EditFile,len(EditFile)-len(mdbtype))
 response.write "<option value="""&timefile&""">"&timefile&"</option>"&vbcrlf
end if
 rs.MoveNext 
Wend 
if ii<1 then
 response.write "<option value="""">����Ա��δ�ϴ�����Ŷ</option>"&vbcrlf
end if
rs.Close 
Set rs = Nothing 
%></select></div>
<div class="so_box" id="11">
<input name="name" type="text" class="txts" id="name" value="<%=duansss%>" placeholder="<%=duansss%>" onfocus="st('name',1)" onBlur="startRequest(2)" />
</div>
<%if yanzhenma="1" then
mamas="+��֤��"%>
<div class="so_box" id="33">
<input name="code" type="text" class="txts" id="code" value="<%=mamasss%>" placeholder="<%=mamasss%>" onfocus="this.value=''" onBlur="startRequest(3)" />
<div class="more" id="clearkey">
<img src="inc/Code.asp?t=<%=timer%>" id="Codes" onClick="this.src='inc/Code.asp?t='+new Date();"/>
</div></div><%end if%>
<div class="so_boxes">
<input type="submit" name="button" class="buts" id="sub" value="������ѯ" />
</div>
<div class="so_boxex" id="tishi"><strong>���˵��: </strong>
<%=duan%><%=mamas%>��������ȷ����ʾ��Ӧ�����
<!---�������˵����������ӣ���ʼ--->
<!---�������˵����������ӣ�����--->
</div>
<div id="tishi1" style="display:none;"><%=duansss%></div>
<div id="tishi4" style="display:none;"><%=mamasss%></div>
</form>
<%
else
datas=""&UpDir&"/"&times&""&mdbtype

if yanzhenma="1" then
if len(codes)<>4 or codes<>Session("GetCode") Then
 call AlertBack("��������ȷ����֤��Ŷ��") 
End if
end if

if filekey(times)>0 then
 call AlertBack("���������ִ���")
end if

if len(names)<1 and len(names)>18 Then
 call AlertBack("������"&duan&"��") 
End if

if IsFile(datas)=True then
else
 call AlertBack("������ʱû���ϴ����߲�����Ŷ��") 
end if

set conn=server.createobject("adodb.connection")
conn.open "driver={Microsoft Excel Driver (*.xls)};DBQ="&server.mappath(datas)
set rst=conn.OpenSchema(20)
first_table=rst("Table_Name").Value
Response.Write "<!---"&first_table&"--->"&vbcrlf
Response.Write "<!---"&duan&"='"&names&"'--->"&vbcrlf
Response.Write "<!--startprint--><table cellspacing=""0"">"&vbcrlf
 Response.Write "<caption align='center'>"&times&" ��ѯ���</caption>"&vbcrlf
set rst=nothing
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from ["&first_table&"] where "&duan&"='"&names&"'",conn,1,1
Response.Write "<tr class=""tt"">"&vbcrlf
    For x3=0 To rs.Fields.Count-1
       tnames=tnames&"---"&rs.Fields(x3).Name&"---"
       Response.Write "<td>" & rs.Fields(x3).Name & "</td>"&vbcrlf
    Next
Response.Write "</tr>"&vbcrlf
if instr(tnames,"---"&tiaojian1&"---")>0 then
else
 call AlertBack("���������õĲ�ѯ����["&tiaojian1&"]�Ƿ���ڣ�") 
end if
do while not rs.eof
lis = rs.Fields.Count
 yyy=yyy+1
 Response.Write "<tr>"&vbcrlf
For rx=0 To lis-1
 curValue = trim(rs.Fields(rx).Value)
 If IsNull(curValue) or len(curValue)<1 Then
 curValue="&nbsp;"
 End If
 curValue = CStr(curValue)
 z=rx+1
ll=rs.Fields(rx).Name
 'Response.Write "<td class=""r"">" & ll & "</td>"&vbcrlf
 Response.Write "<td class=""span"">" & curValue & "</td>"&vbcrlf
Next
 Response.Write "</tr>"&vbcrlf
rs.movenext
loop
rs.close
if yyy<1 then
 'Response.Write "<caption align='center'> ��ѯʧ�� </caption>"&vbcrlf
 Response.Write "<tr>"&vbcrlf
 Response.Write "<td colspan="""&lis&""" align=""center"">"&vbcrlf
 Response.Write "<span>��ѯʧ��: ��ѯ��������ڣ�</span></td>"&vbcrlf
 Response.Write "</tr>"&vbcrlf
end if

Response.Write "</table>"&vbcrlf
Response.Write "<!--endprint-->"&vbcrlf
%><div class="so_boxexe">
<input type="button" name="print" value="Ԥ������ӡ" onclick="preview()" class="buts">
<input type="button" value="�� ��" class="buts" onclick="location.href='index.asp';" id="reset"></div>
<%end if
endtime=timer()%></div>
<div class="foot">
 <div class="title">
 <span>&copy;<%=year(now)%>&nbsp; <a href="<%=copysu%>" target="_blank"><%=copysr%></a>
 <!---����汾�뿴:http://test.96448.cn/--->
 <a href="http://aiyaha.taobao.com/" target="_blank">����֧��</a></span>
<!---

ҳ��ִ��ʱ�䣺<%=FormatNumber((endtime-startime)*1000,3)%>����

��ر���������Ϣ��

��ѿ�ͨ��http://add.12391.net [����С�û����]
�޳�������yujianyue@126.com (֧�����˺�)
���ø��ѣ�https://aiyaha.taobao.com
Ӧ����ϵ��15058593138 �����ϱ���+�������ʣ�

--->
 </div>
</div>
</body>
</html><script type="text/javascript" src="../index_cha.js?v=ADA_W3T"></script>