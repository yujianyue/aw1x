<%
'修改以下几项后 无需修改其他内容


title="某某da中学成绩查询系统"			'设置查询标题,相信你懂的。

copysr="某某da中学"				'设置底部版权简短的文字,相信你懂的。
copysu="http://www.96448.cn/"			'设置底部版权连接完整网址,相信你懂的。

'设置查询条件
tiaojian1="姓名"				'汉字是列标题，跟excel一致

yanzhenma="0"					'是否使用验证码。填1使用0不用
UpDir="shujukufangzheli"			'设置数据库所在目录(文件夹名称),修改后相应更名对于文件夹。

'#########################################################
'修改以上几项后 无需修改以下内容
'#########################################################

'以下内容是系列查询系统导航，以后总能用到，敬请保留

'#	通用成绩查询系统解决方案(简单通用易用):
'#	
'#	方案4:微信公众号N选1个查询条件直接查询工资、成绩、水电费等
'#	自助开通试用:http://new.12391.net/ 
'#	视频教程下载:http://pan.baidu.com/s/1ge6BPEr 
'#	代码购买:https://item.taobao.com/item.htm?id=520496908275 
'#	整体服务:https://item.taobao.com/item.htm?id=529624346797 
'#	
'#	方案3(荐):微信公众号一对一绑定才可以查询工资、成绩、水电费等
'#	自助开通试用:http://add.96cha.com/ 
'#	代码购买:https://item.taobao.com/item.htm?id=44248394675 
'#	整体服务:https://item.taobao.com/item.htm?id=528187132312 
'#	
'#	方案2(荐):用户在线登录查询工资成绩水电费等，可自助修改密码
'#	自助开通试用:http://add.dbcha.com/ 
'#	视频教程下载:http://pan.baidu.com/s/1boANMwv 
'#	代码购买:https://item.taobao.com/item.htm?id=43193387085 
'#	整体服务:https://item.taobao.com/item.htm?id=528108807297 
'#	
'#	方案1:直接通过设定的（1-3个）查询条件查询
'#	免费即时开通:http://add.12391.net/ 
'#	视频教程下载:http://pan.baidu.com/s/1eSoDn26 
'#	代码购买:https://item.taobao.com/item.htm?id=528692002051 
'#	整体服务:https://item.taobao.com/item.htm?id=520023732507 

'#	代码版：不加密，无域名限制，无时间限制,一次付费一直可用(域名和网站空间费用另外自理)
'#	整体服务：无需域名 无需空间 无需代码 无需技术人员 无需备案，即开即用 ,按时间付费

'#	通用模糊检索系统解决方案(简单通用易用):

'#	方案1:通用多选一模糊查询系统单输入框版
'#	自助开通试用:http://add.xuelikai.com:1111/ 
'#	视频教程下载:http://pan.baidu.com/s/1ge6BPEr (只参考第一步第二步)
'#	代码购买:https://item.taobao.com/item.htm?id=520167788658 
'#	整体服务:https://item.taobao.com/item.htm?id=529624346797 

'#	方案2:通用多选一模糊查询系统下拉可选条件版
'#	自助开通试用:http://add.xuelikai.com:2222/
'#	视频教程下载:http://pan.baidu.com/s/1ge6BPEr (只参考第一步第二步)
'#	代码购买:https://item.taobao.com/item.htm?id=520167788658 
'#	整体服务:https://item.taobao.com/item.htm?id=528187132312 

'#	方案3(荐):通用多选一模糊查询系统多输入框版
'#	自助开通试用:http://add.xuelikai.com:3333/
'#	视频教程下载:http://pan.baidu.com/s/1ge6BPEr (只参考第一步第二步)
'#	代码购买:https://item.taobao.com/item.htm?id=520167788658 
'#	整体服务:https://item.taobao.com/item.htm?id=528108807297 
'#	
'#	以上三个模糊检索方案也可以用于多选一精准查询系统

'#	成绩工资水电费通用查询系统其他方案:
'#	
'#	50元:asp无需后台版(12款选1):
'#	https://item.taobao.com/item.htm?id=45703415332 
'#	
'#	60元:PHP无需后台版(36款选1):
'#	https://item.taobao.com/item.htm?id=45808268273 
'#	
'#	108元:PHP多级下拉版(8款选1):
'#	https://item.taobao.com/item.htm?id=43263796985 
'#	
'#	N选1模糊查询系统解决方案(超过8款68元起):
'#	https://item.taobao.com/item.htm?id=520167788658 

'#	N选1精准查询系统解决方案(超过8款68元起):
'#	https://item.taobao.com/item.htm?id=520167788658
'#	
'#	其他相关查询系统:

'#	网店发货查询系统后台导入通用版网页版加微信自动回复机器人查询
'#	即时开通试用:http://add.mabida.cn:1111/ 
'#	整体服务:https://item.taobao.com/item.htm?id=528166721835 
'#	源码购买:https://item.taobao.com/item.htm?id=44194950836 
'#	
'#	农产品质量溯源查询系统二维码防伪查询微信防伪查询
'#	即时开通试用:http://add.mabida.cn/ 
'#	整体服务:https://item.taobao.com/item.htm?id=43525422046 
'#	源码购买:https://item.taobao.com/item.htm?id=43525422046 

duan=tiaojian1
mdbtype=".xls" '只能是xls格式文件哦不要修改

if len(tiaojian1)=2 then
 qianmian1=left(tiaojian1,1)&"&nbsp;&nbsp;"&right(tiaojian1,1)
else
 qianmian1=tiaojian1
end if

Function IsFile(FilePath)
 Set Fso=Server.CreateObject("Scri"&"pting.File"&"Sys"&"temObject")
 If (Fso.FileExists(Server.MapPath(FilePath))) Then
 IsFile=True
 Else
 IsFile=False
 End If
 Set Fso=Nothing
End Function

Function filekey(texts)
 filekey=0
 rekey="-/-\-%-@-.-"
 keyes=split(rekey,"-")
 nnnnn=Ubound(keyes)
 For m=1 To Ubound(keyes)-1
 rekeys=keyes(m)
 rekeys=trim(rekeys)
 if instr(texts,rekeys)>0 and len(rekeys)>0 then
 filekey=filekey+1
 end if
 next
End Function

'定义获取排序文件列表的函数 
Function getSortedFiles(folderPath) 
 Dim rs, fso, folder, File 
 Const adInteger = 3 
 Const adDate = 7 
 Const adVarChar = 200 
 Set rs = Server.CreateObject("ADODB.Recordset") 
 Set fso = Server.CreateObject("Scripting.FileSystemObject") 
 Set folder = fso.GetFolder(folderPath) 
 Set fso = Nothing 
 With rs.Fields 
 .Append "Name", adVarChar, 200 
 .Append "Type", adVarChar, 200 
 .Append "DateCreated", adDate 
 .Append "DateLastAccessed", adDate 
 .Append "DateLastModified", adDate 
 .Append "Size", adInteger 
 .Append "TotalFileCount", adInteger 
 End With 
 rs.Open 
 For Each File In folder.Files 
 rs.AddNew 
 rs("Name") = File.Name 
 rs("Type") = File.Type 
 rs("DateCreated") = File.DateCreated 
 rs("DateLastAccessed") = File.DateLastAccessed 
 rs("DateLastModified") = File.DateLastModified 
 rs("Size") = File.Size 
 rs.Update 
 Next 
 '设置排序规则：按名称排序 
 rs.Sort = "DateLastModified DESC" 
 '设置排序规则：依次按文件大小倒序，按修改日期倒序 
 'rs.Sort = "Size DESC, DateLastModified DESC" 
 rs.MoveFirst 
 Set folder = Nothing 
 Set getSortedFiles = rs 
End Function 
'==============================
'函 数 名： AlertUrl(AlertStr,Url) 
'作 用：警告后转入指定页面
'==============================
Function AlertUrl(AlertStr,Url) 
 Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" &vbcrlf
 Response.Write "<script>" &vbcrlf
 Response.Write "alert('"&AlertStr&"');" &vbcrlf
 Response.Write "location.href='"&Url&"';" &vbcrlf
 Response.Write "</script>" &vbcrlf
 Response.End()
End Function
'==============================
'函 数 名： AlertBack(AlertStr)
'作 用：警告后返回上一页面
'==============================
Function AlertBack(AlertStr) 
 Response.Write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"" />" &vbcrlf
 Response.Write "<script>" &vbcrlf
 Response.Write "alert('"&AlertStr&"');" &vbcrlf
 Response.Write "history.go(-1)" &vbcrlf
 Response.Write "</script>"&vbcrlf
 Response.End()
End Function

%>