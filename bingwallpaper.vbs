'On error resume next 'comment it for debug
set FSO=CreateObject ("Scripting.FileSystemObject")
bingfile = fso.GetSpecialFolder(2): if right(bingfile,1)<>"\" then bingfile=bingfile & "\" : bingfile = bingfile & "bing.jpg"

sUrlRequest = "https://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1"
Set oXMLHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
oXMLHTTP.Open "GET", sUrlRequest, False
oXMLHTTP.Send
xmlfile=oXMLHTTP.Responsetext
Set oXMLHTTP = Nothing

beg=instr(lcase(xmlfile),"<urlbase>")
ef=instr(lcase(xmlfile),"</urlbase>")
lnk=mid(xmlfile,beg+9,ef-beg-9)
url="http://www.bing.com/"+lnk+"_1920x1080.jpg"

Set oXMLHTTP2 = CreateObject("WinHttp.WinHttpRequest.5.1")
oXMLHTTP2.Open "GET", url, False
oXMLHTTP2.Send
Set oADOStream = CreateObject("ADODB.Stream")
oADOStream.Mode = 3 'RW
oADOStream.Type = 1 'Binary
oADOStream.Open
oADOStream.Write oXMLHTTP2.ResponseBody

oADOStream.SaveToFile bingfile, 2

Set objWshShell = WScript.CreateObject("Wscript.Shell")
'use OS to set wallpaper
'objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\Wallpaper", bingfile, "REG_SZ"
'objWshShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, False
Set objFile = FSO.GetFile(bingfile)
Set objShellApp = CreateObject("Shell.Application")
Set objFolder = objShellApp.Namespace(FSO.GetParentFolderName(objFile))
objFolder.ParseName(FSO.GetFileName(objFile)).InvokeVerb "setdesktopwallpaper"


'use irfanview if you want
'objWshShell.Run "c:\Programs\IrfanView\i_view64.exe """ & bingfile & """ /wall=0 /killmesoftly", 1, False 


Set oXMLHTTP2 = Nothing
Set oADOStream = Nothing
Set FSO = Nothing
Set objWshShell = Nothing
