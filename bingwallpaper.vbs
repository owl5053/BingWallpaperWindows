set FSO=CreateObject ("Scripting.FileSystemObject")
bingfile = fso.GetSpecialFolder(2): if right(bingfile,1)<>"\" then bingfile=bingfile & "\" : bingfile = bingfile & "bing.jpg"

sUrlRequest = "https://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1"
Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP")
oXMLHTTP.Open "GET", sUrlRequest, False
oXMLHTTP.Send
xmlfile=oXMLHTTP.Responsetext
Set oXMLHTTP = Nothing

beg=instr(lcase(xmlfile),"<urlbase>")
ef=instr(lcase(xmlfile),"</urlbase>")
lnk=mid(xmlfile,beg+9,ef-beg-9)
url="http://www.bing.com/"+lnk+"_1920x1080.jpg"

Set oXMLHTTP2 = CreateObject("MSXML2.XMLHTTP")
oXMLHTTP2.Open "GET", url, False
oXMLHTTP2.Send
Set oADOStream = CreateObject("ADODB.Stream")
oADOStream.Mode = 3 'RW
oADOStream.Type = 1 'Binary
oADOStream.Open
oADOStream.Write oXMLHTTP2.ResponseBody

oADOStream.SaveToFile bingfile, 2

Set objWshShell = WScript.CreateObject("Wscript.Shell")
objWshShell.RegWrite "HKEY_CURRENT_USER\Control Panel\Desktop\Wallpaper", bingfile, "REG_SZ"
objWshShell.Run "%windir%\System32\RUNDLL32.EXE user32.dll,UpdatePerUserSystemParameters", 1, False

Set oXMLHTTP2 = Nothing
Set oADOStream = Nothing
Set FSO = Nothing
Set objWshShell = Nothing