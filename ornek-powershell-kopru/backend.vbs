' ============================================================
'  backend.vbs  -  VBScript is mantigi (COM / WMI / FileSystemObject / ADODB)
'
'  Cagri  :  cscript //nologo backend.vbs <ciktiDosyasi> <action> [p1] [p2]
'  Sonuc  :  UTF-8 (BOM'suz) JSON olarak <ciktiDosyasi>'na yazilir.
'
'  Bu dosya, modern tarayicidan gelen istekleri server.ps1 (PowerShell)
'  koprusu uzerinden alir. Sizin tum VBScript yatiriminiz (ADODB, Access,
'  dosya islemleri...) burada yasamaya devam eder.
'
'  NOT: Hardcoded metinler bilerek ASCII tutuldu (cscript ANSI okuyabilir).
'       Kullanicidan gelen veriler (Turkce dahil) sorunsuz islenir.
' ============================================================
Option Explicit

Dim args, outFile, action, p1, p2
Set args = WScript.Arguments
If args.Count < 2 Then WScript.Quit 1

outFile = args(0)
action  = LCase(args(1))
p1 = "" : p2 = ""
If args.Count > 2 Then p1 = args(2)
If args.Count > 3 Then p2 = args(3)

' --- Ortak nesneler / veri deposu (FSO ornegi) ---
Dim fso, dataDir, dataFile, tmpTs
Set fso = CreateObject("Scripting.FileSystemObject")
dataDir = fso.BuildPath(fso.GetParentFolderName(WScript.ScriptFullName), "data")
If Not fso.FolderExists(dataDir) Then fso.CreateFolder dataDir
dataFile = fso.BuildPath(dataDir, "users.txt")
If Not fso.FileExists(dataFile) Then
    Set tmpTs = fso.CreateTextFile(dataFile, True, True)   ' unicode = True
    tmpTs.Close
End If

' --- Yonlendirme (whitelist mantigi: sadece bilinen action'lar) ---
Dim result
Select Case action
    Case "info"  : result = GetInfo()
    Case "list"  : result = ListUsers()
    Case "add"   : result = AddUser(p1, p2)
    Case "del"   : result = DelUser(p1)
    Case "files" : result = ListFiles()
    Case Else    : result = "{""error"":""bilinmeyen action: " & J(action) & """}"
End Select

WriteResult result
WScript.Quit 0

' ============================================================
'  Yardimcilar
' ============================================================

' JSON string kacisi
Function J(s)
    Dim t : t = CStr(s)
    t = Replace(t, "\",   "\\")
    t = Replace(t, """",  "\""")
    t = Replace(t, vbCr,  "\r")
    t = Replace(t, vbLf,  "\n")
    t = Replace(t, vbTab, "\t")
    J = t
End Function

' Sonucu UTF-8 (BOM'suz) olarak cikti dosyasina yazar  (ADODB.Stream).
' Boylece Turkce karakterler tarayiciya bozulmadan ulasir.
Sub WriteResult(text)
    If outFile = "" Then
        WScript.StdOut.Write text   ' (yedek: dogrudan calistirildiysa)
        Exit Sub
    End If
    Dim s1, bytes, s2
    Set s1 = CreateObject("ADODB.Stream")
    s1.Type = 2 : s1.Charset = "utf-8" : s1.Open   ' adTypeText
    s1.WriteText text
    s1.Position = 0
    s1.Type = 1                                     ' adTypeBinary
    s1.Position = 3                                 ' UTF-8 BOM (3 byte) atla
    bytes = s1.Read
    s1.Close

    Set s2 = CreateObject("ADODB.Stream")
    s2.Type = 1 : s2.Open
    s2.Write bytes
    s2.SaveToFile outFile, 2                        ' adSaveCreateOverWrite
    s2.Close
End Sub

' --- COM / WMI ornegi: sistem bilgisi ---
Function GetInfo()
    Dim net, comp, usr, os, wmi, item
    Set net = CreateObject("WScript.Network")
    comp = net.ComputerName
    usr  = net.UserName
    os   = ""
    On Error Resume Next
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    For Each item In wmi.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")
        os = item.Caption : Exit For
    Next
    On Error GoTo 0
    GetInfo = "{""computer"":""" & J(comp) & """,""user"":""" & J(usr) & _
              """,""os"":""" & J(os) & """,""engine"":""VBScript / cscript.exe""}"
End Function

' --- FSO ornegi: kullanicilari oku ---
Function ListUsers()
    Dim ts, line, parts, json, first
    json = "[" : first = True
    Set ts = fso.OpenTextFile(dataFile, 1, False, -1)   ' ForReading, Unicode
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        If Len(Trim(line)) > 0 Then
            parts = Split(line, "|")
            If Not first Then json = json & ","
            json = json & "{""id"":" & CLng(parts(0)) & _
                   ",""name"":""" & J(parts(1)) & """,""role"":""" & J(parts(2)) & """}"
            first = False
        End If
    Loop
    ts.Close
    ListUsers = json & "]"
End Function

' --- FSO ornegi: kullanici ekle (yaz) ---
Function AddUser(name, role)
    name = Trim(name)
    If Len(name) = 0 Then AddUser = "{""error"":""ad bos olamaz""}" : Exit Function
    If Len(Trim(role)) = 0 Then role = "Kullanici"
    Dim ts, line, parts, maxId
    maxId = 0
    Set ts = fso.OpenTextFile(dataFile, 1, False, -1)   ' ForReading, Unicode
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        If Len(Trim(line)) > 0 Then
            parts = Split(line, "|")
            If CLng(parts(0)) > maxId Then maxId = CLng(parts(0))
        End If
    Loop
    ts.Close
    Set ts = fso.OpenTextFile(dataFile, 8, True, -1)    ' ForAppending, Unicode
    ts.WriteLine (maxId + 1) & "|" & Replace(name, "|", "/") & "|" & Replace(role, "|", "/")
    ts.Close
    AddUser = "{""ok"":true,""id"":" & (maxId + 1) & "}"
End Function

' --- FSO ornegi: kullanici sil ---
Function DelUser(idStr)
    Dim ts, line, parts, keep, id
    id = CLng("0" & Trim(idStr))
    keep = ""
    Set ts = fso.OpenTextFile(dataFile, 1, False, -1)   ' ForReading, Unicode
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        If Len(Trim(line)) > 0 Then
            parts = Split(line, "|")
            If CLng(parts(0)) <> id Then keep = keep & line & vbCrLf
        End If
    Loop
    ts.Close
    Set ts = fso.OpenTextFile(dataFile, 2, True, -1)    ' ForWriting, Unicode
    ts.Write keep
    ts.Close
    DelUser = "{""ok"":true}"
End Function

' --- FSO ornegi: bu klasordeki dosyalari listele ---
Function ListFiles()
    Dim folder, f, json, first
    Set folder = fso.GetFolder(fso.GetParentFolderName(WScript.ScriptFullName))
    json = "[" : first = True
    For Each f In folder.Files
        If Not first Then json = json & ","
        json = json & "{""name"":""" & J(f.Name) & """,""size"":" & f.Size & "}"
        first = False
    Next
    ListFiles = json & "]"
End Function

' ============================================================
'  GERCEK PROJEDE: Access/ADODB ornegi (sablon)
'  FSO deposu yerine kendi veritabaninizi kullanmak icin yukaridaki
'  ListUsers/AddUser yerine asagidakine benzer kod yazin:
'
'  Function ListUsersDB()
'      Dim conn, rs, json, first
'      Set conn = CreateObject("ADODB.Connection")
'      conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
'                fso.BuildPath(dataDir, "app.mdb")
'      Set rs = conn.Execute("SELECT id, name, role FROM Users")
'      json = "[" : first = True
'      Do Until rs.EOF
'          If Not first Then json = json & ","
'          json = json & "{""id"":" & rs("id") & ",""name"":""" & _
'                 J(rs("name")) & """,""role"":""" & J(rs("role")) & """}"
'          first = False : rs.MoveNext
'      Loop
'      conn.Close
'      ListUsersDB = json & "]"
'  End Function
'
'  ONEMLI: Jet 4.0 (.mdb) yalnizca 32-bit'tir. 64-bit Windows'ta
'  server.ps1 icindeki $cscript degerini SysWOW64'teki 32-bit surume
'  ayarlayin:  $cscript = "$env:WINDIR\SysWOW64\cscript.exe"
'  (veya .accdb + ACE OLEDB 64-bit saglayicisini kurun.)
' ============================================================
