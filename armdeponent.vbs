Option Explicit
Dim args, launcher, exe, WshShell, objNet, objFSO, inifile, cob
Dim inidir, ini
dim sfs_db, sfs_host, sfs_port
dim mark_db, mark_host, mark_port

Set objNet = CreateObject("WScript.Network")
Set WshShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

exe = WshShell.CurrentDirectory & "\bin\arm_deponent.exe"
inidir = WshShell.CurrentDirectory & "\bin\users"
ini = WshShell.CurrentDirectory & "\bin\users\usr.ini"

If objfso.FileExists(WshShell.CurrentDirectory & "\cob.txt") then
    sfs_db = "DBNAME1"
    sfs_host = "HOST1.NET"
    sfs_port = "10080"
    mark_db = "DBNAM2"
    mark_host = "HOST1.NET"
    mark_port = "10080"
    msgbox "Launching with CoB settings!", vbOKOnly+vbInformation, "CoB"
else
    sfs_db = "DBNAME1"
    sfs_host = "HOST2.NET"
    sfs_port = "10080"
    mark_db = "DBNAME2"
    mark_host = "HOST2.NET"
    mark_port = "10080"
end if

Set args = Wscript.Arguments
if args.count = 1 Then
    'TROPS roles
    dim tropspriv
    tropspriv = priv(3) + priv(4) + priv(5) + priv(6) + priv(7) + priv(8) + priv(9) + priv(20) + priv(21)
    'SFS roles
    dim sfspriv
    sfspriv = priv(11) + priv(12) + priv(13) + priv(14) + priv(16) + priv(16) + priv(17) + priv(18) + priv(19)
    'ISA roles
    dim isapriv
    isapriv = priv(2) + priv(10)
    if tropspriv > 0 and sfspriv > 0 then
        msgbox "More than one department is assigned!", vbOKOnly+vbCritical, "Error"
    elseif tropspriv = 0 and sfspriv = 0 then
        if isapriv > 1 then
            dim dbchose
            dbchose = dbselect()
            if dbchose = 1 then
                call launch("SFS")
            elseif dbchose = 2 then
                call launch("Markets")
            end if
        elseif isapriv = 1 then
            if priv(2) = 1 Then
                call launch("Markets")
            elseif priv(10) = 1 Then
                call launch("Markets")
            end if
        else
            msgbox "No departments assigned!", vbOKOnly+vbCritical, "Error"
        end if
    else
        if isapriv > 0 then
            msgbox "Duty segregation violation detected!", vbOKOnly+vbCritical, "Error"
        else
            if tropspriv > 0 and sfspriv = 0 then
                call launch("Markets")
            elseif tropspriv = 0 and sfspriv > 0 then
                call launch("SFS")
            end if
        end if
    end if
elseif args.count = 0 then
    msgbox "No entitlements string was recieved!", vbOKOnly+vbCritical, "Error"
else
    msgbox "Wrong arguments recieved!", vbOKOnly+vbCritical, "Error"
end if

function priv(id)
    if mid(args(0), id, 1) = 1 then
        priv = 1
    else
        priv = 0
    end if
end function

function launch(depart)
    if dircheck() = true then
        if inicreat(depart) = true then
            launcher = WshShell.Run (exe, 1, true)
            call delini()
        else
            msgbox "Failed to write configuration file!", vbOKOnly+vbCritical, "Error"
        end if
    else
        msgbox "Failed to create configuration directory!", vbOKOnly+vbCritical, "Error"
    end if
end function

function delini()
    dim source, errcount
    source = ini
    If objfso.FileExists(source) then
        On Error Resume Next
        errcount = 1
        do
            objFSO.DeleteFile source, 1
            If Err.Number <> 0 then
                errcount = errcount + 1
                if errcount = 6 then
                    exit do
                else
                    call sleep(2)
                end if
            else
                errcount = 0
                exit do
            end if
        loop until errcount = 6
        On Error GoTo 0
    end if
    source = inidir
    If objfso.FolderExists(source) then
        On Error Resume Next
        errcount = 1
        do
            objFSO.DeleteFolder source, 1
            If Err.Number <> 0 then
                errcount = errcount + 1
                if errcount = 6 then
                    exit do
                else
                    call sleep(2)
                end if
            else
                errcount = 0
                exit do
            end if
        loop until errcount = 6
        On Error GoTo 0
    end if
end function

function inicreat(depart)
    dim writelog, logtowrite, logline
    dim errcount
    On Error Resume Next
    logtowrite = ini
    errcount = 1
    do
        Set writelog=objFSO.openTextfile(logtowrite, 2, true)
        If Err.Number <> 0 then
            errcount = errcount + 1
            if errcount = 6 then
                exit do
            else
                call sleep(2)
            end if
        else
            errcount = 0
            exit do
        end if
    loop until errcount = 6
    if errcount = 6 then
        On Error GoTo 0    
        inicreat = false    
    else
        dim database, host, port
        if (depart = "SFS") Then
            database = sfs_db
            host = sfs_host
            port = sfs_port
        elseif (depart = "Markets") Then
            database = mark_db
            host = mark_host
            port = mark_port
        end if
        logline = "[DataBase]" & vbcrlf & "Database =" & database & vbcrlf & "LogID =" & username() & vbcrlf & "Host =" & host & vbcrlf & "Port =" & port & vbcrlf & "Database_connect=(DESCRIPTION =(ADDRESS = (PROTOCOL = TCP)(HOST = " & host & ")(PORT =" & port & "))(CONNECT_DATA =(SERVER = DEDICATED)(SERVICE_NAME = " & database & ")))"
        writelog.WriteLine logline
        writelog.close
        On Error GoTo 0
        inicreat = true
    end if
end function

function username()
    Dim objService, Process, Process2, strNameOfUser, strUserDomain, Return
    dim prnameuser, usercurr
    set objService = getobject("winmgmts:")
    usercurr = ""
    prnameuser = "launchingapp.exe"
    for each Process in objService.InstancesOf("Win32_process")
        If lcase(Process.name) = prnameuser Then
            return = Process.GetOwner(strNameOfUser,strUserDomain)
            usercurr = strNameOfUser
        end if
    next
    if usercurr = "" or usercurr = "\" Then
        usercurr = objnet.UserName
    end if
    username = usercurr
end function

function dircheck()
    On Error Resume Next
    dim dirtocr
    dirtocr = inidir
    If not objfso.FolderExists(dirtocr) Then
        objFSO.CreateFolder(dirtocr)
    end if
    On Error Goto 0
    If not objfso.FolderExists(dirtocr) Then
        dircheck = false
    else
        dircheck = true
    end if
end function

function Sleep(seconds)
    dim countdown
    countdown = 0
    with createobject("wscript.shell")
        do
                 .run "timeout " & 1, 0, True
            countdown = countdown + 1
        loop until countdown = seconds
    end with
End function


Function dbselect()
    ' This function uses Internet Explorer to create a dialog.
    Dim objIE, sTitle, iErrorNum
    ' Create an IE object
    Set objIE = CreateObject( "InternetExplorer.Application" )
    ' specify some of the IE window's settings
    objIE.Navigate "about:blank"
    sTitle="Database selection" & String( 80, "." ) 'Note: the String( 80,".") is to push "Internet Explorer" string off the window
    objIE.Document.title = sTitle
    objIE.MenuBar        = False
    objIE.ToolBar        = False
    objIE.AddressBar     = false
    objIE.Resizable      = False
    objIE.StatusBar      = False
    objIE.Width          = 250
    objIE.Height         = 160
    ' Center the dialog window on the screen
    With objIE.Document.parentWindow.screen
        objIE.Left = (.availWidth  - objIE.Width ) \ 2
        objIE.Top  = (.availHeight - objIE.Height) \ 2
    End With
    ' Wait till IE is ready
    Do While objIE.Busy
        WScript.Sleep 200
    Loop
    ' Insert the HTML code to prompt for user input
    objIE.Document.body.innerHTML = "<div align=""center"">" & vbcrlf _
                    & "<p><input type=""hidden"" id=""OK"" name=""OK"" value=""0"">" _
                    & "<input type=""submit"" value=""SFS (" & sfs_db & ")"" onClick=""VBScript:OK.value=1""></p>" _
                    & "<input type=""submit"" value=""Markets (" & mark_db & ")"" onClick=""VBScript:OK.value=2""></p>" _
                    & "<p><input type=""hidden"" id=""Cancel"" name=""Cancel"" value=""0"">" _
                    & "<input type=""submit"" id=""CancelButton"" value=""       Cancel       "" onClick=""VBScript:Cancel.value=-1""></p></div>"

    ' Hide the scrollbars
    objIE.Document.body.style.overflow = "auto"
    ' Make the window visible
    objIE.Visible = True
    ' Set focus on Cancel button
    objIE.Document.all.CancelButton.focus
    'CAVEAT: If user click red X to close IE window instead of click cancel, an error will occur.
    '        Error trapping Is Not doable For some reason
    On Error Resume Next
    Do While objIE.Document.all.OK.value = 0 and objIE.Document.all.Cancel.value = 0
        WScript.Sleep 200
        iErrorNum=Err.Number
        If iErrorNum <> 0 Then    'user clicked red X (or alt-F4) to close IE window
            objIE.Quit
            Set objIE = Nothing
            dbselect="bad"
        End if
    Loop
    On Error Goto 0
    objIE.Visible = False
    ' Read the user input from the dialog window
    if dbselect <> "bad" Then
        dbselect = objIE.Document.all.OK.value
    else
        dbselect = 0
    end if
    ' Close and release the object
    objIE.Quit
    Set objIE = Nothing
End Function
