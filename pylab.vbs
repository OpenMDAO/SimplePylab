' Set your settings
Set wshShell = CreateObject( "WScript.Shell" )
Set objFSO = Createobject("Scripting.FileSystemObject")

' Environment Variables
HOME = wshShell.ExpandEnvironmentStrings( "%USERPROFILE%" )
PATH = wshShell.ExpandEnvironmentStrings( "%PATH%" )

Function isInPATH( fileName )
    dim fso
    set fso = CreateObject( "Scripting.FileSystemObject" )
    paths = Split( PATH, ";")
    For i=lbound(paths) to ubound(paths):
        If isInDirectory( paths(i), fileName ) Then
            found = True
        End If
    Next
    
    isInPath = found
End Function

Function isInDirectory( dirPath, fileName )
    Dim fso
    set fso = CreateObject("Scripting.FileSystemObject")
    isInDirectory = fso.FileExists( dirPath & "\" & fileName )
End Function

Function getAbsPath(path)
    Dim fso
    set fso = CreateObject("Scripting.FileSystemObject")
    getAbsPath = fso.GetAbsolutePathName(path)
End Function

Function getCwd()
    Dim wsh
    Set wsh = CreateObject( "WScript.Shell" )
    getCwd = wsh.CurrentDirectory
End Function

Sub downloadfile(Byref XMLHTTP, filePath)
    Set objADOStream = CreateObject("ADODB.Stream")

    objADOStream.Open
    objADOStream.Type = 1 'adTypeBinary

    objADOStream.Write XMLHTTP.ResponseBody
    objADOStream.Position = 0 'Set the stream position to the start

    Wscript.Echo filePath
    objADOStream.SaveToFile filePath
    objADOStream.Close

    Set objADOStream = Nothing
End Sub
    
Sub getHttp(ByVal URL, ByVal filePath)
    Dim objXMLHTTP
    Set objXMLHTTP = CreateObject("MSXML2.XMLHTTP.3.0")

    objXMLHTTP.Open "GET", URL, False
    objXMLHTTP.send()

    If objXMLHTTP.Status = 200 Then
        downloadFile objXMLHTTP, filePath
    End if

    Set objXMLHTTP = Nothing
End Sub

' Capture command line 
Set args = WScript.Arguments
python_version = "2.7.3"

If args.count > 0 Then
    For i=Lbound(args) to Ubound(args.count)
        ' If args(i) = "/v" then
        '    ' Specific version of python to download
        '    python_version = args(i+1)
        '    i = i + 1
        ' Elseif args(i) = "/p" then
        '     ' Overrides /v option
        '     ' Path to directory containing python.exe
        '     python_dir = getAbsPath(args(i+1))
        '     i = i + 1
        ' Elseif args(i) = "/e" then
        '     ' Path to virtualenv script
        '     virtualenv_dir = getAbsPath(args(i+1))
        '     i = i + 1
        ' Elseif args(i) = "/d" then
        '     num_package = getAbsPath(args(i+1))
        '     i = i + 1
        If args(i) = "/h" then
            ' Print help message
            WScript.Echo "pylab.vbs " & vbCRLF & vbcrlf &_ 
            "A script that downloads and installs " & vbcrlf & _
            "pylab, a collection of python packages including:" & vbcrlf & vbcrlf &_
            "    Numpy" & vbcrlf &_
            "    Scipy" & vbcrlf &_
            "    Matplotlib" & vbcrlf &_
            "    Ipython" & vbcrlf & vbcrlf &_
            "Python and the pylab installer" & vbcrlf &_
            "can be downloaded/installed if they " & vbcrlf &_
            "do not exist on the local machine." & vbcrlf & vbcrlf &_
            "pylab.vbs [ /h ]" & vbcrlf & vbcrlf &_
            "/h:                Prints this help message" & vbcrlf &_
            "                   and exits."
            WScript.Quit()
        Else
            WScript.Echo "Command incorect. run pylab.vbs /h for correct usage."
            WScript.Quit()
        End If
    Next
End If

pyversion = "2.7.3"
python_msi = "python-" & pyversion & ".msi"
pythonURL = "http://www.python.org/ftp/python/" & pyversion & "/" & python_msi
virutalenvURL = "http://gvis.grc.nasa.gov/sciapps/pylab_virtualenv.py"

WScript.Echo "Looking for python interpreter."
If isInPATH("python.exe") or isInDirectory( getCwd(), "python.exe" ) Then
    WScript.Echo "Found python interpreter."
Else
        PYTHON_HOME = getCwd() & "\Python27"
        WScript.Echo "Cannot find python interpreter (python.exe)"
        WScript.Echo "Downloading and installing python interpreter to " & getAbsPath(PYTHON_HOME)
        getHttp pythonURL, python_msi
        wshShell.Run "msiexec /qn /i " & python_msi & " TARGETDIR=" & PYTHON_HOME & "ADDLOCAL=Tools,TclTk", 0, True
        If not objFSO.FolderExists( PYTHON_HOME ) Then
            objFSO.MoveFolder HOME & "\Python27", PYTHON_HOME
        End If
        WScript.Echo "Installation of Python is complete."
End If

WScript.Echo "Looking for pylab installer."
If isInPATH("pylab_virtualenv.py") or isInDirectory(getCwd(), "pylab_virtualenv.py") Then
    WScript.Echo "Found pylab installer."
Else
    WScript.Echo "Cannot find pylab installer (pylab_virtualenv.py)"
    WScript.Echo "Downloading pylab installer to " & getAbsPath(getCwd())
    PYLAB_HOME = getCwd()
    virtualenv_script =  pylab_home & "\" & "pylab_virtualenv.py"
    getHttp virtualenvURL, virtualenv_script
    WScript.Echo "Finished downloading Pylab installer."
End If

WScript.Echo "Installing pylab to " & getCwd() & "\pylab"
wshShell.Run "cmd /c " & PYTHON_HOME & " python.exe pylab_virtualenv.py pylab"