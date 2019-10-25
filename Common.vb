Module Common

    Public Const INIFILE_ERR_FILENAME_REQUIRED = -1000
    Public Const INIFILE_ERR_KEY_REQUIRED = -1001

    Function AppPath() As String
        '-- Return application folder ended with "\"
        '   C:\
        '   C:\Project\

        Dim sPath As String = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)

        If Right(sPath, 1) <> "\" Then
            sPath = sPath & "\"
        End If

        Return sPath
    End Function
End Module
