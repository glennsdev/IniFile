Module Common

    Public Const INIFILE_ERR_FILENAME_REQUIRED = -1000
    Public Const INIFILE_ERR_KEY_REQUIRED = -1001

    Function AppPath() As String
        '-- Return application folder ended with "\"
        '   C:\
        '   C:\Project\

        Return AppDomain.CurrentDomain.BaseDirectory
    End Function
End Module
