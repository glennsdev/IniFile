Imports System.IO
Imports System.Text

Public Class IniFile
    Dim m_aText() As String
    Dim m_nArrayCount As Int32

    Dim m_sPathFileName As String = ""

    Public Function LoadFile(ByVal FileName As String) As Boolean
        Dim fs As FileStream
        Dim sr As StreamReader

        Dim sLine As String

        FileName = Trim(FileName)

        Try
            If FileName = "" Then
                Err.Raise(INIFILE_ERR_FILENAME_REQUIRED,, "Missing file name parameter")
            End If

            If InStr(FileName, "\") = 0 Then        '-- Only filename without path
                '-- Add current EXE path
                m_sPathFileName = AppPath() & FileName
            Else
                m_sPathFileName = FileName
            End If

            fs = New FileStream(m_sPathFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.ReadWrite)
            sr = New StreamReader(fs, Encoding.UTF8)

            m_nArrayCount = 0

            While True
                sLine = sr.ReadLine()

                m_nArrayCount = m_nArrayCount + 1

                ReDim Preserve m_aText(m_nArrayCount)

                m_aText(m_nArrayCount - 1) = sLine

                If sr.EndOfStream Then
                    Exit While
                End If
            End While

            sr.Close()
            fs.Close()

            sr = Nothing
            fs = Nothing

            Return True

        Catch ex As Exception
            sr = Nothing
            fs = Nothing

            Return False
        End Try
    End Function

    Public Function SaveFile(Optional FileName As String = "") As Boolean
        Dim fs As FileStream
        Dim sw As StreamWriter

        Dim sWrite As String
        Dim i As Int32

        FileName = Trim(FileName)

        Try
            If FileName = "" And m_sPathFileName = "" Then
                Err.Raise(INIFILE_ERR_FILENAME_REQUIRED, , "Missing file name parameter")
            End If

            If FileName <> "" Then      '-- Change INI file name
                If InStr(FileName, "\") = 0 Then        '-- Only filename without path
                    '-- Add current EXE path
                    m_sPathFileName = AppPath() & FileName
                Else
                    m_sPathFileName = FileName
                End If
            End If

            sWrite = ""

            For i = 1 To m_nArrayCount
                sWrite = sWrite & m_aText(i - 1) & vbCrLf
            Next


            fs = New FileStream(m_sPathFileName, FileMode.Create, FileAccess.Write)
            sw = New StreamWriter(fs, Encoding.UTF8)

            sw.Write(sWrite)

            sw.Close()
            fs.Close()

            sw = Nothing
            fs = Nothing

            Return True

        Catch ex As Exception
            sw = Nothing
            fs = Nothing

            Return False
        End Try
    End Function

    Default Public Property Items(ByVal Key As String) As String
        Get
            Dim nIndex As Int32
            Dim sRetVal As String = ""
            Dim nPos As Int32

            Key = Trim(Key)

            If Key = "" Then
                Err.Raise(INIFILE_ERR_KEY_REQUIRED, , "Missing key parameter")
            End If

            nIndex = _GetIndex(Key)

            If nIndex = -1 Then     '-- Not found
                Return ""
            End If


            '-- Found

            sRetVal = m_aText(nIndex)

            nPos = InStr(sRetVal, "=")

            If (nPos > 0) Then
                sRetVal = sRetVal.Substring(nPos)
            Else

            End If

            Return sRetVal
        End Get

        Set(value As String)
            Dim nIndex As Integer

            Key = Trim(Key)

            If Key = "" Then
                Err.Raise(INIFILE_ERR_KEY_REQUIRED, , "Missing key parameter")
            End If

            nIndex = _GetIndex(Key)

            If nIndex = -1 Then     '-- Not found
                '-- Add new. Put key value at the bottom
                m_nArrayCount = m_nArrayCount + 1

                ReDim Preserve m_aText(m_nArrayCount)

                nIndex = m_nArrayCount - 1
            End If

            m_aText(nIndex) = Key & "=" & value
        End Set
    End Property

    Private Function _GetIndex(ByVal Key As String) As Int32
        Dim i As Integer
        Dim sSearchKey As String
        Dim nPos As Int32

        sSearchKey = Key & "="   '-- "key="

        For i = 1 To m_nArrayCount
            nPos = InStr(m_aText(i - 1), sSearchKey, CompareMethod.Text)

            If nPos = 1 Then
                Return i - 1
            End If
        Next

        Return -1   '-- Not Found
    End Function
End Class