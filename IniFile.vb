﻿Imports System.IO
Imports System.Text

Public Class IniFile
    Dim m_aItems() As String
    Dim m_nItemsCount As Int32

    Dim m_sPathFileName As String = ""

    Public Function LoadFile(ByVal FileName As String) As Boolean
        Dim fs As FileStream
        Dim sr As StreamReader

        Dim sLine As String

        FileName = Trim(FileName)

        Try
            If FileName = "" Then
                Err.Raise(INIFILE_ERR_FILENAME_REQUIRED, , "Missing file name parameter")
            End If

            If InStr(FileName, "\") = 0 Then        '-- Only filename without path
                '-- Add current EXE path
                m_sPathFileName = AppPath() & FileName
            Else
                m_sPathFileName = FileName
            End If

            '-- Initialize array items
            m_nItemsCount = 0

            If System.IO.File.Exists(m_sPathFileName) = False Then      '-- Not found
                Return False
            End If

            fs = New FileStream(m_sPathFileName, FileMode.OpenOrCreate, FileAccess.Read, FileShare.ReadWrite)
            sr = New StreamReader(fs, Encoding.UTF8)

            While True
                sLine = Trim(sr.ReadLine())

                m_nItemsCount = m_nItemsCount + 1

                ReDim Preserve m_aItems(m_nItemsCount)

                m_aItems(m_nItemsCount - 1) = sLine

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

            For i = 1 To m_nItemsCount
                sWrite = sWrite & m_aItems(i - 1) & vbCrLf
            Next


            fs = New FileStream(m_sPathFileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
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

            sRetVal = m_aItems(nIndex)

            nPos = InStr(sRetVal, "=")

            If (nPos > 0) Then
                sRetVal = Trim(sRetVal.Substring(nPos))
            End If

            Return sRetVal
        End Get

        Set(value As String)
            Dim nIndex As Integer

            Key = Trim(Key)
            value = Trim(value)

            If Key = "" Then
                Err.Raise(INIFILE_ERR_KEY_REQUIRED, , "Missing key parameter")
            End If

            nIndex = _GetIndex(Key)

            If nIndex = -1 Then     '-- Not found
                '-- Add new. Put key value at the bottom
                m_nItemsCount = m_nItemsCount + 1

                ReDim Preserve m_aItems(m_nItemsCount)

                nIndex = m_nItemsCount - 1
            End If

            m_aItems(nIndex) = Key & "=" & value
        End Set
    End Property

    Private Function _GetIndex(ByVal Key As String) As Int32
        Dim i As Integer
        Dim sItem As String
        Dim sCheckKey As String
        Dim nPos As Int32


        Key = LCase(Key)

        For i = 1 To m_nItemsCount
            sItem = m_aItems(i - 1)

            If Left(sItem, 1) <> ";" Then       '-- Ignore remark char ";"
                nPos = InStr(sItem, "=")        '-- Find "=" char

                If nPos > 0 Then
                    '-- Get Key
                    sCheckKey = Trim(LCase(Left(m_aItems(i - 1), nPos - 1)))

                    If sCheckKey = Key Then
                        '-- Return actual array index (zero based array)
                        Return (i - 1)
                    End If
                End If

            End If
        Next

        Return -1   '-- Not Found
    End Function
End Class