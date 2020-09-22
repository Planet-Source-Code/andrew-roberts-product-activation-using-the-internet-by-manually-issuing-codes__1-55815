Attribute VB_Name = "Module1"
'**************************************
' Name: Embed Files In Executables
' Description:This module will embed any
'     number of files into your executable whi
'     ch can later be extracted by it. Useful


'     for packing multiple files into one inst
    '     aller.
' By: Sean Ferguson
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=34370&lngWId=1'for details.'**************************************



Public Function ExtractEmbeddedFile(strFileName As String, strDestinationFile As String, Optional strSourceFile As String) As Boolean
    On Error GoTo handleError
    Dim fileListStart As Long
    Dim lngFileSize As Long
    Dim lngFilePos As Long
    Dim lngCurPos As Long
    Dim strCurFile As String
    Dim iFreeFile As Integer
    Dim oFreeFile As Integer
    Dim fileData As String
    iFreeFile = FreeFile()
    If Not Len(Trim(strSourceFile)) > 0 Then strSourceFile = App.Path & "\" & App.EXEName & ".exe"
    Open strSourceFile For Binary As iFreeFile
    Get iFreeFile, LOF(iFreeFile) - 3, fileListStart


    If fileListStart = 0 Then
        Close iFreeFile
        Exit Function
    End If


    Do
        Get iFreeFile, fileListStart, lngFilePos
        fileListStart = fileListStart + 4
        Get iFreeFile, fileListStart, lngFileSize
        fileListStart = fileListStart + 4
        strCurFile = String$(255, Chr$(0))
        Get iFreeFile, fileListStart, strCurFile
        If Mid(strCurFile, 1, 1) = Chr$(0) Then strCurFile = Mid(strCurFile, 2)
        strCurFile = Mid(strCurFile, 1, InStr(1, strCurFile, Chr$(0)) - 1)
        fileListStart = fileListStart + Len(strCurFile) + 5


        If lngFilePos = 0 Or lngFileSize = 0 Or Trim(strCurFile) = "" Then
            Close iFreeFile
            Exit Function
        ElseIf strCurFile = strFileName Then
            oFreeFile = FreeFile()
            Open strDestinationFile For Binary As oFreeFile


            If lngFileSize > 1000000 Then
                lngCurPos = -1000000


                Do
                    lngCurPos = lngCurPos + 1000000


                    If lngCurPos + 1000000 > lngFileSize Then
                        fileData = String(lngFileSize - lngCurPos, Chr$(0))
                    Else
                        fileData = String(1000000, Chr$(0))
                    End If
                    Get iFreeFile, lngCurPos + lngFilePos, fileData
                    Put oFreeFile, lngCurPos + 1, fileData
                Loop Until lngCurPos + 999999 >= lngFileSize
            Else
                fileData = String(lngFileSize, Chr$(0))
                Get iFreeFile, lngFilePos, fileData
                Put oFreeFile, 1, fileData
            End If
            Close oFreeFile
            Close iFreeFile
            ExtractEmbeddedFile = True
            Exit Function
        End If
    Loop Until fileListStart >= (LOF(iFreeFile) - 7)
    Close iFreeFile
    Exit Function
handleError:
    Close
    'MsgBox Err.Description
    ExtractEmbeddedFile = False
    Exit Function
End Function


Public Function AddEmbeddedFile(strSourceFile As String, strAddName As String, strDestinationFile As String) As Boolean
    On Error GoTo handleError
    Dim fileListStart As Long
    Dim fileList As String
    Dim lngFileSize As Long
    Dim lngFilePos As Long
    Dim lngCurPos As Long
    Dim strCurFile As String
    Dim iFreeFile As Integer
    Dim oFreeFile As Integer
    Dim fileData As String
    'MsgBox strSourceFile
    If Not FileLen(strSourceFile) > 0 Then Exit Function
    oFreeFile = FreeFile()
    Open strDestinationFile For Binary As oFreeFile
    Get oFreeFile, LOF(oFreeFile) - 3, fileListStart


    If fileListStart = 0 Then
        fileListStart = LOF(oFreeFile) + 1
        fileList = ""
    Else
        fileList = String(LOF(oFreeFile) - fileListStart - 3, Chr$(0))
        Get oFreeFile, fileListStart, fileList
    End If
    lngFilePos = fileListStart
    lngFileSize = FileLen(strSourceFile)
    iFreeFile = FreeFile()
    Open strSourceFile For Binary As iFreeFile


    If LOF(iFreeFile) > 1000000 Then
        lngCurPos = -1000000


        Do
            lngCurPos = lngCurPos + 1000000


            If lngCurPos + 999999 > LOF(iFreeFile) Then
                fileData = String(LOF(iFreeFile) - lngCurPos, Chr$(0))
            Else
                fileData = String(1000000, Chr$(0))
            End If
            Get iFreeFile, lngCurPos, fileData
            Put oFreeFile, fileListStart, fileData
            fileListStart = fileListStart + Len(fileData) + 1
        Loop Until lngCurPos + 999999 > LOF(iFreeFile)
    Else
        fileData = String(LOF(iFreeFile), Chr$(0))
        Get iFreeFile, 1, fileData
        Put oFreeFile, fileListStart, fileData
        fileListStart = fileListStart + Len(fileData) + 1
    End If
    Close iFreeFile
    strAddName = strAddName & Chr$(0)
    Put oFreeFile, fileListStart, fileList
    Put oFreeFile, fileListStart + Len(fileList), lngFilePos
    Put oFreeFile, fileListStart + Len(fileList) + 4, lngFileSize
    Put oFreeFile, fileListStart + Len(fileList) + 8, strAddName
    Put oFreeFile, fileListStart + Len(fileList) + 12 + Len(strAddName), fileListStart
    Close oFreeFile
    AddEmbeddedFile = True
    Exit Function
handleError:
    Close
    'MsgBox Err.Description
    AddEmbeddedFile = False
    Exit Function
End Function

