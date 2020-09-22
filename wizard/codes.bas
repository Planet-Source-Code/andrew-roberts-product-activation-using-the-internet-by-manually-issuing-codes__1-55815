Attribute VB_Name = "modPiracy"
'Public Declare Function InternetGetConnectedState Lib "wininet" (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean

Global Const productid = "File_Explorer_3"

Public seed1 As String
Public val11 As String
Public val21 As Variant
Public val31 As String

Public Function createCode(val1, val3, seed)
10       On Error GoTo createCode_Error

20        On Error Resume Next
30        createcode2 = seed * val1 / 3 * 200 * val3 / 5 * 137
          
40        createCode = seed & "-" & val1 & "-" & Hex(createcode2) & "-" & val3
          
50        seed1 = seed
60        val11 = val1
70        val21 = CStr(Hex(createcode2))
80        val31 = val3

90       On Error GoTo 0
100      Exit Function

createCode_Error:
110       writeLog err.Number, err.Description, "modPiracy::createCode::" & Erl
120       frmError.txtError = "Error Number: " & err.Number & Chr(13) & Chr(10) & err.Description & Chr(13) & Chr(10) & "Source: " & err.Source & Chr(13) & Chr(10) & "Sub: modPiracy::createCode::" & Erl
130       frmError.Show vbModal
          'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure createCode of Module modPiracy" & vbcrlf & "Line Number: " & erl
          
End Function

Public Function validateCode(seed, val1, val2, val3) As Boolean
10       On Error GoTo validateCode_Error

          'convert val2 to value
20        On Error Resume Next
30        val2 = Val("&H" & val2)
40        On Error Resume Next
50        Temp1 = seed * val1 / 3 * 200 * val3 / 5 * 137
          
60        If Temp1 = val2 Then validateCode = True Else validateCode = False

70       On Error GoTo 0
80       Exit Function

validateCode_Error:
90        writeLog err.Number, err.Description, "modPiracy::validateCode::" & Erl
100       frmError.txtError = "Error Number: " & err.Number & Chr(13) & Chr(10) & err.Description & Chr(13) & Chr(10) & "Source: " & err.Source & Chr(13) & Chr(10) & "Sub: modPiracy::validateCode::" & Erl
110       frmError.Show vbModal
          'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure validateCode of Module modPiracy" & vbcrlf & "Line Number: " & erl
          
End Function
