VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmRegWiz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KTK License Registration"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7065
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDone 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   6855
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Label Label14 
         Caption         =   $"Form1.frx":000C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   6615
      End
      Begin VB.Label Label13 
         Caption         =   $"Form1.frx":00B6
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   6735
      End
      Begin VB.Label Label12 
         Caption         =   "Thankyou"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         TabIndex        =   12
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.PictureBox picWait 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   120
      ScaleHeight     =   3615
      ScaleWidth      =   6855
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   6855
      Begin VB.Label Label11 
         Caption         =   "License Registration is now linking this product to your computer. This process may take a few minutes to complete"
         Height          =   855
         Left            =   1080
         TabIndex        =   10
         Top             =   1560
         Width           =   5775
      End
      Begin VB.Label Label10 
         Caption         =   "Please Wait..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   1200
         Width           =   2055
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1080
      TabIndex        =   6
      Top             =   3720
      Width           =   5775
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4680
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      RequestTimeout  =   30
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   5160
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   6360
      Picture         =   "Form1.frx":0172
      ScaleHeight     =   525
      ScaleWidth      =   525
      TabIndex        =   2
      Top             =   360
      Width           =   525
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "NOTE: You must have an active internet connection to complete product registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Width           =   6855
   End
   Begin VB.Label Label5 
      Caption         =   "Product registration is required. We only need your name to complete registration."
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   2880
      Width           =   6855
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":1078
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   6855
   End
   Begin VB.Label Label3 
      Caption         =   "Why complete product registration?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   6855
   End
   Begin VB.Label lblHardwareFingerPrint 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label10"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   960
      Width           =   6975
   End
   Begin VB.Label Label8 
      Caption         =   "Name:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Register Product"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "KTK License Registration Wizard"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.Image Image2 
      Height          =   30
      Left            =   -3840
      Picture         =   "Form1.frx":1147
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   10935
   End
   Begin VB.Image Image1 
      Height          =   30
      Left            =   -2040
      Picture         =   "Form1.frx":11C1
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   9135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   -1920
      Top             =   -1080
      Width           =   9495
   End
End
Attribute VB_Name = "frmRegWiz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const id = "we"

Const EncryptName = "Putencyptstringhere "

Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'String to hold Registry Computer Name
Public SysInfoPath As String

Public productCode As String

Const productID = "File_Explorer_3"

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "System\CurrentControlSet\Control\ComputerName\ComputerName"
Const gREGKEYSYSINFO = "System\CurrentControlSet\Control\ComputerName\ComputerName"
Const gREGVALSYSINFO = "ComputerName"
Const RegKey = "Reg"
Public Register As String

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'Put your project name here
'This is an entry in the registry that is created
Const RegPath = "SOFTWARE\KTK File Explorer 3.00"

Public header As String
Public licenseNo As String
Public licenseHolder As String
Public hardwareFingerPrint As String
Public evalDays As String
Public dateRegistered As String
Public dateLastUsed As String
Public footer As String


Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdNext_Click()
    'MsgBox fingerPrint
    
    If txtName = "" Then
        MsgBox "You must enter your name!", vbCritical, "Product Registration"
        picWait.Visible = False
        cmdNext.Visible = True: cmdCancel.Visible = True
        Exit Sub
    End If
    
    'the code is correct.
    
    'check if we are connected to the Internet first
    'If InternetGetConnectedState(Flags, 0) = False Then
    '    MsgBox "The following error was encountered while performing online lookup:" & Chr(13) & Chr(13) & "Active Internet connection not found!", vbCritical, "KTK License Registration Wizard"
    '    picWait.Visible = False
    '    cmdNext.Visible = True: cmdCancel.Visible = True
    '    Exit Sub
    'End If
    
    'MsgBox InternetGetConnectedState(flags, 0)
    
    'check to see if this serial code has already been issued
    
    f = FreeFile
    
    Open App.Path & "\test.dat" For Output As #f
    'Debug.Print "members.lycos.co.uk/ktkpiracy/" & productid & "/" & txtSeed & "_" & txtVal1 & "_" & txtVal2 & "_" & txtVal3 & ".txt"
    d = Inet1.OpenURL("members.lycos.co.uk/ktkpiracy/" & productID & "/" & productCode & ".txt")
    Print #f, d
    Close #f
    
    'MsgBox d
    
    extractData App.Path & "\test.dat"
    
    'check to see if this license has been registered
    
    If header = "KTK" Then
        If licenseHolder = "" Then
            'not registered
            s = 0
        ElseIf fingerPrint = hardwareFingerPrint Then
            s = 0 'this user has already this product
        ElseIf Not (fingerPrint = hardwareFingerPrint) Then
            'we must do a day lookup. The computer can be changed once
            'every 30 days!
            If DateDiff("d", CDate(dateLastUsed), Date, vbUseSystemDayOfWeek, vbUseSystem) > 30 Then
                MsgBox "This product has already been registered on another computer!" & Chr(13) & "License Registration restricts the usage of licenses already registered on another computer. However, this license has been inactive for over 30 days so you can re-register it.", vbInformation, "KTK License Registration Wizard"
                s = 0
            Else
                s = 1
                MsgBox "This product has already been registered on another computer!" & Chr(13) & "License Registration restricts the usage of licenses already registered on another computer. The license you are trying to use must be inactive for atleast 30 days before it can be re-registered." & Chr(13) & Chr(13) & "You will be able to re-register this product on " & Format(CDate(dateLastUsed) + 31, "dddd dd mmmm yyyy", vbUseSystemDayOfWeek, vbUseSystem), vbInformation, "KTK License Registered Wizard"
            End If
        Else
            s = 1
            'has been registered
        End If
    Else
        'not registered
        s = 0
    End If
    
    If s = 1 Then
        MsgBox "We could not activate this product because it has already been activated!"
        picWait.Visible = False
        cmdNext.Visible = True: cmdCancel.Visible = True
        Exit Sub
    End If
    
    Inet1.OpenURL "members.lycos.co.uk/ktkpiracy/writeFile.php?productID=" & productCode & "&licenseNo=" & ProductName & "&licenseHolder=" & txtName.Text & "&licenseHardwareKey=" & fingerPrint & "&daysEval=0&dateRegistered=" & Date & "&lastUsed=" & Date
    Me.picDone.Visible = True
    cmdCancel.Visible = True
    cmdCancel.Caption = "Close"
End Sub

Sub extractData(file)
    On Error Resume Next
    f = FreeFile
    Open file For Input As #f
    Input #f, header
    Input #f, licenseNo
    Input #f, licenseHolder
    Input #f, hardwareFingerPrint
    Input #f, evalDays
    Input #f, dateRegistered
    Input #f, dateLastUsed
    Input #f, footer
    Close #f
End Sub

Function fingerPrint()
          Dim TempStr As String
          Dim RegStr As String
          Dim I As Integer
          Dim SerialNumber As Long
          
10       On Error GoTo fingerPrint_Error

          'Get The Computer Name in the registry
          'StartSysInfo
20        SerialNumber = GetSerialNumber("C:\")
30        SysInfoPath = Str(SerialNumber)
          
          'For encrypting purposes make the length
          'of it no more than 20 character
40        If Len(SysInfoPath) > 20 Then
50            SysInfoPath = Left$(SysInfoPath, 20)
60        End If
          'invert the computer name
70        InvertIt
80        EncryptIt
90        EncipherIt
100       GetSubKey
          
110       fingerPrint = SysInfoPath

120      On Error GoTo 0
130      Exit Function

fingerPrint_Error:
140       'writeLog Err.Number, Err.Description, "frmRegWiz::fingerPrint::" & Erl
150       'frmError.txtError = "Error Number: " & Err.Number & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Source: " & Err.Source & Chr(13) & Chr(10) & "Sub: frmRegWiz::fingerPrint::" & Erl
160       'frmError.Show vbModal
          'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fingerPrint of Form frmRegWiz" & vbcrlf & "Line Number: " & erl
          
End Function

'GetSerialNumber Procedure - Put this in the module or form where it is called.
Function GetSerialNumber(strDrive As String) As Long
          Dim SerialNum As Long
          Dim Res As Long
          Dim Temp1 As String
          Dim Temp2 As String
10       On Error GoTo GetSerialNumber_Error

          'initialise the strings
20        Temp1 = String$(255, Chr$(0))
30        Temp2 = String$(255, Chr$(0))
          'call the API function
40        Res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
50        GetSerialNumber = SerialNum

60       On Error GoTo 0
70       Exit Function

GetSerialNumber_Error:
80        'writeLog Err.Number, Err.Description, "frmRegWiz::GetSerialNumber::" & Erl
90        'frmError.txtError = "Error Number: " & Err.Number & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Source: " & Err.Source & Chr(13) & Chr(10) & "Sub: frmRegWiz::GetSerialNumber::" & Erl
100       'frmError.Show vbModal
          'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetSerialNumber of Form frmRegWiz" & vbcrlf & "Line Number: " & erl
          
End Function

Sub EncipherIt()
          Dim Temp As Integer
          Dim Hold As String
          Dim I As Integer
          Dim J As Integer
          Dim TempStr As String
          Dim Temp1 As String
          
10       On Error GoTo EncipherIt_Error

20        TempStr = ""
30        For I = 1 To Len(SysInfoPath)
40            Temp = Asc(Mid$(SysInfoPath, I, 1))
50            Temp1 = Hex(Temp)
60            If Len(Temp1) = 1 Then
70                Temp1 = "0" & Temp1
80            End If
90            For J = 1 To 2
100               Hold = Mid$(Temp1, J, 1)
110               Select Case Hold
                      Case "0"
120                       TempStr = TempStr + "7"
130                   Case "1"
140                       TempStr = TempStr + "B"
150                   Case "2"
160                       TempStr = TempStr + "F"
170                   Case "3"
180                       TempStr = TempStr + "D"
190                   Case "4"
200                       TempStr = TempStr + "1"
210                   Case "5"
220                       TempStr = TempStr + "9"
230                   Case "6"
240                       TempStr = TempStr + "3"
250                   Case "7"
260                       TempStr = TempStr + "A"
270                   Case "8"
280                       TempStr = TempStr + "6"
290                   Case "9"
300                       TempStr = TempStr + "5"
310                   Case "A"
320                       TempStr = TempStr + "E"
330                   Case "B"
340                       TempStr = TempStr + "8"
350                   Case "C"
360                       TempStr = TempStr + "0"
370                   Case "D"
380                       TempStr = TempStr + "C"
390                   Case "E"
400                       TempStr = TempStr + "2"
410                   Case "F"
420                       TempStr = TempStr + "4"
430               End Select
440           Next J
450       Next I
460       SysInfoPath = TempStr

470      On Error GoTo 0
480      Exit Sub

EncipherIt_Error:
490       'writeLog Err.Number, Err.Description, "frmRegWiz::EncipherIt::" & Erl
500       'frmError.txtError = "Error Number: " & Err.Number & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Source: " & Err.Source & Chr(13) & Chr(10) & "Sub: frmRegWiz::EncipherIt::" & Erl
510       'frmError.Show vbModal
          'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EncipherIt of Form frmRegWiz" & vbcrlf & "Line Number: " & erl
End Sub

Sub EncryptIt()
          Dim Temp As Integer
          Dim Temp1 As Integer
          Dim Hold As Integer
          Dim I As Integer
          Dim J As Integer
          Dim TempStr As String

10       On Error GoTo EncryptIt_Error

20        TempStr = ""
30        For I = 1 To Len(EncryptName)
40            Hold = 0
50            Temp = Asc(Mid$(EncryptName, I, 1))
60            For J = 1 To Len(SysInfoPath)
70                Temp1 = Asc(Mid$(SysInfoPath, J, 1))
80                Hold = Temp Xor Temp1
90             Next J
100           TempStr = TempStr + Chr(Hold)
110       Next I
          
120       SysInfoPath = TempStr

130      On Error GoTo 0
140      Exit Sub

EncryptIt_Error:
150       'writeLog Err.Number, Err.Description, "frmRegWiz::EncryptIt::" & Erl
160       'frmError.txtError = "Error Number: " & Err.Number & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Source: " & Err.Source & Chr(13) & Chr(10) & "Sub: frmRegWiz::EncryptIt::" & Erl
170       'frmError.Show vbModal
          'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EncryptIt of Form frmRegWiz" & vbcrlf & "Line Number: " & erl
End Sub

Sub InvertIt()
          Dim Temp As Integer
          Dim Hold As Integer
          Dim I As Integer
          Dim TempStr As String
              
10       On Error GoTo InvertIt_Error

20        TempStr = ""
30        For I = 1 To Len(SysInfoPath)
40            Temp = Asc(Mid$(SysInfoPath, I, 1))
50            Hold = 0
Top:
60        Select Case Temp
              Case Is > 127
70                Hold = Hold + 1
80                Temp = Temp - 128
90                GoTo Top
100           Case Is > 63
110               Hold = Hold + 2
120               Temp = Temp - 64
130               GoTo Top
140           Case Is > 31
150               Hold = Hold + 4
160               Temp = Temp - 32
170               GoTo Top
180           Case Is > 15
190               Hold = Hold + 8
200               Temp = Temp - 16
210               GoTo Top
220           Case Is > 7
230               Hold = Hold + 16
240               Temp = Temp - 8
250               GoTo Top
260           Case Is > 3
270               Hold = Hold + 32
280               Temp = Temp - 4
290               GoTo Top
300           Case Is > 1
310               Hold = Hold + 64
320               Temp = Temp - 2
330               GoTo Top
340           Case Is = 1
350               Hold = Hold + 128
                  
360       End Select
370           Temp = 255 Xor Hold
380           TempStr = TempStr + Chr(Temp)
390       Next I
          
400       SysInfoPath = TempStr

410      On Error GoTo 0
420      Exit Sub

InvertIt_Error:
430       'writeLog Err.Number, Err.Description, "frmRegWiz::InvertIt::" & Erl
440       'frmError.txtError = "Error Number: " & Err.Number & Chr(13) & Chr(10) & Err.Description & Chr(13) & Chr(10) & "Source: " & Err.Source & Chr(13) & Chr(10) & "Sub: frmRegWiz::InvertIt::" & Erl
450       'frmError.Show vbModal
          'MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InvertIt of Form frmRegWiz" & vbcrlf & "Line Number: " & erl
End Sub

Public Sub GetSubKey()

    If Not GetKeyValue(HKEY_LOCAL_MACHINE, RegPath, RegKey, Register) Then
        'Rem Not in registry
        
    End If
    
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim I As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For I = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, I, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Load()
    lblHardwareFingerPrint.Caption = "Computer Reference: " & fingerPrint
    If Module1.ExtractEmbeddedFile("serialCode", App.Path & "\temp.dat", App.Path & "\about.exe") = True Then
        Dim data As String
        Open App.Path & "\temp.dat" For Input As #1
        Input #1, data
        Close #1
        productCode = data
        Kill App.Path & "\temp.dat"
    Else
        MsgBox "Error: ProductName could not be located!", vbCritical, "KTK License Registration"
        End
    End If
    
    'msgboc productCode
    
End Sub
