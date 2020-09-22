VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Add Code"
      Height          =   735
      Left            =   1920
      TabIndex        =   1
      Top             =   1080
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   960
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Serial Code:"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    dialog.ShowOpen
    Text1 = dialog.FileName
End Sub

Private Sub Command2_Click()
    Open App.Path & "\temp.dat" For Output As #1
    Print #1, Text1
    Close #1
    
    'check if a code has already been added to this file so that
    'the first code does not override the new code
    If Module1.ExtractEmbeddedFile("serialCode", App.Path & "\temp.dat", App.Path & "\about.exe") = True Then
        MsgBox "A serial code has already been added to the about.exe file! You should make copies of this file so that you can produce new codes on the fly!", vbCritical, "ERROR"
        Exit Sub
    End If
    
    If Module1.AddEmbeddedFile(App.Path & "\temp.dat", "serialCode", App.Path & "\about.exe") = True Then
        MsgBox "Serial Code added successfully!" & vbCrLf & vbCrLf & "NOTE: If you intend to use this activation method in your own programs you may want to encrypt the serial number in the file first so that nobody can change it!", vbInformation, "Done"
    Else
        MsgBox "Error adding serial code!", vbCritical, "Error"
    End If
    
    Kill App.Path & "\temp.dat"
    
End Sub

'Private Sub Command3_Click()
    'MsgBox Module1.ExtractEmbeddedFile(Text3, App.Path & "\" & Text3, App.Path & "\about.exe")
'End Sub

