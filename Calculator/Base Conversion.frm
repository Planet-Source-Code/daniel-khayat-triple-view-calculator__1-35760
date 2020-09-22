VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Base 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Base Conversion"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "Base Conversion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4050
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdF 
      Caption         =   "F"
      Height          =   495
      Left            =   3360
      TabIndex        =   22
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdE 
      Caption         =   "E"
      Height          =   495
      Left            =   2640
      TabIndex        =   21
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdD 
      BackColor       =   &H8000000B&
      Caption         =   "D"
      Height          =   495
      Left            =   1920
      TabIndex        =   20
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdC 
      Caption         =   "C"
      Height          =   495
      Left            =   1200
      TabIndex        =   19
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdB 
      Caption         =   "B"
      Height          =   495
      Left            =   3360
      TabIndex        =   18
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton cmdA 
      Caption         =   "A"
      Height          =   495
      Left            =   2640
      TabIndex        =   17
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   1920
      TabIndex        =   16
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   1200
      TabIndex        =   15
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3360
      TabIndex        =   14
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   2640
      TabIndex        =   13
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   1920
      TabIndex        =   12
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   1200
      TabIndex        =   11
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   3360
      TabIndex        =   10
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   2640
      TabIndex        =   9
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   8
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton Digits 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   7
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   600
      Left            =   5280
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   3555
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   915
      Begin VB.OptionButton optoct 
         Caption         =   "Oct"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   675
      End
      Begin VB.OptionButton optbin 
         Caption         =   "Bin"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   2880
         Width           =   585
      End
      Begin VB.OptionButton opthex 
         Caption         =   "Hex"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   645
      End
      Begin VB.OptionButton optdec 
         Caption         =   "Dec"
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   705
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      MaxLength       =   20
      MousePointer    =   3  'I-Beam
      TabIndex        =   0
      Top             =   135
      Width           =   3855
   End
   Begin MSForms.CommandButton cmdclearf2 
      Height          =   495
      Index           =   2
      Left            =   1200
      TabIndex        =   26
      Top             =   720
      Width           =   1215
      ForeColor       =   255
      Caption         =   "Clear"
      Size            =   "2143;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdback 
      Height          =   495
      Index           =   1
      Left            =   1200
      TabIndex        =   25
      Top             =   3720
      Width           =   1335
      ForeColor       =   255
      Caption         =   "Back"
      Size            =   "2355;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton CommandButton1 
      Height          =   495
      Index           =   0
      Left            =   2640
      TabIndex        =   24
      Top             =   3720
      Width           =   1335
      ForeColor       =   255
      Caption         =   "Exit"
      Size            =   "2355;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton backspace 
      Height          =   495
      Left            =   2640
      TabIndex        =   23
      Top             =   720
      Width           =   1335
      ForeColor       =   255
      Caption         =   "Backspace"
      Size            =   "2355;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "Base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim currentbase As Byte
Dim ClearDisplay As Boolean

Private Sub backspace_Click()
 
 If Len(Text1.Text) < "1" Then
    Text1.Text = ""
 End If
    If Len(Text1.Text) = "1" Then
        Text1.Text = ""
    End If
        If Len(Text1.Text) > 1 Then
          Text1.Text = Left$(Text1.Text, Len(Text1.Text) - 1)
        End If
Text1.SetFocus

End Sub


Private Sub cmdA_Click()
    
    Text1.Text = Text1.Text + "A"

End Sub

Private Sub cmdB_Click()
    
    Text1.Text = Text1.Text + "B"

End Sub

Private Sub cmdback_Click(Index As Integer)
    
    Calfrm.Show
    Base.Hide

End Sub

Private Sub cmdC_Click()
    
    Text1.Text = Text1.Text + "C"

End Sub

Private Sub cmdclearf2_Click(Index As Integer)
    
    Text1.Text = ""
    Text2.Text = ""
    Text1.SetFocus

End Sub

Private Sub cmdD_Click()
    
    Text1.Text = Text1.Text + "D"

End Sub

Private Sub cmdE_Click()
    
    Text1.Text = Text1.Text + "E"

End Sub

Private Sub cmdF_Click()
    
    Text1.Text = Text1.Text + "F"

End Sub

Private Sub CommandButton1_Click(Index As Integer)
    
    End

End Sub

Private Sub Digits_Click(Index As Integer)
  
  If ClearDisplay Then
      Text1.Text = "" And Text2.Text = ""
  End If
  Text1.Text = Text1.Text + Digits(Index).Caption

End Sub

Private Sub Form_Load()
    currentbase = 1
    cmdA.Enabled = False
    cmdB.Enabled = False
    cmdC.Enabled = False
    cmdD.Enabled = False
    cmdE.Enabled = False
    cmdF.Enabled = False
End Sub

Private Sub optdec_Click()
    Text2.Text = ""
    cmdA.Enabled = False
    cmdB.Enabled = False
    cmdC.Enabled = False
    cmdD.Enabled = False
    cmdE.Enabled = False
    cmdF.Enabled = False
    Digits(0).Enabled = True
    Digits(1).Enabled = True
    Digits(2).Enabled = True
    Digits(3).Enabled = True
    Digits(4).Enabled = True
    Digits(5).Enabled = True
    Digits(6).Enabled = True
    Digits(7).Enabled = True
    Digits(8).Enabled = True
    Digits(9).Enabled = True

On Error GoTo errorHandler
If Text1.Text <> "" Then
    Select Case currentbase
        Case 2
            Text1.Text = HexToDec(Text1.Text)
        Case 3
            Text1.Text = OctToDec(Text1.Text)
        Case 4
            Text1.Text = BinToDec(Text1.Text)
    End Select
End If
currentbase = 1
Text1.SetFocus
Exit Sub
errorHandler:
    Text1.Text = "ERROR"
    MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error"  'the vbCrlf constant inserts a line break between the literal string and the error's description
    ClearDisplay = True
    currentbase = 1
    Text1.SetFocus
End Sub

Private Sub opthex_Click()
    Text2.Text = ""
    cmdA.Enabled = True
    cmdB.Enabled = True
    cmdC.Enabled = True
    cmdD.Enabled = True
    cmdE.Enabled = True
    cmdF.Enabled = True
    Digits(0).Enabled = True
    Digits(1).Enabled = True
    Digits(2).Enabled = True
    Digits(3).Enabled = True
    Digits(4).Enabled = True
    Digits(5).Enabled = True
    Digits(6).Enabled = True
    Digits(7).Enabled = True
    Digits(8).Enabled = True
    Digits(9).Enabled = True

On Error GoTo errorHandler

If Text1.Text <> "" Then
        Select Case currentbase
            Case 1
                Text1.Text = Hex(Val(Text1.Text))
            Case 3
                Text1.Text = Hex(OctToDec(Text1.Text))
            Case 4
                Text1.Text = Hex(BinToDec(Text1.Text))
        End Select
End If
        currentbase = 2
        Text1.SetFocus
    Exit Sub
errorHandler:
  Text1.Text = "ERROR"
 MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error"  'the vbCrlf constant inserts a line break between the literal string and the error's description
 ClearDisplay = True
 currentbase = 2
 Text1.SetFocus

End Sub

Private Sub optoct_Click()
    Text2.Text = "" 'text2 is an invisible textbox as a temporary variable for the binary conversion
    cmdA.Enabled = False
    cmdB.Enabled = False
    cmdC.Enabled = False
    cmdD.Enabled = False
    cmdE.Enabled = False
    cmdF.Enabled = False
    Digits(0).Enabled = True
    Digits(1).Enabled = True
    Digits(2).Enabled = True
    Digits(3).Enabled = True
    Digits(4).Enabled = True
    Digits(5).Enabled = True
    Digits(6).Enabled = True
    Digits(7).Enabled = True
    Digits(8).Enabled = False
    Digits(9).Enabled = False
   
   On Error GoTo errorHandler
If Text1.Text <> "" Then
         Select Case currentbase
        Case 1
            Text1.Text = Oct(Val(Text1.Text))
        Case 2
            Text1.Text = Oct(HexToDec(Text1.Text))
        Case 4
            Text1.Text = Oct(BinToDec(Text1.Text))
            
    End Select
End If
    currentbase = 3
    Text1.SetFocus
Exit Sub

errorHandler:
    Text1.Text = "ERROR"
    MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error"  'the vbCrlf constant inserts a line break between the literal string and the error's description
    ClearDisplay = True
    currentbase = 3
    Text1.SetFocus

End Sub

Private Sub optbin_Click()
   If ClearDisplay Then
        Text2.Text = ""
   End If
    cmdA.Enabled = False
    cmdB.Enabled = False
    cmdC.Enabled = False
    cmdD.Enabled = False
    cmdE.Enabled = False
    cmdF.Enabled = False
    Digits(0).Enabled = True
    Digits(1).Enabled = True
    Digits(2).Enabled = False
    Digits(3).Enabled = False
    Digits(4).Enabled = False
    Digits(5).Enabled = False
    Digits(6).Enabled = False
    Digits(7).Enabled = False
    Digits(8).Enabled = False
    Digits(9).Enabled = False
   On Error GoTo errorHandler
Select Case currentbase
        Case 1
            DecToBin (Val(Text1.Text))
            Text1.Text = Text2.Text
        Case 2
            DecToBin (HexToDec(Text1.Text))
            Text1.Text = Text2.Text
        Case 3
            DecToBin (OctToDec(Text1.Text))
            Text1.Text = Text2.Text
    End Select
    currentbase = 4
    Text1.SetFocus
Exit Sub

errorHandler:
  Text1.Text = "ERROR"
  MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error"  'the vbCrlf constant inserts a line break between the literal string and the error's description
  ClearDisplay = True
  currentbase = 4
  Text1.SetFocus

End Sub

Private Sub Text1_Change()
    Text1.SetFocus
    Text1.Text = UCase$(Text1.Text)
End Sub

Private Sub Text1_GotFocus()
    
    Text1.SelStart = Len(Text1.Text)

End Sub

Private Sub text1_KeyDown(KeyCode As Integer, Shift As Integer)
If mustClear = True Then
      mustClear = False
      Text1.Text = ""
End If
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
        
            If currentbase = 1 Then
            Select Case KeyAscii
                       Case 8 '8 =ASCII of backspace
                         backspace_Click
                        KeyAscii = 0
                   Case Else
                 If (KeyAscii < 48 Or KeyAscii > 57) Then 'the ASCII code of numeric digit start at 48 (for the digit 0)and end at 57 (for the digit 9). I allow the user to enter only numeric digits in the text1
                 KeyAscii = 0
                 End If
                 End Select
                
            End If
         If currentbase = 2 Then
        Select Case KeyAscii
                       Case 8 '8 =ASCII of backspace
                             backspace_Click
                            KeyAscii = 0
                       Case Else
                     If (KeyAscii < 48 Or KeyAscii > 57) And ((KeyAscii < 65 Or KeyAscii > 70)) Then 'the ASCII code of numeric digit starts at 48 (for the digit 0)and end at 57 (for the digit 9). I allow the user to enter only numeric digits in the text1
                        KeyAscii = 0
                     End If
          End Select
          End If
         If currentbase = 3 Then
            Select Case KeyAscii
            Case 8
               backspace_Click
               KeyAscii = 0
               Case Else
             If (KeyAscii < 48 Or KeyAscii > 55) Then
               KeyAscii = 0
            End If
            End Select
         End If
         If currentbase = 4 Then
            Select Case KeyAscii
                       Case 8 '8 =ASCII of backspace
                             backspace_Click
                            KeyAscii = 0
                       Case Else
                     If (KeyAscii < 48 Or KeyAscii > 49) Then 'the ASCII code of numeric digit starts at 48 (for the digit 0)and end at 57 (for the digit 9). I allow the user to enter only numeric digits in the text1
                        KeyAscii = 0
                     End If
            End Select
         End If
End Sub

