VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Calfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daniel's Calculator"
   ClientHeight    =   3480
   ClientLeft      =   1335
   ClientTop       =   2325
   ClientWidth     =   3705
   FillStyle       =   0  'Solid
   Icon            =   "calculatrice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   14.501
   ScaleMode       =   0  'User
   ScaleWidth      =   30.875
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   5280
      TabIndex        =   42
      Top             =   360
      Width           =   855
      Begin VB.CheckBox checkinv 
         Caption         =   "Inv"
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   5280
      TabIndex        =   38
      Top             =   960
      Width           =   855
      Begin VB.OptionButton optgrad 
         Caption         =   "Gra"
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton optrad 
         Caption         =   "Rad"
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton optdeg 
         Caption         =   "Deg"
         ForeColor       =   &H00FF00FF&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Digits 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   720
      TabIndex        =   34
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton Digits 
      Caption         =   "1"
      Height          =   495
      Index           =   1
      Left            =   720
      TabIndex        =   33
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton DotBttn 
      Caption         =   "."
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
      Left            =   1920
      TabIndex        =   10
      Top             =   2880
      Width           =   495
   End
   Begin VB.CommandButton PlusMins 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   1320
      TabIndex        =   9
      Top             =   2880
      Width           =   495
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
      Left            =   1320
      TabIndex        =   8
      Top             =   2280
      Width           =   495
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
      Left            =   1920
      TabIndex        =   7
      Top             =   2280
      Width           =   495
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
      Left            =   720
      TabIndex        =   6
      Top             =   1680
      Width           =   495
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
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   495
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
      Left            =   1920
      TabIndex        =   4
      Top             =   1680
      Width           =   495
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
      Left            =   720
      TabIndex        =   3
      Top             =   1080
      Width           =   495
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
      Left            =   1320
      TabIndex        =   2
      Top             =   1080
      Width           =   495
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
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.TextBox textbox 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.000E+00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   6
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      MaxLength       =   25
      MousePointer    =   3  'I-Beam
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
   Begin MSForms.CommandButton cmdbaseconv 
      Height          =   495
      Left            =   3720
      TabIndex        =   37
      Top             =   480
      Width           =   1455
      ForeColor       =   16711680
      Caption         =   "Base Conversion"
      Size            =   "2566;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton ClearBttn 
      Height          =   495
      Left            =   2280
      TabIndex        =   36
      Top             =   480
      Width           =   1335
      ForeColor       =   255
      Caption         =   "C"
      Size            =   "2355;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdpi 
      Height          =   495
      Index           =   3
      Left            =   5280
      TabIndex        =   35
      Top             =   2880
      Width           =   855
      ForeColor       =   16711935
      Caption         =   "pi"
      Size            =   "1508;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdbackspace 
      Height          =   495
      Left            =   720
      TabIndex        =   32
      Top             =   480
      Width           =   1455
      ForeColor       =   255
      Caption         =   "Backspace"
      Size            =   "2566;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Equals 
      Height          =   495
      Index           =   24
      Left            =   3120
      TabIndex        =   31
      Top             =   2880
      Width           =   495
      ForeColor       =   255
      Caption         =   "="
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Over 
      Height          =   495
      Index           =   23
      Left            =   3120
      TabIndex        =   30
      Top             =   2280
      Width           =   495
      ForeColor       =   16711680
      Caption         =   "1/x"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdperc 
      Height          =   495
      Index           =   22
      Left            =   3120
      TabIndex        =   29
      Top             =   1680
      Width           =   495
      ForeColor       =   16711680
      Caption         =   "%"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton sqrcmd 
      Height          =   495
      Index           =   21
      Left            =   3120
      TabIndex        =   28
      Top             =   1080
      Width           =   495
      ForeColor       =   16711680
      Caption         =   "sqrt"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Plus 
      Height          =   495
      Index           =   20
      Left            =   2520
      TabIndex        =   27
      Top             =   2880
      Width           =   495
      ForeColor       =   255
      Caption         =   "+"
      Size            =   "873;873"
      FontHeight      =   195
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Minus 
      Height          =   495
      Index           =   19
      Left            =   2520
      TabIndex        =   26
      Top             =   2280
      Width           =   495
      ForeColor       =   255
      Caption         =   "-"
      Size            =   "873;873"
      FontHeight      =   195
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Times 
      Height          =   495
      Index           =   18
      Left            =   2520
      TabIndex        =   25
      Top             =   1680
      Width           =   495
      ForeColor       =   255
      Caption         =   "*"
      Size            =   "873;873"
      FontHeight      =   195
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton Div 
      Height          =   495
      Index           =   17
      Left            =   2520
      TabIndex        =   24
      Top             =   1080
      Width           =   495
      ForeColor       =   255
      BackColor       =   -2147483637
      Caption         =   "/"
      PicturePosition =   131072
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdexp 
      Height          =   495
      Index           =   13
      Left            =   4560
      TabIndex        =   23
      Top             =   2880
      Width           =   495
      ForeColor       =   255
      Caption         =   "x^n"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdsquare 
      Height          =   495
      Index           =   12
      Left            =   4560
      TabIndex        =   22
      Top             =   2280
      Width           =   495
      ForeColor       =   255
      Caption         =   "x^2"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdlog 
      Height          =   495
      Index           =   11
      Left            =   4560
      TabIndex        =   21
      Top             =   1680
      Width           =   495
      ForeColor       =   255
      Caption         =   "log"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdln 
      Height          =   495
      Index           =   10
      Left            =   4560
      TabIndex        =   20
      Top             =   1080
      Width           =   495
      ForeColor       =   255
      Caption         =   "ln"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdeng 
      Height          =   495
      Index           =   9
      Left            =   3840
      TabIndex        =   19
      Top             =   2880
      Width           =   495
      ForeColor       =   255
      Caption         =   "e"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton sincmd 
      Height          =   495
      Index           =   8
      Left            =   3840
      TabIndex        =   18
      Top             =   1080
      Width           =   495
      ForeColor       =   16711935
      Caption         =   "sin"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton coscmd 
      Height          =   495
      Index           =   7
      Left            =   3840
      TabIndex        =   17
      Top             =   1680
      Width           =   495
      ForeColor       =   16711935
      Caption         =   "cos"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdtan 
      Height          =   495
      Index           =   1
      Left            =   3840
      TabIndex        =   16
      Top             =   2280
      Width           =   495
      ForeColor       =   16711935
      Caption         =   "tan"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdmemo 
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   495
      ForeColor       =   255
      Caption         =   "MR"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdmemplus 
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2280
      Width           =   495
      ForeColor       =   255
      Caption         =   "M+"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdmeminus 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   495
      ForeColor       =   255
      Caption         =   "M-"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.CommandButton cmdmemclear 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   495
      ForeColor       =   255
      Caption         =   "MC"
      Size            =   "873;873"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin VB.Label memlbl 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   11
      Top             =   480
      Width           =   495
   End
   Begin VB.Menu filemnu 
      Caption         =   "&File"
      Begin VB.Menu exitmnu 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu viewmnu 
      Caption         =   "&View"
      WindowList      =   -1  'True
      Begin VB.Menu nrmaviewmnu 
         Caption         =   "S&tandard View"
      End
      Begin VB.Menu scviewmnu 
         Caption         =   "&Scientific View"
      End
      Begin VB.Menu basemnuview 
         Caption         =   "&Base Conversion"
      End
   End
   Begin VB.Menu aboutmnu 
      Caption         =   "&About"
      Begin VB.Menu abtauthormnu 
         Caption         =   "&About the Author"
      End
   End
End
Attribute VB_Name = "Calfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''
'Copyright© 2001-2002 FAC. Sc. & Comp. Eng. KASLIK'
'$Source:/cours/inf336/Calculator                 '
'$Version 0.3                                     '
'Date: Spring 2002                                '
'$Author: Daniel J. Khayat                        '
'''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Dim dOperand1, dOperand2, result As Double
Dim Operator As String
Dim ClearDisplay As Boolean
Dim Memo As Double
Dim sum As Double
Dim mustClear As Boolean
Dim check As Boolean
Const pi = 3.14159265358979
Dim flaj As Integer
Dim tempo As Double 'tempo is a temporary variable that holds the textbox content
'tempo is a variable to hold the value of the numerical number written in the textbox
Dim flagr, flag1, flag2 As Integer
' flagr to keep track that we passed through the function equal
' flag1 to keep track of dOperand1. When the Operand one is the first element typed in the
'textbox, then flag1 would be zero
' when we are doing a+b+c...the doperand1 will be "a" and doperand2 will be "b"
' then doperand1 will be "result" and doperand2 will be "c"
' flag1=0 means that doperand1 should be "a"
' flag2 to keep track of dOperand2

Private Sub abtauthormnu_Click()
    
    MsgBox "Copyright© 2001-2002 FAC. Sc. & Comp. Eng. KASLIK " & Chr(13) & "Source:/cours/inf336/Calculator" & Chr(13) & "Version 0.5" & Chr(13) & "Date: Spring 2002" & Chr(13) & "Author: Daniel J. Khayat", vbInformation
      
End Sub

Private Sub basemnuview_Click()

    Calfrm.Hide
    Base.Show

End Sub

Private Sub checkinv_Click()

    textbox.SetFocus

End Sub

Private Sub ClearBttn_Click()
    
    checkinv.value = False
    Form1_Load
    
End Sub

Private Sub cmdbackspace_Click()
    
    If Len(textbox.Text) < "1" Then
       textbox.Text = ""
    End If
        If Len(textbox.Text) = "1" Then
            textbox.Text = ""
        End If
    If Len(textbox.Text) > 1 Then
      textbox.Text = Left$(textbox.Text, Len(textbox.Text) - 1)
    End If

End Sub

Private Sub cmdbaseconv_Click()
    
    Calfrm.Hide
    Base.Show

End Sub

Private Sub cmdeng_Click(Index As Integer)
    'textbox.Text=log
    textbox.Text = 2.71828182845905
    ClearDisplay = True
    mustClear = True
    
End Sub

Private Sub cmdexp_Click(Index As Integer)
    
    dOperand1 = Val(textbox.Text) 'Converts the value of the textbox to a numeric value
    Operator = "x^n"
    ClearDisplay = True
    mustClear = True
    textbox.Text = ""

End Sub

Private Sub cmdln_Click(Index As Integer)

On Error GoTo errorHandler
    textbox.Text = Log(Val(textbox.Text))
    ClearDisplay = True
    mustClear = True
    Exit Sub

errorHandler:
    textbox.Text = "ERROR"
    MsgBox "The operation result in the following error " & vbCrLf & Err.Description, vbExclamation, " Error" 'the vbCrlf constant inserts a line break between the literal string and the error's description
    ClearDisplay = True

End Sub
Private Sub cmdlog_Click(Index As Integer)

On Error GoTo errorHandler
        textbox.Text = (Log(Val(textbox.Text))) / (Log(10)) 'The formulas to convert from the base e to base 10
        ClearDisplay = True
        mustClear = True
        Exit Sub

errorHandler:
 textbox.Text = "ERROR"
 MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error"  'the vbCrlf constant inserts a line break between the literal string and the error's description
 ClearDisplay = True

End Sub

Private Sub cmdmemclear_Click(Index As Integer)
    
    Memo = 0
    memlbl.Caption = ""
    textbox.SetFocus

End Sub

Private Sub cmdmeminus_Click(Index As Integer)

    If Len(textbox.Text) = 0 Then
        textbox.Text = ""
    Else: Memo = Memo - textbox.Text
        If (Memo <> 0) Then
            memlbl.Caption = "M"
        End If
    End If

End Sub

Private Sub cmdmemo_Click(Index As Integer)
 
textbox.Text = Memo
 If (Memo <> 0) Then
    memlbl.Caption = "M"
 End If

End Sub

Private Sub cmdmemplus_Click(Index As Integer)

If Len(textbox.Text) = 0 Then
    textbox.Text = ""
Else: Memo = Memo + textbox.Text
    If (Memo <> 0) Then
     memlbl.Caption = "M"
    End If
End If

End Sub
Private Sub cmdperc_Click(Index As Integer)

If (flag2 = 0 And flagr = 1) Then GoTo NOOP5
    tempo = Val(textbox.Text)
    flagr = 0
    
    If flag1 = 0 Then
        dOperand1 = tempo 'the first time we press this operator
        flag1 = flag1 + 1
    
    ElseIf (flagr <> 1 And flag1 > 1) Then
        dOperand2 = tempo
        flag2 = 1 'to know that doperand2 has been assigned a value before we clear the texbox
    End If
    
NOOP5: flag1 = flag1 + 1
    textbox.Text = ""
    Operator = "%"
    If (flag2 = 1) Then
        flag2 = 2
        Equals_Click (flag2)
    End If
mustClear = True
    
End Sub
Private Sub cmdpi_Click(Index As Integer)
    
    textbox.Text = 4 * Atn(1)
    ClearDisplay = True
    mustClear = True

End Sub
Private Sub cmdsquare_Click(Index As Integer)

On Error GoTo errorHandler
  textbox.Text = textbox.Text * textbox.Text
  ClearDisplay = True
  mustClear = True
  Exit Sub

errorHandler:
 textbox.Text = "ERROR"
 MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error"  'the vbCrlf constant inserts a line break between the literal string and the error's description
 ClearDisplay = True

End Sub
Private Sub cmdtan_Click(Index As Integer)
 
 On Error GoTo errorHandler
  If optdeg.value Then
            If checkinv.value Then
                        textbox.Text = Atn(Val(textbox.Text)) * 180 / pi
                        ClearDisplay = True
                        mustClear = True
                        checkinv.value = False
            
            Else
                        textbox.Text = Tan(Val(textbox.Text) * (pi) / (180)) 'I did this multiplication by pi and the division by 180 to convert from radian to degree coz the internal function cos gives the result cos in radian
                        ClearDisplay = True
                        mustClear = True
            End If
            Exit Sub
  End If
  
  If optrad.value Then
            If checkinv.value Then
                        textbox.Text = Atn(Val(textbox.Text))
                        ClearDisplay = True
                        mustClear = True
                        checkinv.value = False
            
            Else
                        textbox.Text = Tan(Val(textbox.Text))
                        ClearDisplay = True
                        mustClear = True
            End If
            Exit Sub
 End If
   
   If optgrad.value Then
            If checkinv.value Then
                        textbox.Text = Atn(Val(textbox.Text)) * 200 / pi
                        ClearDisplay = True
                        mustClear = True
                        checkinv.value = False
            
            Else
                        textbox.Text = Tan(Val(textbox.Text) * (pi / 200)) 'I did this multiplication by pi and the division by 200 to convert from radian to grad coz the internal function sin gives the result sin in radian
                        ClearDisplay = True
                        mustClear = True
            End If
            Exit Sub
End If

errorHandler:
                        textbox.Text = "ERROR"
                        MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error" 'the vbCrlf constant inserts a line break between the literal string and the error's description
                        ClearDisplay = True
End Sub

Private Sub coscmd_Click(Index As Integer)
 
 On Error GoTo errorHandler
  If optdeg.value Then
            If checkinv.value Then
                        textbox.Text = ((Atn(Val(-textbox.Text) / Sqr(-Val(textbox.Text) * Val(textbox.Text) + 1))) + 2 * Atn(1)) * 180 / pi
                        ClearDisplay = True
                        mustClear = True
                        checkinv.value = False
            
            Else
                        textbox.Text = Cos(Val((textbox.Text) * pi) / (180)) 'I did this multiplication by pi and the division by 180 to convert from radian to degree coz the internal function cos gives the result cos in radian
                        ClearDisplay = True
                        mustClear = True
            End If
            Exit Sub
  End If
  
  If optrad.value Then
            If checkinv.value Then
                        textbox.Text = ((Atn(Val(-textbox.Text) / Sqr(-Val(textbox.Text) * Val(textbox.Text) + 1))) + 2 * Atn(1))
                        ClearDisplay = True
                        mustClear = True
                        checkinv.value = False
            
            Else
                        textbox.Text = Cos(Val(textbox.Text))
                        ClearDisplay = True
                        mustClear = True
            End If
            Exit Sub
 End If
  
  If optgrad.value Then
            If checkinv.value Then
                        textbox.Text = ((Atn(Val(-textbox.Text) / Sqr(-Val(textbox.Text) * Val(textbox.Text) + 1))) + 2 * Atn(1)) * 200 / pi
                        ClearDisplay = True
                        mustClear = True
                        checkinv.value = False
            
            Else
                        textbox.Text = Cos(Val(textbox.Text) * (pi / 200))  'I did this multiplication by pi and the division by 200 to convert from radian to grad coz the internal function sin gives the result sin in radian
                        ClearDisplay = True
                        mustClear = True
            End If
            Exit Sub
End If

errorHandler:
                        textbox.Text = "ERROR"
                        MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error" 'the vbCrlf constant inserts a line break between the literal string and the error's description
                        ClearDisplay = True
  
End Sub

Private Sub Digits_Click(Index As Integer)
  
  If ClearDisplay Then
    textbox.Text = ""
    ClearDisplay = False
  End If
textbox.Text = textbox.Text + Digits(Index).Caption
textbox.SetFocus

End Sub

Private Sub DotBttn_Click(Index As Integer)
        
        If InStr(textbox.Text, ".") Then 'InStr Function returns a value that indicates the position of the Dot
            textbox.SetFocus
            Exit Sub
        Else
            textbox.Text = textbox.Text + "."
        End If
textbox.SetFocus

End Sub

Private Sub Equals_Click(Index As Integer)

If (flag2 = 0) Then dOperand2 = Val(textbox.Text)
'to know that doperand2 was assigned a value before we clear the textbox
flagr = 1 'to know that we passed in the equals function
On Error GoTo errorHandler
    Select Case Operator
    Case "+"
        result = dOperand1 + dOperand2
    Case "-"
        result = dOperand1 - dOperand2
    Case "*"
        result = dOperand1 * dOperand2
    Case "x^n"
        result = dOperand1 ^ dOperand2
    Case "%"
        result = dOperand1 * dOperand2 / 100
    Case "/"
        result = dOperand1 / dOperand2
    End Select
   
   'If Operator = "/" Then    'condition that we do not devide by zero to avoid a runtime error
   dOperand1 = result
   textbox.Text = result
   ClearDisplay = True
   mustClear = True
   Exit Sub
errorHandler:
  textbox.Text = "ERROR"
 MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error" 'the vbCrlf constant inserts a line break between the literal string and the error's description
 ClearDisplay = True
  End Sub

Private Sub Form1_Load()
    
    tempo = 0
    flagr = 0
    flag1 = 0
    flag2 = 0
    dOperand1 = 0
    dOperand2 = 0
    result = 0
    textbox.SetFocus
    textbox.Text = ""
    mustClear = False

End Sub

Private Sub exitmnu_Click()
    
    End

End Sub

Private Sub Form_Load()
    
    mustClear = False
    optdeg.value = True

End Sub

Private Sub Minus_Click(Index As Integer)
    If (flag2 = 0 And flagr = 1) Then GoTo NOOP2
        tempo = Val(textbox.Text)
        flagr = 0
    If flag1 = 0 Then
        dOperand1 = tempo 'the first time we press this operator
        flag1 = flag1 + 1
    ElseIf (flagr <> 1 And flag1 > 1) Then
        dOperand2 = tempo
        flag2 = 1 'to know that doperand2 has been assigned a value before we clear the texbox
    End If

NOOP2: flag1 = flag1 + 1
    Operator = "-"
    If (flag2 = 1) Then
        flag2 = 2
        Equals_Click (flag2)
    End If
    
    ClearDisplay = True
    mustClear = True

End Sub

Private Sub nrmaviewmnu_Click()
    
    Me.Width = 3810
    Me.ScaleHeight = 15
    textbox.Width = 29.125

End Sub

Private Sub optdeg_Click()
'    textbox.SetFocus
End Sub

Private Sub optgrad_Click()
    textbox.SetFocus
End Sub

Private Sub optrad_Click()
    textbox.SetFocus
End Sub

Private Sub Over_Click(Index As Integer)

On Error GoTo errorHandler
  textbox.Text = 1 / Val(textbox.Text)
  ClearDisplay = True
  mustClear = True
Exit Sub
errorHandler:
  textbox.Text = "ERROR"
 MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error" 'the vbCrlf constant inserts a line break between the literal string and the error's description
   
End Sub

Private Sub Plus_Click(Index As Integer)
    
    If (flag2 = 0 And flagr = 1) Then GoTo NOOP1 'this means that we did a+b= and then we are doing + again
    tempo = Val(textbox.Text)
    flagr = 0
    
    If flag1 = 0 Then
        dOperand1 = tempo 'the first time we press this operator
        flag1 = flag1 + 1
    ElseIf (flagr <> 1 And flag1 > 1) Then
    'that means we are doing the a+b dOperand1=a and dOperand2=b
    dOperand2 = tempo
    flag2 = 1 'to know that doperand2 has been assigned a value before we clear the texbox
    End If
    
NOOP1: flag1 = flag1 + 1
    Operator = "+"
    
    If (flag2 = 1) Then
        flag2 = 2 'this is to know that the equals after pressing "+" the second time like a+b"+"c
        ' this differentiate it from the action when we press "="
        Equals_Click (flag2)
    End If
    
    mustClear = True
    ClearDisplay = True
    textbox.SetFocus
    
End Sub
    

Private Sub PlusMins_Click(Index As Integer)
   If Len(textbox.Text) = 0 Then
        textbox.Text = ""
  Else: textbox.Text = -Val(textbox.Text)
  
  End If

End Sub
Private Sub scviewmnu_Click()

    Me.Width = 6330
    textbox.Width = 50.125

End Sub

Private Sub sincmd_Click(Index As Integer)

On Error GoTo errorHandler
  If optdeg.value Then
            If checkinv.value Then
                        textbox.Text = (Atn(Val(textbox.Text) / Sqr(-Val(textbox.Text) * Val(textbox.Text) + 1))) * 180 / pi
                        ClearDisplay = True
                        mustClear = True
                        checkinv.value = False
            
            Else
                        textbox.Text = Sin(Val(textbox.Text) * (pi) / (180)) 'I did this multiplication by pi and the division by 180 to convert from radian to degree coz the internal function sin gives the result sin in radian
                        ClearDisplay = True
                        mustClear = True
            End If
            Exit Sub
  End If
  If optrad.value Then
            If checkinv.value Then
                        textbox.Text = (Atn(Val(textbox.Text) / Sqr(-Val(textbox.Text) * Val(textbox.Text) + 1)))
                        ClearDisplay = True
                        mustClear = True
                        checkinv.value = False
            
            Else
                        textbox.Text = Sin(Val(textbox.Text))  'I did this multiplication by pi and the division by 180 to convert from radian to degree coz the internal function sin gives the result sin in radian
                        ClearDisplay = True
                        mustClear = True
            End If
            Exit Sub
 End If
   If optgrad.value Then
            If checkinv.value Then
                        textbox.Text = (Atn(Val(textbox.Text) / Sqr(-Val(textbox.Text) * Val(textbox.Text) + 1))) * 200 / pi
                        ClearDisplay = True
                        mustClear = True
                        checkinv.value = False
            
            Else
                        textbox.Text = Sin(Val(textbox.Text) * (pi) / (200)) 'I did this multiplication by pi and the division by 180 to convert from radian to degree coz the internal function sin gives the result sin in radian
                        ClearDisplay = True
                        mustClear = True
            End If
            Exit Sub
End If
errorHandler:
                        textbox.Text = "ERROR"
                        MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error" 'the vbCrlf constant inserts a line break between the literal string and the error's description
                        ClearDisplay = True
  
End Sub
Private Sub sqrcmd_Click(Index As Integer)

On Error GoTo errorHandler
    If Len(textbox.Text) = 0 Then
        textbox.Text = ""
    Else: textbox.Text = Sqr(Val(textbox.Text))
        ClearDisplay = True
        mustClear = True
    End If
    Exit Sub

errorHandler:
 textbox.Text = "ERROR"
 MsgBox "The operation resulted in the following error " & vbCrLf & Err.Description, vbExclamation, " Error" 'the vbCrlf constant inserts a line break between the literal string and the error's description
 ClearDisplay = True

End Sub

Private Sub textbox_GotFocus()

textbox.SelStart = Len(textbox.Text)
'If optdeg.value Then
'    textbox.SetFocus
'End If

End Sub

Private Sub textbox_KeyDown(KeyCode As Integer, Shift As Integer)

If mustClear = True Then
      mustClear = False
      textbox.Text = ""
End If

End Sub

Private Sub textbox_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 42 '42=ASCII of *
            KeyAscii = 0
            Times_Click (flag2)
        Case 43 '43=ASCII of +
            KeyAscii = 0
             Plus_Click (flag2)
        Case 45 '45=ASCII of -
            KeyAscii = 0
             Minus_Click (flag2)
        Case 47 '47=ASCII of /
            KeyAscii = 0
             Div_Click (flag2)
        Case 13 '13=ASCII of return and 61=ASCII of =
            KeyAscii = 0
             Equals_Click (flag2)
        Case 8 '8 =ASCII of backspace
             cmdbackspace_Click
            KeyAscii = 0
        Case 46 '46= ASCII of dot
             DotBttn_Click (flag2)
            KeyAscii = 0
        Case Else
     If (KeyAscii < 48 Or KeyAscii > 57) Then 'the ASCII code of numeric digit start at 48 (for the digit 0)and end at 57 (for the digit 9). I allow the user to enter only numeric digits in the texbox
     'textbox.Text = "0"
        KeyAscii = 0
     End If
     End Select
     textbox.SelStart = Len(textbox.Text)
End Sub

Private Sub Times_Click(Index As Integer)
    If (flag2 = 0 And flagr = 1) Then GoTo NOOP3
    
    tempo = Val(textbox.Text)
    flagr = 0
    If flag1 = 0 Then
        dOperand1 = tempo 'the first time we press this operator
        flag1 = flag1 + 1
    ElseIf (flagr <> 1 And flag1 > 1) Then
        dOperand2 = tempo
        flag2 = 1 'to know that doperand2 has been assigned a value before we clear the texbox
    End If

NOOP3: flag1 = flag1 + 1
    Operator = "*"
    If (flag2 = 1) Then
        flag2 = 2
        Equals_Click (flag2)
    End If
mustClear = True
ClearDisplay = True
End Sub
    
Private Sub Div_Click(Index As Integer)
    
    If (flag2 = 0 And flagr = 1) Then GoTo NOOP4
    tempo = Val(textbox.Text)
    flagr = 0
    If flag1 = 0 Then
        dOperand1 = tempo 'the first time we press this operator
        flag1 = flag1 + 1
    ElseIf (flagr <> 1 And flag1 > 1) Then
    dOperand2 = tempo
    flag2 = 1 'to know that doperand2 has been assigned a value before we clear the texbox
    End If

NOOP4: flag1 = flag1 + 1
    Operator = "/"
    If (flag2 = 1) Then
        flag2 = 2
        Equals_Click (flag2)
    End If
    
mustClear = True
ClearDisplay = True

End Sub

Public Sub type_times()
    
    Times_Click (1)

End Sub


