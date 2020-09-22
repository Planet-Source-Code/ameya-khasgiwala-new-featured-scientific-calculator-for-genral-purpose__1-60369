VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ameya's Calculator"
   ClientHeight    =   3840
   ClientLeft      =   2205
   ClientTop       =   3105
   ClientWidth     =   7110
   Icon            =   "Calc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7110
   Begin VB.TextBox txtOP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtANS 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtNO2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtNO1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   120
      Width           =   2055
   End
   Begin VB.CheckBox chkINV 
      Caption         =   "INV"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   42
      Top             =   1200
      Width           =   735
   End
   Begin VB.CheckBox chkHYP 
      Caption         =   "HYP"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   41
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton btnPOW 
      Caption         =   "x^y"
      Height          =   375
      Left            =   120
      TabIndex        =   40
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton btnPOW2 
      Caption         =   "x^2"
      Height          =   375
      Left            =   840
      TabIndex        =   39
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton btnINV 
      Caption         =   "1/x"
      Height          =   375
      Left            =   1560
      TabIndex        =   38
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton btnTAN 
      Caption         =   "tan"
      Height          =   375
      Left            =   1560
      TabIndex        =   37
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton btnCOS 
      Caption         =   "cos"
      Height          =   375
      Left            =   840
      TabIndex        =   36
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton btnSIN 
      Caption         =   "sin"
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton btnCOT 
      Caption         =   "cot"
      Height          =   375
      Left            =   1560
      TabIndex        =   34
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton btnSEC 
      Caption         =   "sec"
      Height          =   375
      Left            =   840
      TabIndex        =   33
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton btnCOSEC 
      Caption         =   "cosec"
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton btnPOW3 
      Caption         =   "x^3"
      Height          =   375
      Left            =   1560
      TabIndex        =   31
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton btnFACTORIAL 
      Caption         =   "x!"
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton btnLN 
      Caption         =   "ln"
      Height          =   375
      Left            =   2280
      TabIndex        =   29
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton btnEXP 
      Caption         =   "e^x"
      Height          =   375
      Left            =   2280
      TabIndex        =   28
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton btnPI 
      Caption         =   "Pi()"
      Height          =   375
      Left            =   2280
      TabIndex        =   27
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton btnSQRT 
      Caption         =   "sqrt(x)"
      Height          =   375
      Left            =   840
      TabIndex        =   26
      Top             =   3120
      Width           =   615
   End
   Begin VB.CommandButton btnLOG 
      Caption         =   "log"
      Height          =   375
      Left            =   2280
      TabIndex        =   25
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton btnce 
      Caption         =   "CE"
      Height          =   495
      Left            =   5400
      TabIndex        =   24
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton btnbksp 
      Caption         =   "BackSpace"
      Height          =   495
      Left            =   4080
      TabIndex        =   23
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton btnC 
      Caption         =   "C"
      Height          =   495
      Left            =   6240
      TabIndex        =   22
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton btnor 
      Caption         =   "OR"
      Height          =   375
      Left            =   6480
      TabIndex        =   21
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnMOD 
      Caption         =   "%"
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton btnnot 
      Caption         =   "NOT"
      Height          =   375
      Left            =   6480
      TabIndex        =   19
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton btnand 
      Caption         =   "AND"
      Height          =   375
      Left            =   6480
      TabIndex        =   18
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btnsub 
      Caption         =   "-"
      Height          =   375
      Left            =   5880
      TabIndex        =   17
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btnmul 
      Caption         =   "*"
      Height          =   375
      Left            =   5880
      TabIndex        =   16
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton btndiv 
      Caption         =   "/"
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton btnEqual 
      Caption         =   "="
      Height          =   375
      Left            =   5280
      TabIndex        =   14
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "show"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton btnadd 
      Caption         =   "+"
      Height          =   375
      Left            =   5880
      TabIndex        =   12
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btndot 
      Caption         =   "."
      Height          =   375
      Left            =   4080
      TabIndex        =   11
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton btn3 
      Caption         =   "3"
      Height          =   375
      Left            =   5280
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton btn4 
      Caption         =   "4"
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btn5 
      Caption         =   "5"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btn6 
      Caption         =   "6"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   2280
      Width           =   495
   End
   Begin VB.CommandButton btn7 
      Caption         =   "7"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btn8 
      Caption         =   "8"
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btn9 
      Caption         =   "9"
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton btn0 
      Caption         =   "0"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton btn2 
      Caption         =   "2"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton btn1 
      Caption         =   "1"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   2760
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   0
         Format          =   $"Calc.frx":030A
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   " ="
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
      Left            =   4680
      TabIndex        =   46
      Top             =   120
      Width           =   255
   End
   Begin VB.Menu file 
      Caption         =   "File"
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "Edit"
      Begin VB.Menu cut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu helptopic 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu about 
         Caption         =   "About Ameya's Calculator"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public no1 As Double, no2 As Double, WhichNo As Boolean, Dot As Boolean, op As String, answer As Double, DotVal As Long
Sub textvalue(num As Long)
    On Error GoTo OvrFlowError
    If (Not WhichNo) Then
        If (Dot) Then
            DotVal = DotVal + 1
            temp = num
            For i = 1 To DotVal
                temp = (temp / 10)
            Next
            no1 = no1 + temp
        Else
            temp = no1 * 10
            no1 = temp + num
        End If
        Text1.Text = no1
    Else
        If (Dot) Then
            DotVal = DotVal + 1
            temp = num
            For i = 1 To DotVal
                temp = (temp / 10)
            Next
            no2 = no2 + temp
        Else
            temp = no2 * 10
            no2 = temp + num
        End If
        Text1.Text = no2
    End If
    txtNO1.Text = no1
    txtNO2.Text = no2
    txtOP.Text = op
    txtANS.Text = answer
    Exit Sub
OvrFlowError:
    MsgBox "Overflow occured!" & Chr(13) & "Please restart your job.", vbExclamation, "Error - Ameya's Calculator"
End Sub
Sub calc()
    On Error GoTo aritherror
    Select Case op
    Case "+"
        answer = (no1 + no2)
    Case "-"
        answer = (no1 - no2)
    Case "*"
        answer = (no1 * no2)
    Case "/"
        answer = (no1 / no2)
    Case "%"
        'answer = (no1 % no2)
    Case "&"
        answer = (no1 & no2)
    Case "|"
        answer = (no1 & no2)
    Case "~"
        answer = (Not no1)
    Case "sin"
        answer = (Sin(no1))
    Case "cos"
        answer = (Cos(no1))
    Case "tan"
        answer = (Tan(no1))
    Case "cosec"
        answer = (1 / Sin(no1))
    Case "sec"
        answer = (1 / Cos(no1))
    Case "cot"
        answer = (1 / Tan(no1))
    Case "ln"
        answer = (Log(no1))
    Case "log"
        answer = (Log(no1) / 2.30258509299405)
    Case "^"
        answer = (no1 ^ no2)
    Case "!"
        answer = 1
        For i = 2 To no1
            answer = answer * i
        Next
    End Select
    txtOP.Text = ""
    WhichNo = False
    Text1.Text = answer
    txtNO1.Text = no1
    txtNO2.Text = no2
    txtOP.Text = op
    txtANS.Text = answer
    no1 = answer
    Exit Sub
aritherror:
    MsgBox "Arithmetic error occured!. Possibly Overflow." & Chr(13) & "Please restart your job.", vbExclamation, "Error - Ameya's Calculator"
End Sub
Private Sub error(errorno As Long)
    Select Case errorno
    Case 1
        MsgBox "Divide by zero error!"
    Case 2
        MsgBox "Operator Overflow!"
    End Select
End Sub

Private Sub about_Click()
    frmAbout.Show
End Sub
Private Sub btn1_Click()
    textvalue (1)
End Sub

Private Sub btn2_Click()
    textvalue (2)
End Sub
Private Sub btn3_Click()
    textvalue (3)
End Sub
Private Sub btn4_Click()
    textvalue (4)
End Sub
Private Sub btn5_Click()
    textvalue (5)
End Sub
Private Sub btn6_Click()
    textvalue (6)
End Sub
Private Sub btn7_Click()
    textvalue (7)
End Sub
Private Sub btn8_Click()
    textvalue (8)
End Sub
Private Sub btn9_Click()
    textvalue (9)
End Sub
Private Sub btn0_Click()
    textvalue (0)
End Sub
Private Sub btnADD_Click()
    op = "+"
    WhichNo = True
    Text1.Text = ""
    DotVal = 0
    Dot = False
    no2 = 0
End Sub
Private Sub btnSUB_Click()
    op = "-"
    WhichNo = True
    Text1.Text = ""
    DotVal = 0
    Dot = False
    no2 = 0
End Sub
Private Sub btnMUL_Click()
    op = "*"
    WhichNo = True
    Text1.Text = ""
    DotVal = 0
    Dot = False
    no2 = 0
End Sub
Private Sub btnDIV_Click()
    op = "/"
    WhichNo = True
    Text1.Text = ""
    DotVal = 0
    Dot = False
    no2 = 0
End Sub
Private Sub btnAND_Click()
    op = "&"
    WhichNo = True
    Text1.Text = ""
    DotVal = 0
    Dot = False
End Sub
Private Sub btnOR_Click()
    op = "|"
    WhichNo = True
    Text1.Text = ""
    DotVal = 0
    Dot = False
End Sub
Private Sub btnMOD_Click()
    op = "%"
    WhichNo = True
    Text1.Text = ""
    DotVal = 0
    Dot = False
End Sub
Private Sub btnNOT_Click()
    op = "~"
    WhichNo = True
    Text1.Text = ""
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnSIN_Click()
    op = ""
    If (chkINV.Value = 1) Then
        op = "a"
    End If
    op = op & "sin"
    If (chkHYP.Value = 1) Then
        op = op & "h"
    End If
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnCOS_Click()
    op = ""
    If (chkINV.Value = 1) Then
        op = "a"
    End If
    op = op & "cos"
    If (chkHYP.Value = 1) Then
        op = op & "h"
    End If
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnTAN_Click()
    op = ""
    If (chkINV.Value = 1) Then
        op = "a"
    End If
    op = op & "tan"
    If (chkHYP.Value = 1) Then
        op = op & "h"
    End If
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnCOSEC_Click()
    op = ""
    If (chkINV.Value = 1) Then
        op = "a"
    End If
    op = op & "cosec"
    If (chkHYP.Value = 1) Then
        op = op & "h"
    End If
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnSEC_Click()
    op = ""
    If (chkINV.Value = 1) Then
        op = "a"
    End If
    op = op & "sec"
    If (chkHYP.Value = 1) Then
        op = op & "h"
    End If
    Text1.Text = ""
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnCOT_Click()
    op = ""
    If (chkINV.Value = 1) Then
        op = "a"
    End If
    op = op & "cot"
    If (chkHYP.Value = 1) Then
        op = op & "h"
    End If
    Text1.Text = ""
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnLN_Click()
    If (no1 <= 0) Then
        MsgBox ("logarithm is only defined for positive numbers." & Chr(13) & "Please enter a valid no and then take logarithm.")
    Else
        op = "ln"
        Text1.Text = ""
        DotVal = 0
        Dot = False
        calc
    End If
End Sub
Private Sub btnLOG_Click()
    If (no1 <= 0) Then
        MsgBox ("logarithm is only defined for positive numbers." & Chr(13) & "Please enter a valid no and then take logarithm.")
    Else
        op = "log"
        Text1.Text = ""
        DotVal = 0
        Dot = False
        calc
    End If
End Sub
Private Sub btnPOW_Click()
    op = "^"
    WhichNo = True
    Text1.Text = ""
    DotVal = 0
    Dot = False
End Sub
Private Sub btnPOW2_Click()
    op = "^"
    no2 = 2
    Text1.Text = ""
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnPOW3_Click()
    op = "^"
    no2 = 3
    Text1.Text = ""
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnINV_Click()
    op = "^"
    no2 = -1
    Text1.Text = ""
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnFACTORIAL_Click()
    op = "!"
    Text1.Text = ""
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnSQRT_Click()
    If (no1 < 0) Then
        MsgBox ("Square Root is not defined for negative numbers.")
    Else
        op = "^"
        no2 = 0.5
        Text1.Text = ""
        DotVal = 0
        Dot = False
        calc
    End If
End Sub
Private Sub btnEXP_Click()
    op = "^"
    no2 = no1
    no1 = 2.30258509299405
    Text1.Text = ""
    DotVal = 0
    Dot = False
    calc
End Sub
Private Sub btnPI_Click()
    If (Not WhichNo) Then
        no1 = 3.14159265358979
    Else
        no2 = 3.14159265358979
    End If
    Text1.Text = "3.14159265358979"
    DotVal = 0
    Dot = False
End Sub
Private Sub btnC_Click()
    no1 = 0
    no2 = 0
    answer = 0
    Text1.Text = "0"
    op = ""
    DotVal = 0
    Dot = False
    WhichNo = False
    txtNO1.Text = no1
    txtNO2.Text = no2
    txtOP.Text = op
    txtANS.Text = answer
End Sub
Private Sub btnCE_Click()
    If (temp) Then
        no1 = 0
    Else
        no2 = 0
    End If
    Text1.Text = ""
    DotVal = 0
    Dot = False
End Sub
Private Sub btnbksp_Click()
    If (Not WhichNo) Then
        If (Len(Str(no1)) > 1) Then
            no1 = FormatNumber(Left(Str(no1), Len(Text1.Text) - 1))
            Text1.Text = no1
        End If
    Else
        If (Len(Str(no2)) > 0) Then
            no2 = FormatNumber(Left(Str(no2), Len(Text1.Text) - 1))
            Text1.Text = no2
        End If
    End If
    If (DotVal > 0) Then
        DotVal = ditval - 1
    End If
    txtNO1.Text = no1
    txtNO2.Text = no2
    txtOP.Text = op
    txtANS.Text = answer
End Sub
Private Sub btnDOT_Click()
    If (Dot = False) Then
        Dot = True
        Text1.Text = Text1.Text & "."
        DotVal = 0
    End If
End Sub
Private Sub btnEqual_Click()
    calc
    no1 = answer
    WhichNo = False
    DotVal = 0
    Dot = False
End Sub
Private Sub Command2_Click()
    MsgBox ("Temp=" & Str(temp) & " No1=" & Str(no1) & " No2=" & Str(no2) & " op=" & op)
End Sub
Private Sub copy_Click()
    Clipboard.SetText (Text1.Text)
End Sub
Private Sub cut_Click()
    Clipboard.SetText (Text1.Text)
    Text1.Text = ""
End Sub
Private Sub exit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    Dim no1, no2, op, WhichNo, Dot, temp
    no1 = 0
    no2 = 0
    answer = 0
    Dot = False
    DotVal = 0
    WhichNo = False
    btnC_Click
End Sub

Private Sub helptopic_Click()
        MsgBox ("No help found." & Chr(13) & "We are SORRY for the inconvenience!")
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    If (IsNumeric(Text1.Text)) Then
        If (Not WhichNo) Then
            no1 = FormatNumber(Text1.Text)
        Else
            no2 = FormatNumber(Text1.Text)
        End If
    End If
    txtNO1.Text = no1
    txtNO2.Text = no2
    txtOP.Text = op
    txtANS.Text = answer
End Sub
