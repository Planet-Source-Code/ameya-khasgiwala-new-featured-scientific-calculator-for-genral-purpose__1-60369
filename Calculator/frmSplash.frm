VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4650
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4410
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   200
         Left            =   120
         Top             =   240
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4560
         TabIndex        =   10
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Please Wait...       Loading"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   4080
         Width           =   2295
      End
      Begin VB.Shape Shape2 
         Height          =   315
         Left            =   3690
         Top             =   4050
         Width           =   3055
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         FillColor       =   &H80000018&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3720
         Top             =   4080
         Width           =   15
      End
      Begin VB.Image imgLogo 
         Height          =   1785
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   795
         Width           =   1815
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright(c). All Rights Reserved"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   4
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Ameya Companies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   3
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         Caption         =   $"frmSplash.frx":0316
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   150
         TabIndex        =   2
         Top             =   3480
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version 1.0.0.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4710
         TabIndex        =   5
         Top             =   2640
         Width           =   2145
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Windows"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5475
         TabIndex        =   6
         Top             =   2280
         Width           =   1380
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Calculator 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2520
         TabIndex        =   8
         Top             =   1140
         Width           =   4230
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         Caption         =   "free version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Ameya Companies"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2355
         TabIndex        =   7
         Top             =   705
         Width           =   3195
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public i As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    i = 0
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    Form1.Show
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Timer1_Timer()
    If (i < 10) Then
        i = i + 1
        Shape1.Width = (i) * 300
        Label2.Caption = Str(((i + 1) * 10) - 10) & "%"
    End If
    If (i >= 10) Then
        i = i + 1
    End If
    If (i > 20) Then
        Unload Me
    End If
End Sub
