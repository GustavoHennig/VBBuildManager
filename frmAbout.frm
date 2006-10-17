VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sobre..."
   ClientHeight    =   4065
   ClientLeft      =   4560
   ClientTop       =   3450
   ClientWidth     =   5730
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805.736
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   4245
      TabIndex        =   0
      Top             =   3570
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rtbMSG 
      Height          =   2505
      Left            =   120
      TabIndex        =   2
      Top             =   495
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   4419
      _Version        =   393217
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmAbout.frx":000C
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "gustavohe@gmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   150
      MouseIcon       =   "frmAbout.frx":008E
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3090
      Width           =   1620
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   2329.486
      Y2              =   2329.486
   End
   Begin VB.Label lblTitle 
      Caption         =   "Compilador VB"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   90
      TabIndex        =   1
      Top             =   75
      Width           =   5325
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   2339.839
      Y2              =   2339.839
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rc As New clsRecursos
    
    rtbMSG.TextRTF = rc.fg_getStringRecurso(cRTF_Sobre)
   ' rtbMSG.TextRTF
    Me.Caption = Me.Caption & "  V. " & App.Major & "." & App.Minor & "." & App.Revision
    Set rc = Nothing
End Sub

Private Sub lbl_Click()
    On Error Resume Next
    ShellExecute Me.hwnd, vbNullString, "mailto:gustavohe@gmail.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub
