VERSION 5.00
Begin VB.Form frmResultado 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "Resultado da compilação"
   ClientHeight    =   6285
   ClientLeft      =   3345
   ClientTop       =   2535
   ClientWidth     =   7140
   Icon            =   "frmResultado.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   7140
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   5685
      TabIndex        =   1
      Top             =   5760
      Width           =   1395
   End
   Begin VB.TextBox txtResultado 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5625
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   30
      Width           =   7050
   End
End
Attribute VB_Name = "frmResultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()

Me.Caption = LoadResString(1019)

    Dim intFile As Integer

    intFile = FreeFile
    Dim linha As String
    Dim linhaant As String
    
    Open App.path & "\saida.txt" For Input As #intFile
    
    While Not EOF(intFile)
        Line Input #intFile, linha
        
        If linha <> "" Then
            If linhaant <> linha Then
                txtResultado.Text = txtResultado.Text & linha & vbNewLine
            End If
        End If
        linhaant = linha
    Wend
    
    Close
    
End Sub

Private Sub Form_Resize()
    txtResultado.Move 30, 30, Me.Width - 200, Me.Height - 1000
    cmdOK.Move Me.Width - cmdOK.Width - 120, Me.Height - cmdOK.Height - 400
End Sub
