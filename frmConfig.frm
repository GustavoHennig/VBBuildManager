VERSION 5.00
Begin VB.Form frmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuração"
   ClientHeight    =   3780
   ClientLeft      =   2925
   ClientTop       =   4455
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   6585
   Begin VB.CommandButton cmdCanc 
      Caption         =   "&Cancelar"
      Height          =   540
      Left            =   5265
      TabIndex        =   3
      Top             =   3135
      Width           =   1230
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Enabled         =   0   'False
      Height          =   540
      Left            =   3930
      TabIndex        =   2
      Top             =   3135
      Width           =   1230
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmConfig.frx":0000
      Left            =   90
      List            =   "frmConfig.frx":0013
      TabIndex        =   0
      Text            =   "Selecione Idioma"
      Top             =   390
      Width           =   2520
   End
   Begin VB.Label lblIdioma 
      Caption         =   "O idioma será detectado automaticamente."
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4005
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Ok As Boolean

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If Combo1.ListIndex >= 0 Then
       ' SaveSetting "CompiladorVB", "Config", "Idioma", Combo1.ItemData(Combo1.ListIndex)
       ' Idioma = Combo1.ItemData(Combo1.ListIndex)
        Ok = True
        Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    lblIdioma.Caption = LoadResString(1020)
    Me.Caption = LoadResString(1017)
    cmdCanc.Caption = LoadResString(1021)
End Sub
