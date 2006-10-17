VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmLegenda 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Legenda..."
   ClientHeight    =   4350
   ClientLeft      =   3285
   ClientTop       =   5400
   ClientWidth     =   6495
   Icon            =   "frmLegenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6495
   Begin RichTextLib.RichTextBox rtbLegenda 
      Height          =   4185
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   6330
      _ExtentX        =   11165
      _ExtentY        =   7382
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmLegenda.frx":000C
   End
End
Attribute VB_Name = "frmLegenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = LoadResString(1011)
    Dim rc As New clsRecursos
    rtbLegenda.TextRTF = rc.fg_getStringRecurso(cRTF_Legenda)
    Set rc = Nothing
    
End Sub



