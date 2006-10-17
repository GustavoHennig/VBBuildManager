VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmCompilador 
   Caption         =   "Conpusis Compilador"
   ClientHeight    =   7845
   ClientLeft      =   2970
   ClientTop       =   2070
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   7845
   ScaleWidth      =   7815
   Begin VB.CommandButton cmdExecuta 
      Caption         =   "Command1"
      Height          =   570
      Left            =   4335
      TabIndex        =   14
      Top             =   6945
      Width           =   2235
   End
   Begin VB.Frame fraBibliotecas 
      Caption         =   "Bibliotecas"
      Height          =   1935
      Left            =   180
      TabIndex        =   10
      Top             =   675
      Width           =   5985
      Begin VB.CommandButton cmdRemLib 
         Caption         =   "Del"
         Height          =   480
         Left            =   4995
         TabIndex        =   13
         Top             =   1155
         Width           =   825
      End
      Begin VB.CommandButton cmdAddBib 
         Caption         =   "Add"
         Height          =   465
         Left            =   4980
         TabIndex        =   12
         Top             =   480
         Width           =   870
      End
      Begin VB.ListBox lstLib 
         Height          =   1425
         ItemData        =   "frmCompilador.frx":0000
         Left            =   195
         List            =   "frmCompilador.frx":0007
         TabIndex        =   11
         Top             =   285
         Width           =   4725
      End
   End
   Begin VB.Frame fraArquivosFonte 
      Caption         =   "ArquivosFonte"
      Height          =   1935
      Left            =   180
      TabIndex        =   6
      Top             =   2670
      Width           =   5985
      Begin VB.CommandButton cmdDelFontes 
         Caption         =   "Del"
         Height          =   480
         Left            =   4965
         TabIndex        =   9
         Top             =   960
         Width           =   825
      End
      Begin VB.CommandButton cmdAddFontes 
         Caption         =   "..."
         Height          =   360
         Left            =   4905
         TabIndex        =   8
         Top             =   450
         Width           =   870
      End
      Begin VB.ListBox lstFontes 
         Height          =   1035
         ItemData        =   "frmCompilador.frx":0013
         Left            =   150
         List            =   "frmCompilador.frx":001A
         TabIndex        =   7
         Top             =   345
         Width           =   4725
      End
   End
   Begin VB.Frame fraResource 
      Caption         =   "Resources"
      Height          =   1935
      Left            =   180
      TabIndex        =   2
      Top             =   4665
      Width           =   5985
      Begin VB.CommandButton cmdDelResources 
         Caption         =   "Del"
         Height          =   480
         Left            =   4890
         TabIndex        =   5
         Top             =   1110
         Width           =   825
      End
      Begin VB.CommandButton cmdAddRes 
         Caption         =   "..."
         Height          =   465
         Left            =   4875
         TabIndex        =   4
         Top             =   390
         Width           =   870
      End
      Begin VB.ListBox lstResources 
         Height          =   1425
         ItemData        =   "frmCompilador.frx":0026
         Left            =   105
         List            =   "frmCompilador.frx":002D
         TabIndex        =   3
         Top             =   315
         Width           =   4665
      End
   End
   Begin VB.CommandButton cmdProcArq 
      Caption         =   "Procurar"
      Height          =   465
      Left            =   4530
      TabIndex        =   1
      Top             =   105
      Width           =   1005
   End
   Begin VB.TextBox txtCaminho 
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   4305
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6255
      Top             =   15
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCompilador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

