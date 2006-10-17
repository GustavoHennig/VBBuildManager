VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCompilaVB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "Gerenciador de compilação VB"
   ClientHeight    =   6690
   ClientLeft      =   2400
   ClientTop       =   3090
   ClientWidth     =   8430
   HasDC           =   0   'False
   Icon            =   "frmCompilaVB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   8430
   Begin MSComctlLib.ImageList imglstDisabled 
      Left            =   6630
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ilLVW 
      Left            =   7815
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompilaVB.frx":08CA
            Key             =   "compilando"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompilaVB.frx":0C1C
            Key             =   "ok"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompilaVB.frx":0F6E
            Key             =   "nao_prec_compilar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompilaVB.frx":12C0
            Key             =   "nao_verificado"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompilaVB.frx":1612
            Key             =   "compilavel"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCompilaVB.frx":1964
            Key             =   "erro"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   6315
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9922
            MinWidth        =   7408
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgLstToolBar 
      Left            =   7215
      Top             =   150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "novo"
            Object.ToolTipText     =   "Novo"
            Object.Tag             =   "novo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "abrir"
            Object.ToolTipText     =   "Abrir"
            Object.Tag             =   "abrir"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar"
            Object.Tag             =   "salvar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "compilaatual"
            Object.ToolTipText     =   "Compila Selecionados"
            Object.Tag             =   "comp_sel"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "compilatudo"
            Object.ToolTipText     =   "Compila Tudo"
            Object.Tag             =   "comp_all"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sair"
            Object.ToolTipText     =   "Sair"
            Object.Tag             =   "sair"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar pbCompilacao 
      Height          =   345
      Left            =   4470
      TabIndex        =   1
      Top             =   300
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame fraProjetos 
      Caption         =   "Projetos"
      Height          =   5550
      Left            =   45
      TabIndex        =   0
      Top             =   630
      Width           =   8265
      Begin MSComctlLib.ListView lstProjetos 
         Height          =   5145
         Left            =   780
         TabIndex        =   4
         Top             =   225
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   9075
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ilLVW"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "nm_proj"
            Text            =   "Nome Projeto"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "tipo"
            Text            =   "Tipo"
            Object.Width           =   952
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "caminho"
            Text            =   "Caminho Completo"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "erro"
            Text            =   "Erro"
            Object.Width           =   7056
         EndProperty
      End
      Begin MSComctlLib.Toolbar ToolbarFilho 
         Height          =   2970
         Left            =   75
         TabIndex        =   5
         Top             =   240
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   5239
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "procurar"
               Object.ToolTipText     =   "Adicionar..."
               Object.Tag             =   "add"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "apagar"
               Object.ToolTipText     =   "Apagar"
               Object.Tag             =   "del"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
               Object.Width           =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "addallvbp"
               Object.ToolTipText     =   "Adicionar todos VBPs da pasta"
               Object.Tag             =   "find"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "refresh"
               Object.ToolTipText     =   "Refresh"
               Object.Tag             =   "refresh"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "sobe"
               Object.ToolTipText     =   "Mover para Cima"
               Object.Tag             =   "up"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "desce"
               Object.ToolTipText     =   "Mover para baixo"
               Object.Tag             =   "down"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label 
         Caption         =   "Loading..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Left            =   855
         TabIndex        =   6
         Top             =   345
         Width           =   4965
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3705
      Top             =   3885
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   2048
   End
   Begin VB.Menu mnarq 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnnovo 
         Caption         =   "&Novo"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnabrir 
         Caption         =   "A&brir"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnsalvar 
         Caption         =   "&Salvar"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnsep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnAcoes 
      Caption         =   "Açõ&es"
      Begin VB.Menu mncompsel 
         Caption         =   "Compila &Selecionados"
      End
      Begin VB.Menu mncomptudo 
         Caption         =   "Compila &Tudo"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnconf 
         Caption         =   "Configurações"
      End
   End
   Begin VB.Menu mnAjuda 
      Caption         =   "Ajuda"
      Begin VB.Menu mnHP 
         Caption         =   "Home Page"
      End
      Begin VB.Menu mnsep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnLegenda 
         Caption         =   "Legenda"
      End
      Begin VB.Menu mnsep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnsobre 
         Caption         =   "Sobre..."
      End
   End
   Begin VB.Menu mnpOp 
      Caption         =   "Oções"
      Visible         =   0   'False
      Begin VB.Menu mnop 
         Caption         =   " = Opções ="
         Enabled         =   0   'False
      End
      Begin VB.Menu mnpAbrirVB 
         Caption         =   "Abrir no Visual Basic"
      End
      Begin VB.Menu mnpSelTudo 
         Caption         =   "Selecionar Tudo"
      End
   End
End
Attribute VB_Name = "frmCompilaVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private INI As New clsINIParser
Private FSO As New FileSystemObject

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal Tempo As Long)
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Const STATUS_PENDING = &H103&
Const PROCESS_QUERY_INFORMATION = &H400

Private lngDifLstH As Long
Private lngDifFraH As Long
Private lngDifLstW As Long
Private lngDifFraW As Long

Private Titulo As String
Private PathVisualBasic As String

Private Sub sl_Abrir(Optional ByVal path As String)
    On Error GoTo erro
  
    If path = "" Then
        cd.Filter = LoadResString(1024) & " *.vbc|*.vbc"
        cd.CancelError = True
        cd.ShowOpen
        
        INI.ArquivoINI = cd.FileName
    Else
        INI.ArquivoINI = path
    
    End If
    
    Me.Caption = Titulo & " - " & fl_RetNomeProjeto(INI.ArquivoINI)
    

    Dim lngI As Long
    Dim cnt As Long
    cnt = getMaxIni
    
    Dim projeto As String

    lstProjetos.ListItems.Clear

    pbCompilacao.Max = cnt
    StatusBar.Panels(1).Text = LoadResString(1025)

    Dim strProj As String

    If FSO.FileExists(App.path & "\saida.txt") Then
        FSO.DeleteFile (App.path & "\saida.txt")
    End If

    lstProjetos.Visible = False

    For lngI = 1 To cnt
        strProj = INI.Le("Projetos", "prj" & Format$(lngI, "00"), "ERRO")
        sl_AdicionaProjeto strProj
        pbCompilacao.Value = lngI
        DoEvents
    Next
    lstProjetos.Visible = True


    pbCompilacao.Value = 0
    StatusBar.Panels(1).Text = ""
    
    If FSO.FileExists(App.path & "\saida.txt") Then
        frmResultado.Show vbModal
    End If
    
    
Exit Sub
erro:
    lstProjetos.Visible = True
    If Err.Number <> 32755 Then
        MsgBox Err.Description & vbNewLine & "Projeto:" & vbNewLine & strProj, vbCritical
    End If
End Sub

Private Sub sl_AddProj()

    On Error GoTo erro
    
    cd.Filter = "Projetos|*.vbp;*.vbg"
    cd.CancelError = True
    cd.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer
    
    cd.ShowOpen
    
  
    
    Dim var As Variant
    Dim lngI As Long
    Dim Nro As Long
    Dim path As String
    
    'Cria um vetor com os arquivos selecionados
    var = Split(cd.FileName, Chr$(0))

    'Conta o nro de arqv
    Nro = UBound(var)
    
    If Nro > 1 Then
        'O primeiro registro eh o path
        
        path = var(0)
      
        pbCompilacao.Max = Nro
        StatusBar.Panels(1).Text = "Analizando..."
      
        If FSO.FileExists(App.path & "\saida.txt") Then
            FSO.DeleteFile (App.path & "\saida.txt")
        End If
      
        For lngI = 1 To Nro
            pbCompilacao.Value = lngI
            DoEvents
            sl_AdicionaProjeto FSO.BuildPath(path, var(lngI))
        Next
        
        pbCompilacao.Value = 0
        StatusBar.Panels(1).Text = ""
        
        If FSO.FileExists(App.path & "\saida.txt") Then
            frmResultado.Show vbModal
        End If
    
    Else
        'Se soh tiver um ele ja vem com o path construido
        sl_AdicionaProjeto var(0)
    End If

    cd.Flags = 0
    cd.FileName = ""
Exit Sub
erro:
    cd.FileName = ""
    cd.Flags = 0
    
    If Err.Number <> 32755 Then
        MsgBox Err.Description, vbCritical
    End If
    Exit Sub
    Resume
End Sub

Private Sub sl_AdicionaProjeto(ByVal projeto As String)

    
    Dim cVBP As New clsVBPParser
    
    Dim strICOKey As String
    Dim strErro As String
    
    strICOKey = icoNaoVerificado
    
    If FSO.FileExists(projeto) Then
        cVBP.AbreProjeto projeto
        
        If cVBP.PrecisaRecompilar Then
            If cVBP.ErroLeitura Then
                strICOKey = icoNaoVerificado
                strErro = cVBP.UltimoErro
            Else
                strICOKey = icoPrecisaCompilar
            End If
        Else
            strICOKey = icoNaoPrecisaCompilar
        End If
        
    End If
    
    
    sl_AdicionaItemLVW projeto, strICOKey, strICOKey = icoPrecisaCompilar, cVBP.ExtensaoBinario, strErro
    DoEvents
End Sub

Private Function getMaxIni() As Long
    Dim ret As Long

    ret = CLng(INI.Le("DadosProjetos", "Count", "-1"))

    If ret = -1 Then
        INI.Grava "DadosProjetos", "Count", "0"
        ret = 0
    End If
    
    getMaxIni = ret
    

End Function

Private Sub sl_DelProj()
    
    If Not lstProjetos.SelectedItem Is Nothing Then
        If lstProjetos.SelectedItem.Index <> -1 Then
            lstProjetos.ListItems.Remove lstProjetos.SelectedItem.Index
'            'lstProjetos.SelectedItem.EnsureVisible
'            lstProjetos.SelectedItem.Selected = True
'            If lstProjetos.ListItems.Count >= lstProjetos.SelectedItem.Index Then
'                Set lstProjetos.SelectedItem = lstProjetos.ListItems(lstProjetos.SelectedItem.Index)
'            Else
'
'            End If
            
        End If
    End If
End Sub

Private Sub sl_CompilaTudo()

    Dim lngI As Long
    
    On Error GoTo erro
    
    If lstProjetos.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    sl_BloqueiaTela
    
    If FSO.FileExists(App.path & "\saida.txt") Then
        FSO.DeleteFile (App.path & "\saida.txt")
    End If
    
    pbCompilacao.Max = lstProjetos.ListItems.Count
    
    Dim DeuErro As Boolean
    
    For lngI = 1 To lstProjetos.ListItems.Count
        StatusBar.Panels(1).Text = LoadResString(1026) & " '" & fl_RetNomeProjeto(lstProjetos.ListItems.item(lngI).tag) & "' ..."
        DeuErro = fl_ExecutaCompilacao(lstProjetos.ListItems.item(lngI).tag, lngI) Or DeuErro
        pbCompilacao.Value = lngI
    Next

    If DeuErro Then
        frmResultado.Show vbModal
    End If
    
    pbCompilacao.Value = 0

    sl_DesBloqueiaTela
Exit Sub
erro:
 MsgBox Err.Description, vbCritical
 sl_DesBloqueiaTela
End Sub

Private Sub sl_CompilaSelecionado()

    On Error GoTo erro
    
    Dim lngI As Long
    
    If lstProjetos.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    sl_BloqueiaTela
    
    If FSO.FileExists(App.path & "\saida.txt") Then
        FSO.DeleteFile (App.path & "\saida.txt")
    End If
    
    Dim lngMax As Long
    lngMax = 0
    'Calcula o nro de selecionados
    For lngI = 1 To lstProjetos.ListItems.Count
        If lstProjetos.ListItems.item(lngI).checked Then
            lngMax = lngMax + 1
        End If
    
    Next
    
    Dim lngCC As Long
    lngCC = 0
    If lngMax = 0 Then
        sl_DesBloqueiaTela
        Exit Sub
    End If
    
    pbCompilacao.Max = lngMax
    
    Dim DeuErro As Boolean
    
    For lngI = 1 To lstProjetos.ListItems.Count
        If lstProjetos.ListItems.item(lngI).checked Then
            lngCC = lngCC + 1
            StatusBar.Panels(1).Text = LoadResString(1026) & " '" & fl_RetNomeProjeto(lstProjetos.ListItems.item(lngI).tag) & "' ..."
            DeuErro = fl_ExecutaCompilacao(lstProjetos.ListItems.item(lngI).tag, lngI) Or DeuErro
            pbCompilacao.Value = lngCC
        End If
    Next
    
    If DeuErro Then
        frmResultado.Show vbModal
    End If
    
    pbCompilacao.Value = 0
    sl_DesBloqueiaTela
Exit Sub
erro:
    
    If MsgBox(Err.Description & vbNewLine & LoadResString(1027), vbYesNo) = vbYes Then
        Resume
    End If
    sl_DesBloqueiaTela
End Sub

Private Function fl_ExecutaCompilacao(ByVal projeto As String, ByVal Index As Long) As Boolean
    
    Dim item As ListItem
    
    Set item = lstProjetos.ListItems.item(Index)
    item.EnsureVisible
    
    lstProjetos.ListItems.item(Index).SmallIcon = icoCompilando
    lstProjetos.Refresh
    sl_Espera Shell(PathVisualBasic & _
                    " /make " & _
                    """" & projeto & _
                    """ /out """ & App.path & "\saida.txt""", vbHide)
                    
'    MsgBox PathVisualBasic & _
'                    " /make " & _
'                    """" & projeto & _
'                    """ /out """ & App.path & "\saida.txt"""
                    
    item.SmallIcon = fl_AnalizaSaida(projeto)
    item.checked = Not (item.SmallIcon = modConst.icoOK)
    
    fl_ExecutaCompilacao = item.checked
    Set item = Nothing
    DoEvents
    lstProjetos.Refresh
End Function



Private Sub sl_Salvar()
    
    On Error GoTo erro
  
    If lstProjetos.ListItems.Count = 0 Then
        Exit Sub
    End If
  
    If INI.ArquivoINI = "" Then
       cd.Filter = LoadResString(1024) & " *.vbc|*.vbc"
       cd.CancelError = True
       cd.ShowSave
    
       INI.ArquivoINI = cd.FileName
    End If
    
   
    Dim lngI As Long
    
    If FSO.FileExists(INI.ArquivoINI) Then
        FSO.DeleteFile (INI.ArquivoINI)
    End If
    
    pbCompilacao.Max = lstProjetos.ListItems.Count
    
    INI.Grava "DadosProjetos", "Count", CStr(lstProjetos.ListItems.Count)
    
    For lngI = 1 To lstProjetos.ListItems.Count
        INI.Grava "Projetos", "prj" & Format$(lngI, "00"), lstProjetos.ListItems.item(lngI).tag
        pbCompilacao.Value = lngI
    Next

    Me.Caption = Titulo & " - " & fl_RetNomeProjeto(INI.ArquivoINI)
    
    pbCompilacao.Value = 0

Exit Sub
erro:
    If Err.Number <> 32755 Then
        MsgBox Err.Description, vbCritical
    End If
End Sub

Private Sub sl_Desce()

    
    If lstProjetos.SelectedItem Is Nothing Then
        Exit Sub
    End If

    If lstProjetos.SelectedItem.Index < lstProjetos.ListItems.Count Then
    
        sl_TrocaItens lstProjetos.SelectedItem, _
                        lstProjetos.ListItems(lstProjetos.SelectedItem.Index + 1)
                        
        lstProjetos.SelectedItem.Selected = False
        Set lstProjetos.SelectedItem = lstProjetos.ListItems(lstProjetos.SelectedItem.Index + 1)
    End If

End Sub

Private Sub sl_TrocaItens(ByRef item1 As ListItem, ByRef item2 As ListItem)

    Dim nome As String
    Dim caminhocompleto As String
    Dim erro As String
    Dim img As Variant
    Dim checked As Boolean
    Dim tag As String
    Dim tipo As String
    
    nome = item1.Text
    img = item1.SmallIcon
    caminhocompleto = item1.ListSubItems("caminho").Text
    erro = item1.ListSubItems("erro").Text
    checked = item1.checked
    tag = item1.tag
    tipo = item1.ListSubItems("tipo").Text
    
    item1.Text = item2.Text
    item1.SmallIcon = item2.SmallIcon
    item1.ListSubItems("caminho").Text = item2.ListSubItems("caminho").Text
    item1.ListSubItems("erro").Text = item2.ListSubItems("erro").Text
    item1.ListSubItems("tipo").Text = item2.ListSubItems("tipo").Text
    item1.checked = item2.checked
    item1.tag = item2.tag
    
    item2.Text = nome
    item2.SmallIcon = img
    item2.ListSubItems("caminho").Text = caminhocompleto
    item2.ListSubItems("erro").Text = erro
    item2.ListSubItems("tipo").Text = tipo
    item2.checked = checked
    item2.tag = tag
    
    
End Sub

Private Sub sl_Sobe()
    If lstProjetos.SelectedItem Is Nothing Then
        Exit Sub
    End If

    If lstProjetos.SelectedItem.Index > 1 Then
    
        sl_TrocaItens lstProjetos.SelectedItem, _
                        lstProjetos.ListItems(lstProjetos.SelectedItem.Index - 1)
                        
        lstProjetos.SelectedItem.Selected = False
        Set lstProjetos.SelectedItem = lstProjetos.ListItems(lstProjetos.SelectedItem.Index - 1)
    End If

End Sub

Private Sub Form_Load()
    SetParent pbCompilacao.hwnd, StatusBar.hwnd
    
    Me.Caption = LoadResString(1018)
    Titulo = Me.Caption

    lngDifFraH = Me.Height - fraProjetos.Height
    lngDifFraW = Me.Width - fraProjetos.Width
    lngDifLstH = Me.Height - lstProjetos.Height
    lngDifLstW = Me.Width - lstProjetos.Width
    

    
    Dim cIL As New clsCarregaImagens
    cIL.sg_carrega imgLstToolBar, False
    cIL.sg_carrega imglstDisabled, True
    cIL.sg_CarregaImgToolbar ToolbarFilho, imgLstToolBar
    Set Toolbar.DisabledImageList = imglstDisabled
    cIL.sg_CarregaImgToolbar Toolbar, imgLstToolBar
    
    PathVisualBasic = """" & Environ$("ProgramFiles") & "\Microsoft Visual Studio\VB98\vb6.exe"""
    
    Set cIL = Nothing
    
    If Command$ <> "" Then
        sl_Abrir Replace(Command$, """", "")
    End If
    
    Me.Move GetSetting("CompiladorVB", "Posicao", "Left", 1995), _
            GetSetting("CompiladorVB", "Posicao", "Top", 1905), _
            GetSetting("CompiladorVB", "Posicao", "Width", 8550), _
            GetSetting("CompiladorVB", "Posicao", "Height", 7500)

  '  Idioma = GetSetting("CompiladorVB", "Config", "Idioma", 2000)
    
    sl_MudaIdioma
    
    
End Sub

Private Sub Form_Resize()
    
    If Me.WindowState <> 1 Then
        On Error Resume Next
        fraProjetos.Width = Me.Width - lngDifFraW
        lstProjetos.Width = Me.Width - lngDifLstW
        fraProjetos.Height = Me.Height - lngDifFraH
        lstProjetos.Height = Me.Height - lngDifLstH
        'ToolbarFilho.Top = Me.Height - 2400
        'ToolbarFilho.Width = lstProjetos.Width
    '    pbCompilacao.Value = 100
        pbCompilacao.Move StatusBar.Panels(2).Left + 30, 45, StatusBar.Panels(2).Width - 60, StatusBar.Height - 60
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)

    If Me.WindowState = 0 Then
        SaveSetting "CompiladorVB", "Posicao", "Top", Me.Top
        SaveSetting "CompiladorVB", "Posicao", "Left", Me.Left
        SaveSetting "CompiladorVB", "Posicao", "Width", Me.Width
        SaveSetting "CompiladorVB", "Posicao", "Height", Me.Height
    End If
End Sub

Private Sub lstProjetos_DblClick()
    mnpAbrirVB_Click
End Sub

Private Sub lstProjetos_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngI As Long
    
    Dim bolTodos As Boolean
    
    If KeyCode = vbKeySpace Then
        bolTodos = f_VerificaTodosMarcados
        For lngI = 1 To lstProjetos.ListItems.Count
            If lngI <> lstProjetos.SelectedItem.Index Then
                If lstProjetos.ListItems.item(lngI).Selected Then
                    lstProjetos.ListItems.item(lngI).checked = Not bolTodos
                End If
            End If
        Next
    End If
End Sub

Private Sub lstProjetos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
    
        PopupMenu mnpOp, , x + lstProjetos.Left, y + lstProjetos.Top + 300
        
    End If
End Sub

Private Sub mnabrir_Click()
sl_Abrir
End Sub

Private Sub mncompsel_Click()
sl_CompilaSelecionado
End Sub

Private Sub mncomptudo_Click()
sl_CompilaTudo
End Sub

Private Sub mnconf_Click()
    frmConfig.Show vbModal
    sl_MudaIdioma
End Sub

Private Sub mnHP_Click()
    On Error Resume Next
    ShellExecute Me.hwnd, vbNullString, "http://gustavo.somee.com", vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub mnLegenda_Click()
    frmLegenda.Show vbModal
End Sub

Private Sub mnnovo_Click()
    sl_Novo
End Sub

Private Sub mnpAbrirVB_Click()
    If Not lstProjetos.SelectedItem Is Nothing Then
       
       Shell PathVisualBasic & " """ & lstProjetos.SelectedItem.tag & """", vbNormalFocus
       
    End If
End Sub

Private Sub mnpSelTudo_Click()
    Dim lngI As Long
    
    For lngI = 1 To lstProjetos.ListItems.Count
        lstProjetos.ListItems.item(lngI).Selected = True
    Next

End Sub

Private Sub mnSair_Click()
sl_Sair
End Sub

Private Sub mnsalvar_Click()
sl_Salvar
End Sub

Private Sub mnsobre_Click()
    frmAbout.Show vbModal
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    Case "novo"
        sl_Novo
    Case "abrir"
        sl_Abrir
    Case "salvar"
        sl_Salvar
    Case "compilatudo"
        sl_CompilaTudo
    Case "compilaatual"
        sl_CompilaSelecionado
    Case "sair"
        sl_Sair
    End Select

    StatusBar.Panels(1).Text = ""

End Sub


Private Sub sl_Novo()

    INI.ArquivoINI = ""
    Me.Caption = Titulo
    lstProjetos.ListItems.Clear

End Sub


Private Sub sl_Espera(ID As Long)
    
    Dim lProcessId As Long
    Dim lExitCode As Long
    
    
    lProcessId = OpenProcess(PROCESS_QUERY_INFORMATION, False, ID)

    Dim inti As Integer
    
    Call GetExitCodeProcess(lProcessId, lExitCode)
    
    Do While lExitCode = STATUS_PENDING
        DoEvents
        Sleep 50
        inti = inti + 1
        If inti > 20000 Then
            Err.Raise 99, "sl_Espera, 15 min", "Timeout"
            Exit Do
        End If
        Call GetExitCodeProcess(lProcessId, lExitCode)
    Loop
    
End Sub


Private Function fl_RetNomeProjeto(ByVal projeto As String) As String

    fl_RetNomeProjeto = Dir(projeto)
End Function

Private Sub ToolbarFilho_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "procurar"
        sl_AddProj
    Case "apagar"
        sl_DelProj
    Case "sobe"
        sl_Sobe
    Case "desce"
        sl_Desce
    Case "addallvbp"
        sl_AdicionPastaVBPs
    Case "refresh"
        sl_Refresh
    End Select

End Sub

Private Sub sl_AdicionaItemLVW(Texto As String, ByVal ico As String, ByVal checado As Boolean, ByVal tipo As String, Optional ByVal strErro As String = "")

    Dim item As ListItem
    
    Set item = lstProjetos.ListItems.Add(, , Dir(Texto), , ico)
    item.tag = Texto
    item.checked = checado
    
    item.ListSubItems.Add , "tipo", UCase(tipo)
    item.ListSubItems.Add , "caminho", Texto
    item.ListSubItems.Add , "erro", strErro
    
    Set item = Nothing

End Sub



Private Sub sl_Sair()
    Unload Me
    End
End Sub

Private Sub sl_BloqueiaTela()
    'Toolbar.Enabled = False
   ' ToolbarFilho.Enabled = False
    Dim b As Button
    For Each b In Toolbar.Buttons
        b.Enabled = False
    Next
    For Each b In ToolbarFilho.Buttons
        b.Enabled = False
    Next
    'Toolbar.Buttons("abrir").Enabled = False
    mnAcoes.Enabled = False
    mnarq.Enabled = False
    Set b = Nothing
End Sub

Private Sub sl_DesBloqueiaTela()
  '  Toolbar.Enabled = True
   ' ToolbarFilho.Enabled = True
    mnAcoes.Enabled = True
    mnarq.Enabled = True
        Dim b As Button
    For Each b In Toolbar.Buttons
        b.Enabled = True
    Next
    For Each b In ToolbarFilho.Buttons
        b.Enabled = True
    Next
    Set b = Nothing
End Sub


Private Function fl_AnalizaSaida(projeto As String) As String



    On Error GoTo erro

    Dim ret As String
    Dim Texto As String
    Dim strNM_Bin As String
    Dim cVBP As New clsVBPParser
    Dim tx As TextStream
    Dim lngPos As Long
    
    cVBP.AbreProjeto projeto
    strNM_Bin = cVBP.NomeBinario
        
    Set tx = FSO.OpenTextFile(App.path & "\saida.txt", ForReading)
    
    Texto = tx.ReadAll
    
    lngPos = InStr(1, Texto, strNM_Bin)
    
    If Mid$(Texto, lngPos + Len(strNM_Bin) + 2, 9) = "succeeded" Then
        ret = modConst.icoOK
    Else
        ret = modConst.icoerro
    End If
    
    tx.Close
    
    Set tx = Nothing
    Set cVBP = Nothing
    
    fl_AnalizaSaida = ret
Exit Function
erro:
    Set tx = Nothing
    Set cVBP = Nothing
    MsgBox Err.Description, vbCritical
End Function

Private Sub sl_AdicionPastaVBPs()

    Dim vbp As New clsVBPParser
    Dim pasta As String
    Dim clsBrw As New clsBrowseFolder
    
    pasta = clsBrw.fg_ProcuraPasta
    
    
    
    If pasta <> "" Then
    
        If FSO.FileExists(App.path & "\saida.txt") Then
            FSO.DeleteFile (App.path & "\saida.txt")
        End If
    
        StatusBar.Panels(1).Text = LoadResString(1022)
        vbp.CarregaColProjetos pasta
        
        StatusBar.Panels(1).Text = LoadResString(1023)
        If vbp.colProjetos.Count > 0 Then
        
            pbCompilacao.Max = vbp.colProjetos.Count - 1
        
        End If
        
            Dim lngL As Long
            For lngL = 1 To vbp.colProjetos.Count - 1
                sl_AdicionaProjeto vbp.colProjetos(lngL)
                pbCompilacao.Value = lngL
            Next
        
        
        pbCompilacao.Value = 0
        StatusBar.Panels(1).Text = ""
        
        If FSO.FileExists(App.path & "\saida.txt") Then
            frmResultado.Show vbModal
        End If
        
    End If

End Sub

Private Sub sl_MudaIdioma()
On Error GoTo erro
    mnabrir.Caption = LoadResString(2 + Idioma)
    mnnovo.Caption = LoadResString(1 + Idioma)
    mnsalvar.Caption = LoadResString(3 + Idioma)
    mnSair.Caption = LoadResString(4 + Idioma)
    mnarq.Caption = LoadResString(5 + Idioma)
    mncompsel.Caption = LoadResString(6 + Idioma)
    mncomptudo.Caption = LoadResString(7 + Idioma)
    mnAcoes.Caption = LoadResString(8 + Idioma)
    mnAjuda.Caption = LoadResString(9 + Idioma)
    mnsobre.Caption = LoadResString(10 + Idioma)
    mnLegenda.Caption = LoadResString(11 + Idioma)
    mnconf.Caption = LoadResString(1017)
    
    ToolbarFilho.Buttons("procurar").ToolTipText = LoadResString(12 + Idioma)
    ToolbarFilho.Buttons("apagar").ToolTipText = LoadResString(13 + Idioma)
    ToolbarFilho.Buttons("sobe").ToolTipText = LoadResString(15 + Idioma)
    ToolbarFilho.Buttons("desce").ToolTipText = LoadResString(16 + Idioma)
    ToolbarFilho.Buttons("addallvbp").ToolTipText = LoadResString(14 + Idioma)



    Toolbar.Buttons("novo").ToolTipText = LoadResString(1 + Idioma)
    Toolbar.Buttons("abrir").ToolTipText = LoadResString(2 + Idioma)
    Toolbar.Buttons("salvar").ToolTipText = LoadResString(3 + Idioma)
    Toolbar.Buttons("compilatudo").ToolTipText = LoadResString(7 + Idioma)
    Toolbar.Buttons("compilaatual").ToolTipText = LoadResString(6 + Idioma)
    Toolbar.Buttons("sair").ToolTipText = LoadResString(4 + Idioma)
    Exit Sub
erro:
    MsgBox Err.Description, vbCritical
End Sub

Private Function f_VerificaTodosMarcados() As Boolean


    Dim lngI As Long
    Dim bolEncotrouDesmarcado As Boolean
    
    f_VerificaTodosMarcados = False
    
    
    For lngI = 1 To lstProjetos.ListItems.Count
       ' If lngI <> lstProjetos.SelectedItem.Index Then
            If lstProjetos.ListItems.item(lngI).Selected Then
                If Not lstProjetos.ListItems.item(lngI).checked Then
                    bolEncotrouDesmarcado = True
                    Exit For
                End If
            End If
       ' End If
    Next
    
    f_VerificaTodosMarcados = Not bolEncotrouDesmarcado
    
End Function

Private Sub sl_Refresh()


    Dim col As New Collection
    Dim item As ListItem
    Dim v As Variant
    
    On Error GoTo erro
    
    
    For Each item In lstProjetos.ListItems
        col.Add item.ListSubItems("caminho").Text
    Next

    lstProjetos.ListItems.Clear
    lstProjetos.Visible = False

    For Each v In col
        sl_AdicionaProjeto v
    Next
    
    lstProjetos.Visible = True

Exit Sub
erro:
    lstProjetos.Visible = True
End Sub
