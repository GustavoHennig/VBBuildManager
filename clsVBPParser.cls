VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVBPParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private FSO As New FileSystemObject

Public colArquivos As New Collection
Public Propriedades As New Collection
Public PathArqCompilado As String
Public DataUltCompilacao As Date
Public DataUltAquivoModificado As Date
Public NomeBinario As String
Public ErroLeitura As Boolean
Public UltimoErro As String
Public ExtensaoBinario As String

Public colProjetos As New Collection

Public Sub AbreProjeto(projeto As String)

    Dim intFile As Integer
    intFile = FreeFile
    Dim var As Variant
    Dim strLinha As String
    Dim path As String
    Dim path32 As String
    Dim arq_binario As String
    
    path = retPath(projeto)
    
 
    
    Open projeto For Input As #intFile
    
    While Not EOF(intFile)
        Line Input #intFile, strLinha
    
        If Left$(strLinha, 6) = "Class=" Or _
           Left$(strLinha, 7) = "Module=" Then
            
            var = Split(strLinha, ";")
            colArquivos.Add fl_CorrigePath(path, Trim$(var(UBound(var))), path)
            
        ElseIf Left$(strLinha, 5) = "Form=" Then
            colArquivos.Add path & Right$(strLinha, Len(strLinha) - 5)
        ElseIf Left$(strLinha, 12) = "UserControl=" Then
            colArquivos.Add fl_CorrigePath(path, Trim$(Right$(strLinha, Len(strLinha) - 12)), path)
        ElseIf Left$(strLinha, 10) = "ExeName32=" Then
            arq_binario = fl_RemoveAspas(Right$(strLinha, Len(strLinha) - 10))
        ElseIf Left$(strLinha, 7) = "Path32=" Then
            path32 = fl_RemoveAspas(Right$(strLinha, Len(strLinha) - 7))
        End If
        
        PopulaPropriedades strLinha
    Wend

    Close
    
    NomeBinario = arq_binario
    
    If Trim$(NomeBinario) = "" Then
        AdicionaErroSaida "O projeto " & Dir(projeto) & " precisa ser compilado pela IDE do Visual Basic na sua primeira compilação."
        ErroLeitura = True
    Else
        PathArqCompilado = fl_CorrigePath(path32, arq_binario, path)
        ExtensaoBinario = FSO.GetExtensionName(PathArqCompilado)
    End If
    
    sl_ValidaData

    
End Sub

Public Function retPath(projeto) As String
    retPath = Left$(projeto, Len(projeto) - Len(Dir(projeto)))
End Function

Public Function fl_CorrigePath(caminho As String, arq As String, path As String) As String
    
    On Error GoTo erro
    
    Dim curPath As String
    
    curPath = CurDir$
    
    ChDrive Left(path, 2)
    ChDir path
    
    fl_CorrigePath = FSO.GetAbsolutePathName(IIf(Left$(arq, 1) = "\", arq, FSO.BuildPath(caminho, arq)))
    
    
    ChDrive Left(curPath, 2)
    ChDir curPath
    
Exit Function
erro:
    ChDrive Left(curPath, 2)
    ChDir curPath

    MsgBox Err.Description
    Exit Function
    Resume
End Function


Private Sub sl_ValidaData()
    
    Dim L As Long
    Dim F As File
    Dim DataBinario As Date
    
    On Error GoTo erro
    
    DataUltCompilacao = CDate(0)
    DataUltAquivoModificado = CDate(0)

    If FSO.FileExists(PathArqCompilado) Then
        DataUltCompilacao = FSO.GetFile(PathArqCompilado).DateLastModified
    Else
        AdicionaErroSaida "Arquivo " & IIf(PathArqCompilado = "", "Nenhum", PathArqCompilado) & " não existe."
        ErroLeitura = True
    End If
    Dim bolCompilar As Boolean

    For L = 1 To colArquivos.Count
        Set F = FSO.GetFile(colArquivos(L))
        
        If DataUltAquivoModificado < F.DateLastModified Then
            DataUltAquivoModificado = F.DateLastModified
        End If
        
    Next
    
    Set F = Nothing
Exit Sub
erro:
    ErroLeitura = True
    AdicionaErroSaida Err.Description & " Arq: " & colArquivos(L)
    Resume Next
End Sub

Private Sub Class_Initialize()
    DataUltCompilacao = CDate(0)
    DataUltAquivoModificado = CDate(0)
End Sub


Public Sub PopulaPropriedades(linha As String)
    
    Dim var As Variant
    Dim pro As New clsPropVBP
    
    If linha <> "" Then
        If InStr(1, linha, "=") > 0 Then
            var = Split(linha, "=")
            
            pro.Propriedade = Trim$(var(0))
            pro.Valor = Trim$(var(1))
            Propriedades.Add pro
        End If
    End If
End Sub

Private Function fl_RemoveAspas(Texto As String) As String
    fl_RemoveAspas = Replace(Texto, """", "")
End Function


Public Property Get PrecisaRecompilar() As Boolean
    PrecisaRecompilar = DataUltAquivoModificado > DataUltCompilacao
End Property

Public Sub AdicionaErroSaida(erro As String)


    Dim intF As Integer
    intF = FreeFile
    
    UltimoErro = erro
    
    Open App.path & "\saida.txt" For Append As #intF
    
    Print #intF, erro
    
    Close #intF
    
End Sub

Public Sub CarregaColProjetos(ByVal path As String)

    Dim pastaPai As Folder
    Dim pasta As Folder
    
    Set pastaPai = FSO.GetFolder(path)

    fl_procuraEmPastaSimples pastaPai
    
    For Each pasta In pastaPai.SubFolders
        CarregaColProjetos pasta.path
    Next

End Sub

Private Function fl_procuraEmPastaSimples(ByVal pasta As Folder)

    Dim arq As File
    
    For Each arq In pasta.Files
        If LCase(FSO.GetExtensionName(arq.Name)) = "vbp" Then
            colProjetos.Add arq.path
        End If
    Next
    
End Function
