Attribute VB_Name = "modConst"
Option Explicit


Public Const icoOK As String = "ok"
Public Const icoNaoPrecisaCompilar As String = "nao_prec_compilar"
Public Const icoCompilando As String = "compilando"
Public Const icoerro As String = "erro"
Public Const icoNaoVerificado As String = "nao_verificado"
Public Const icoPrecisaCompilar As String = "compilavel"



Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const SW_SHOWNORMAL = 1
