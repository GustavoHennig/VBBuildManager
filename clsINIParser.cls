VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsINIParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public ArquivoINI As String

Public Function Le(Secao As String, Key As String, Optional Default As String = "", Optional mArquivoIni As String) As String
    Dim strRet As String
    Dim varString As Variant
    Dim strAux As String
    Dim intVal As Long
    
    varString = ""
    
    strRet = Space(256)
    strAux = Key
    intVal = GetPrivateProfileString(Secao, strAux, Default, strRet, 256, IIf(mArquivoIni = "", ArquivoINI, mArquivoIni))
    varString = varString & Left$(strRet, intVal)
    
    Le = varString
    'fl_ObtemValorINI = Left$(strRet, intVal)
    
End Function

Public Sub Grava(strLocal As String, Chave As String, ByVal Dado As String, Optional mArquivoIni As String)
    WritePrivateProfileString strLocal, Chave, Dado, IIf(mArquivoIni = "", ArquivoINI, mArquivoIni)
End Sub
