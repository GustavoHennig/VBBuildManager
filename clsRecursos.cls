VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum enuTipoRecurso
    cRTF_Legenda = 101
    cRTF_Sobre = 102
End Enum

Public Function fg_getStringRecurso(ID As enuTipoRecurso) As String

    fg_getStringRecurso = f_ConvByteString(LoadResData(ID, "rtf"))

End Function

Private Function f_ConvByteString(ByVal b As Variant) As String


    Dim strR As String
    Dim i As Long
    
    For i = 0 To UBound(b)
        strR = strR & ChrW$(b(i))
    Next

    f_ConvByteString = strR

End Function
