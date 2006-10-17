VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCarregaImagens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub sg_carrega(ByRef il As ImageList, ByVal disabled_list As Boolean)


'    Dim FSO As New FileSystemObject
'
'    Dim pasta As Folder
'    Dim arq As File
    Dim is_dis As Integer
    
    If disabled_list Then
        is_dis = 200
    Else
        is_dis = 100
    End If
    
'    Set pasta = FSO.GetFolder(App.path & "\" & pasta_arq)
'
'    Dim strArq As String
    
    
    Dim i As Integer
    For i = 1 To 12
        il.ListImages.Add i, LoadResString(i), LoadResPicture(i + is_dis, vbResBitmap)
    Next
'
'    For Each arq In pasta.Files
'
'
'
'        If LCase(FSO.GetExtensionName(arq.Name)) = "bmp" Then
'          '  il.ListImages.Add , arq.Name, LoadPicture(arq.path)
'        End If
'    Next
End Sub

Public Sub sg_CarregaImgToolbar(ByRef tlb As Toolbar, ByRef il As ImageList)

    Dim i As Integer
    'On Error Resume Next
    
    Set tlb.ImageList = il
    'Set tlb.DisabledImageList = il
    Dim b As Button

    For Each b In tlb.Buttons
        If b.Style = tbrDefault Then
            b.Image = b.tag
            
        End If
    Next

End Sub
