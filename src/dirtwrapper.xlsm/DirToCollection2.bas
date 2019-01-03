Attribute VB_Name = "DirToCollection2"
Option Explicit

Sub test()
  Dim files As Collection
  Set files = getFilelistRecursively(ThisWorkbook.Path)

  Dim file As Variant
  For Each file In files
    Debug.Print file
  Next
End Sub

'BaseCollection��getFilelistRecursively���ŏ��������̂Œ���
'�Ԃ�l��BaseCollection�́A�S������Colelction���w���̂ŁA�ǂ�����g���Ă�OK
Function getFilelistRecursively(Path As String, Optional ByRef BaseCollection As Collection) As Collection
  If BaseCollection Is Nothing Then
    Set BaseCollection = New Collection
  End If
  
'Path�Ŏw�肳�ꂽ�t�H���_�z���̃t�H���_���擾
'�t�H���_������΍ċA����
  Dim Folders As Collection
  Set Folders = DirWrapper(Path, "*", vbDirectory)
  
  Dim Folder As Variant
  For Each Folder In Folders
    Call getFilelistRecursively(CStr(Folder), BaseCollection)
  Next
   
'Path�Ŏw�肳�ꂽ�t�H���_�z���̃t�@�C�����擾
  Call DirWrapper(Path, "*.xlsx", vbNormal, BaseCollection)
  
  Set getFilelistRecursively = BaseCollection
  
End Function

'BaseCollection��DirWrapper���ŏ��������̂Œ���
'�Ԃ�l��BaseCollection�́A�S������Colelction���w���̂ŁA�ǂ�����g���Ă�OK
Function DirWrapper(Path As String, Filter As String, Optional Attributes As VbFileAttribute = vbNormal, Optional ByRef BaseCollection As Collection) As Collection
  If BaseCollection Is Nothing Then
    Set BaseCollection = New Collection
  End If
  
  Dim Filename As String
  Filename = Dir(Path & "\" & Filter, Attributes)
    
  Do While Filename <> ""
  'GetAttr(Filename)��Attributes�̊Ԃ�and�́A�r�b�g���Z��And�ł��邱�Ƃɒ���
    If Attributes = vbNormal Or (GetAttr(Path & "\" & Filename) And Attributes) Then
      If Filename <> "." And Filename <> ".." Then
        BaseCollection.Add Path & "\" & Filename
      End If
    End If
    Filename = Dir()
  Loop

  Set DirWrapper = BaseCollection
End Function



