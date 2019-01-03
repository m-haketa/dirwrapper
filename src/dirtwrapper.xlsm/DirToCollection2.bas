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

'BaseCollectionはgetFilelistRecursively内で書き換わるので注意
'返り値とBaseCollectionは、全く同じColelctionを指すので、どちらを使ってもOK
Function getFilelistRecursively(Path As String, Optional ByRef BaseCollection As Collection) As Collection
  If BaseCollection Is Nothing Then
    Set BaseCollection = New Collection
  End If
  
'Pathで指定されたフォルダ配下のフォルダを取得
'フォルダがあれば再帰処理
  Dim Folders As Collection
  Set Folders = DirWrapper(Path, "*", vbDirectory)
  
  Dim Folder As Variant
  For Each Folder In Folders
    Call getFilelistRecursively(CStr(Folder), BaseCollection)
  Next
   
'Pathで指定されたフォルダ配下のファイルを取得
  Call DirWrapper(Path, "*.xlsx", vbNormal, BaseCollection)
  
  Set getFilelistRecursively = BaseCollection
  
End Function

'BaseCollectionはDirWrapper内で書き換わるので注意
'返り値とBaseCollectionは、全く同じColelctionを指すので、どちらを使ってもOK
Function DirWrapper(Path As String, Filter As String, Optional Attributes As VbFileAttribute = vbNormal, Optional ByRef BaseCollection As Collection) As Collection
  If BaseCollection Is Nothing Then
    Set BaseCollection = New Collection
  End If
  
  Dim Filename As String
  Filename = Dir(Path & "\" & Filter, Attributes)
    
  Do While Filename <> ""
  'GetAttr(Filename)とAttributesの間のandは、ビット演算のAndであることに注意
    If Attributes = vbNormal Or (GetAttr(Path & "\" & Filename) And Attributes) Then
      If Filename <> "." And Filename <> ".." Then
        BaseCollection.Add Path & "\" & Filename
      End If
    End If
    Filename = Dir()
  Loop

  Set DirWrapper = BaseCollection
End Function



