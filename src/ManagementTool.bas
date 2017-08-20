Attribute VB_Name = "ManagementTool"
Option Explicit

Private Enum vbext_ComponentType 'VBIDE
    vbext_ct_StdModule = 1
    vbext_ct_ClassModule
    vbext_ct_MSForm
End Enum

Sub ExportModule()
    
    Dim fso As Object 'As Scripting.FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    With fso
        Dim bookPath
        bookPath = ThisWorkbook.FullName
        
        Dim rootDir
        rootDir = .GetParentFolderName(.GetParentFolderName(bookPath))
        
        Dim exportDir
        exportDir = .BuildPath(rootDir, "src")
        
    End With 'FSO
    
    Dim targetBook As Excel.Workbook
    Set targetBook = Excel.ThisWorkbook
    
    Dim extensionDic  As Object 'As Scripting.Dictionary
    Set extensionDic = CreateObject("Scripting.Dictionary")
    With extensionDic
        .Item(vbext_ct_ClassModule) = ".cls"
        .Item(vbext_ct_MSForm) = ".frm"
        .Item(vbext_ct_StdModule) = ".bas"
    End With
    
    Dim tmpModule As Object 'As VBIDE.VBComponent
    For Each tmpModule In targetBook.VBProject.VBComponents
        If extensionDic.Exists(tmpModule.Type) Then
            Call tmpModule.Export(fso.BuildPath(exportDir, tmpModule.Name & extensionDic.Item(tmpModule.Type)))
        End If
    Next tmpModule
End Sub
