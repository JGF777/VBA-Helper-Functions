Attribute VB_Name = "GITHUB"
'*************************************************************************************************************
'
'The following module contains two key parts:
'
'   1 - The following functions/sub routines are implemented as methods of one class that acts as a "helper" for many
'       different projects that I've developed in my job. Their main purpose is to help in other modules
'       that focus on automating monthly/weekly financial reports.
'
'   2 - Userform sub routines that are responsible for making userform resizable or adding a maximize/
'        minimize button. Note that they require prior declaration of other Windows API functions and private
'       constants that are included as a commented block.
'
'DISCLAIMER: I've tried to make those functions as dynamic as possible, however they could
'be further adapted to other projects depending on each person's needs and the singularities of their
'projects.
'
'Developed by Jacobo G.F 2022
'
'**************************************************************************************************************

Public Function Visual_Mods(switch As Boolean)

'Function that controls the main visualization indicators
'It takes one boolean argument

With Application
    .ScreenUpdating = switch
    .DisplayAlerts = switch
End With

End Function

Public Function OpenWb(NameFile As String) As Workbook

'Function that opens a wb through the Microsoft dialogue window
'accepts one string as argument

Dim wb As Workbook
Dim FileName As String
Dim fd As Office.FileDialog

Set fd = Application.FileDialog(msoFileDialogFilePicker)

With fd

    .AllowMultiSelect = False
    .Title = "Select " & NameFile
    .Show

End With

On Error Resume Next

FileName = fd.SelectedItems(1)

On Error GoTo 0

Set wb = Workbooks.Open(FileName)
Set OpenWb = wb

End Function

Public Function Edit_Links(wbk1 As Workbook, wbk2 As Workbook)

'Function that edits workbooks' links to change the references accordingly.
'The purpose is to referece itself, mainly used for periodical update of differente templates
'Takes 2 neccesary workbooks as arguments

Dim filelink As String
Dim NameFile As String
Dim path As String

filelink = wbk1.FullName '(ex: c:\\myDocuments\workbook.xlsx)
NameFile = wbk2.Name '(ex: workbook.xlsx)
path = wbk1.path '(c:\\myDocuments)

ChDir path
wbk2.ChangeLink Name:= _
filelink, NewName:= _
NameFile, Type:=xlExcelLinks

End Function

Public Function Copy_Sheet(WsName As String, WbCopy As Workbook, WbPaste As Workbook)

'Simple function that copies one sheet from the previous template to the current one taking the first position
'as index on the workbook.
'Needs 3 arguments:
'-WsName to indicate which one must be copied
'-WbCopy as the wb source of the data
'-WbPaste as the wb destination of the data

With WbCopy.Sheets(WsName)
    .Activate
    .Select
    .Copy Before:=WbPaste.Sheets(1)
End With

End Function

Public Function ChangeSource(WsName As String, PTName As String, Optional WsofPtName As String)

    'Function that updates the different Databases and their corresponding PTs
    'It takes 2 arguments as strings, the name of the DB's worksheet and the name of its linked PT
    
    Dim pt As PivotTable
    Dim MyData As Range
    
    Set MyData = ActiveWorkbook.Worksheets(WsName).Range("A1").CurrentRegion
    
    For Each pt In ActiveWorkbook.Worksheets(WsofPtName).PivotTables
             pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create _
                (SourceType:=xlDatabase, SourceData:=MyData)
    Next pt
   
   'Final check to save data
    ActiveWorkbook.Worksheets(WsofPtName).PivotTables(1).SaveData = True
        
End Function

Public Function Copy_Table(wbk1 As Workbook, wbk2 As Workbook, WsTemplate As String, wsSource As String)
  
    'Function that copies different pivot tables for monthly reports, it s able to keep the source format.
    'It takes 4 arguments as strings, being two of them the name of the DB's worksheet and name of template's
    'worksheet
    
    wbk2.Worksheets(WsTemplate).Range("A5").CurrentRegion.Clear
    
    wbk1.Worksheets(wsSource).Range("A4").CurrentRegion.Copy
    With wbk2.Worksheets(WsTemplate).Range("A5")
        .PasteSpecial Paste:=xlPasteFormats
        .PasteSpecial Paste:=xlPasteValues
    End With
        
End Function

Public Function Filter_Pt_Bytype(PtField As String, coll As Collection, PTName As String)

    'Function that filters a PT.
    'It takes 2 arguments, a collection with the items that we want to filter and the name
    'of the relevant PT.

    Dim pvtitem As PivotItem
    Dim i As Integer

    For Each pvtitem In ActiveSheet.PivotTables(PTName).PivotFields(PtField).PivotItems
        
        For i = 1 To coll.Count
            
            If coll(i) = pvtitem.Value Then
            
                ActiveSheet.PivotTables(PTName).PivotFields(PtField).PivotItems(pvtitem.Value).Visible = True
        
                Exit For
                
            ElseIf coll(i) <> pvtitem.Value Then
        
                ActiveSheet.PivotTables(PTName).PivotFields(PtField).PivotItems(pvtitem.Value).Visible = False
                
            End If
            
        Next i
        
    Next pvtitem

End Function

Public Function Get_absolutevalue(wbk As Workbook, WsTemplate As String, RangeAbsolute As String)

'Function that gets the absolute value of all the data inputs obtained from the GL for relevant reports such as the GTB ones.
'It takes 3 arguments:
'   -wbk as the workbook of the template
'   -WsTemplate as the worksheet with the raw copied data
'   -RangeAbsolute as string with the range we wanna loop through to convert to absolute values

wbk.Worksheets(WsTemplate).Activate

Dim range_Toloop As Range
Dim cell As Variant

Set range_Toloop = ActiveSheet.Range(RangeAbsolute)

For Each cell In range_Toloop

    If cell.Value <> "" Then
        cell.Value = cell.Value * -1
    End If

Next cell

End Function

Public Function BotWait(Seconds As String)

'Sleep function used in different scraping projects.

Application.Wait Now + TimeValue("00:00:" & Seconds)

End Function

Function GetFile(NameItem As String, FolderPath As String)

'This function is responsible of looping on a specific folder and find a specific workbook.
'It requires two arguments, the name/string of the item that we want to find and
'the folder path of the item

Dim FileItem As File
Dim fso As New FileSystemObject
Dim SourceFolder As Object
Dim wb As Workbook

Set SourceFolder = fso.GetFolder(FolderPath)

For Each FileItem In SourceFolder.Files
  
    If InStr(FileItem.Name, NameItem) <> 0 Then
    
        Set wb = Workbooks.Open(FolderPath & "\" & FileItem.Name)
        
        Set GetFile = wb
        
        Exit Function
        
    End If
        
Next FileItem

End Function
