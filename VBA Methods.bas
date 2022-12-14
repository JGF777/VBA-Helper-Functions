Attribute VB_Name = "GITHUB"
'*******************************************************************************************************************************
'
'     - The following functions/sub routines are implemented as methods of one class that acts as a "helper" for many
'       different projects that I've developed in my job based on OOP proggramming logic. 
'       Their main purpose is to help in other modules that focus on automating financial reports/analysis
'
'DISCLAIMER: I've tried to make those functions as dynamic as possible, however they could
'be further adapted to other projects depending on each person's needs and the singularities of the aforementioned projects
'
'--------------------------------------
'Developed by Jacobo G.F 2022
'--------------------------------------
'
'*******************************************************************************************************************************

Public Function Visual_Mods(switch As Boolean)

    'Function that controls the main visualization indicators
    'It takes one boolean argument

    With Application
        .ScreenUpdating = switch
        .DisplayAlerts = switch
    End With

End Function

'---------------------------------------------------------------------------

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

'---------------------------------------------------------------------------  
        
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
        
'---------------------------------------------------------------------------

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
            
'---------------------------------------------------------------------------

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
           
'---------------------------------------------------------------------------

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
            
'---------------------------------------------------------------------------

Public Function Filter_Pt_Bytype(Field As String, coll As Collection, PTName As String)

    'Function that filters a PT.
    'It takes 2 arguments, a collection with the items that we want to filter and the name
    'of the relevant PT.

    Dim pvtitem As PivotItem
    Dim i As Integer

                With ActiveSheet.PivotTables(PTName).PivotFields(Field)
    
        For Each pvtitem In .PivotItems
            
            For i = 1 To coll.Count
                
                If coll(i) = pvtitem.Value Then
                
                    .PivotItems(pvtitem.Value).Visible = True
            
                    Exit For
                    
                ElseIf coll(i) <> pvtitem.Value Then
            
                    .PivotItems(pvtitem.Value).Visible = False
                    
                End If
                
            Next i
            
        Next pvtitem
        
    End With

End Function
    
'---------------------------------------------------------------------------

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

    For Each cell In rango_Toloop

        With cell

            If .Value <> "" Then
                .Value = Abs(.Value)
            End If

        End With

    Next cell

End Function
    
'---------------------------------------------------------------------------

Public Function BotWait(Seconds As String)

'Sleep function used in different scraping projects.

    Application.Wait Now + TimeValue("00:00:" & Seconds)

End Function
    
'---------------------------------------------------------------------------

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
               
'---------------------------------------------------------------------------                
                
Function ColumnLettersFromRange(rInput As Range) As String

'Gets the string of the column, it s used to pass the ouput for the filldown method
'It takes one argument. Desired range to obtain the columnstring
                    
    ColumnLettersFromRange = Split(rInput.Address, "$")(1)

End Function

'---------------------------------------------------------------------------                          
                
Function FindRowNum(Type As String)

'Function that finds the row number of the totals of the different expenses types
'Returns a number to pass to another sub
    
 Dim rng As Range
 Dim rownumber As Long

 Set rng = ActiveSheet.Columns("B:B").find(What:=Type, _
    LookIn:=xlValues, LookAt:=xlWhole, SearchOrder:=xlByRows, _
    SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
    
    rownumber = rng.Row
    
    FindRowNum = rownumber

End Function
                    
'---------------------------------------------------------------------------
                    
Public Function clear_other_bars(NameBar As String)
                        
'Funcion used in userframes.
'Function that is responsible for dynamically clearing the other comboboxes value whenever the user is selecting
'different comboboxes
'It takes one argument:
'   'Name of the actual bar that we want to control and clear the rest, as a string

Dim cont As Control
Dim cont2 As Control

For Each cont In MainUI.Controls

    If cont.Name = NameBar Then
        
            For Each cont2 In MainUI.Controls
            
                            If TypeName(cont2) = "ComboBox" And cont2.Name <> NameBar Then 'It must fulfil two conditions,
                                                                                            'so we make sure we are looping through
                    If cont2.Value <> 0 Then                                                'only the comboboxes whose value we want
                                                                                            'to clear.
                        cont2.Value = ""
                        
                    End If
    
                End If
                
           Next cont2
           
    End If
    
Next cont

End Function

'---------------------------------------------------------------------------            
            
Public Function populate_dropdown(arr As Variant, bar As Control, counter As Integer)

'Populates a dropdown menu by adding the relevant items from an array
'it takes three arguments:
'   -An array with the strings of the different procedures
'   -The name of each of the dropwdown menus passed as control
'   -A counter to limit the amount of items created in each bar, equal to the total amount of tasks per bar

Dim uB As Integer, lB As Integer, i As Integer

uB = UBound(arr)
lB = LBound(arr)

With bar

    If .ListCount < counter Then
    
        For i = lB To uB
    
            .AddItem arr(i)
        
        Next i
        
    End If

End With

End Function
 
'---------------------------------------------------------------------------        

Sub Sizable()

'Main procedure to make the form resizable, it is called in the form initialize event
            
'Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare PtrSafe Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Const GWL_STYLE As Long = (-16)        
            
    Dim hWndForm As Long
    Dim iStyle As Long

    'Get the userform's window handle
    If Val(Application.Version) < 9 Then
        hWndForm = FindWindow("ThunderXFrame", Me.Caption)  'XL97
    Else
        hWndForm = FindWindow("ThunderDFrame", Me.Caption)  'XL2000
    End If

    'Make the form resizable
    iStyle = GetWindowLong(hWndForm, GWL_STYLE)
    iStyle = iStyle Or WS_THICKFRAME
    SetWindowLong hWndForm, GWL_STYLE, iStyle
 
End Sub   
        
'---------------------------------------------------------------------------    
'
'The following two functions/sub routines are .txt related, particularly       
'about how to read/write files for specific values programmatically
'
'---------------------------------------------------------------------------        

Sub ChangeValue(NEWValue As String)

    'This subroutine updates the  value based on the given argument
    
    Dim OLDValue As String
    OLDValue = FindOldValue

    Open EXCPpath For Input As #1
    c0 = Input(LOF(1), #1)
    Close #1

    Open EXCPpath For Output As #1
    Print #1, Replace(c0, OLDValue, NEWValue)
    Close #1
    
End Sub
        
'---------------------------------------------------------------------------            
 
Function  FindOldValue() As String

    'Finds the relevant line containing the old num value by filtering the "6.." 
    'It returns the old CNUM as string to pass to the main change value.
            
Dim Jno As String
Dim strLine As String
Dim JnoToCheck As String
Dim CleanValue As String

Open EXCPpath For Input As #1
     
    Jno = "<input value=""6"

Do While Not EOF(1)
 
        IngLine = IngLine + 1
        
        Line Input #1, strLine
        
        If InStr(1, strLine, Jno, vbTextCompare) > 1 Then
    
            JnoToCheck = Trim(Split(strLine, "input value=""")(1))      'The logic can be easily adapted for the Jno and JnotoCheck variables. 
                                                                        'depending on the variables that we want to find.
                                                                        'Or the relevant data manipulations that we want to make with it. 
            
            CleanValue = Left(JnoToCheck, 6)
            
            FindOldValue = CleanValue
            
            Close #1
            
            Exit Function

        End If
            
Loop
        
End Function      
            
