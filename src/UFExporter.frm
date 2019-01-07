VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UFExporter 
   Caption         =   "Export Data Range"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   OleObjectBlob   =   "UFExporter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UFExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' # ------------------------------------------------------------------------------
' # Name:        UFExporter.frm
' # Purpose:     Core UserForm for the CSV Exporter Excel VBA Add-In
' #
' # Author:      Brian Skinn
' #                bskinn@alum.mit.edu
' #
' # Created:     24 Jan 2016
' # Copyright:   (c) Brian Skinn 2016-2019
' # License:     The MIT License; see "LICENSE.txt" for full license terms.
' #
' #       http://www.github.com/bskinn/excel-csvexporter
' #
' # ------------------------------------------------------------------------------

Option Explicit

' ===== EVENT-ENABLED APPLICATION =====
Private WithEvents appn As Application
Attribute appn.VB_VarHelpID = -1


' =====  CONSTANTS  =====
Const NoFolderStr As String = "<none>"
Const InvalidSelStr As String = "<invalid selection>"


' =====  GLOBALS  =====
Dim WorkFolder As Folder
Dim fs As FileSystemObject
Dim ExportRange As Range


' =====  EVENT-ENABLED APPLICATION EVENTS  =====

Private Sub appn_SheetActivate(ByVal Sh As Object)
    ' Update the export range object, the
    ' export range reporting text, and the
    ' status of the 'Export' button any time a sheet
    ' is switched to
    setExportRange
    setExportRangeText
    setExportEnabled
End Sub

Private Sub appn_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    ' Update the export range object, the
    ' export range reporting text, and the
    ' status of the 'Export' button any time
    ' a new cell selection is made
    setExportRange
    setExportRangeText
    setExportEnabled
End Sub


' =====  FORM EVENTS  =====

Private Sub BtnClose_Click()
    ' Set the startup-position setting to 'Manual', so that the form
    '  will re-open where the user last placed it instead of in the
    '  center of the Excel window
    Me.StartUpPosition = 0  ' vbStartUpManual
    
    ' Hide the form without Unloading
    UFExporter.Hide
    
End Sub

Private Sub BtnExport_Click()
    
    Dim filePath As String, tStrm As TextStream, mode As IOMode
    
    ' Should only ever be possible to click if form is in a good state for exporting
    
    ' Proofread the range -- only one area allowed
    'If Selection.Areas.Count <> 1 Then
    '    Call MsgBox( _
    '        "Export of multiple areas is not supported!", _
    '        vbExclamation + vbOKOnly, _
    '        "Invalid Selection" _
    '    )
    '
    '    Exit Sub
    'End If
    
    ' Reject if entire column or row is selected
    'If ( _
    '    Selection.Rows.Count = Rows.Count _
    '    Or Selection.Columns.Count = Columns.Count _
    ') Then
    '    Call MsgBox( _
    '        "Cannot output entire rows or columns!", _
    '        vbExclamation + vbOKOnly, _
    '        "Invalid Selection" _
    '    )
    '
    '    Exit Sub
    'End If
    
    ' Store full file path
    filePath = fs.BuildPath(WorkFolder.Path, TxBxFilename.Value)
    
    ' Convert append setting to IOMode
    If ChBxAppend.Value Then
        mode = ForAppending
    Else
        mode = ForWriting
    End If
    
    ' Bind the text stream
    Set tStrm = fs.OpenTextFile(filePath, mode, True, TristateUseDefault)
    
    ' Ready to go. Pass info to writing function
    writeCSV ExportRange, tStrm, TxBxFormat.Value, TxBxSep.Value
    
    ' Close the stream
    tStrm.Close
    
End Sub

Private Sub BtnSelectFolder_Click()

    Dim fd As FileDialog
    Dim result As Long
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .Title = "Choose Output Folder"
        If InStr(UCase(.InitialFileName), "SYSTEM32") Then
            .InitialFileName = Environ("USERPROFILE") & "\Documents"
        End If
        
        result = .Show
    End With
    
    ' Drop if box cancelled
    If result = 0 Then Exit Sub
    
    ' Made it here; update the linked folder and the display textbox
    Set WorkFolder = fs.GetFolder(fd.SelectedItems(1))
    TxBxFolder.Value = WorkFolder.Path
    
    ' Update the Export button
    setExportEnabled

End Sub

Private Sub TxBxFilename_Change()

    ' If filename is nonzero-length and valid, enable Export and set color black
    If validFilename(TxBxFilename.Value) Then
        TxBxFilename.ForeColor = RGB(0, 0, 0)
    Else
        TxBxFilename.ForeColor = RGB(255, 0, 0)
    End If
    
    setExportEnabled
    
End Sub

Private Sub TxBxFormat_Change()
    setExportEnabled
End Sub

Private Sub TxBxSep_Change()
    setExportEnabled
End Sub

Private Sub UserForm_Activate()
    ' Always update the export range info box when
    ' focus is gained
    setExportRange
    setExportRangeText
    
End Sub

Private Sub UserForm_Initialize()
    ' Set to no folder selected
    TxBxFolder.Value = NoFolderStr
    
    ' Link filesystem
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Link Application for events
    Set appn = Application
    
    ' Default is for filename to be empty; disable export button
    BtnExport.Enabled = False
    
    ' Comma is default separator
    TxBxSep.Value = ","
    
    ' General is default number format
    TxBxFormat.Value = "@"
    
End Sub


' =====  FORM MANAGEMENT ROUTINES  =====

Private Sub setExportEnabled()

    If ( _
        Len(TxBxSep.Value) > 0 And _
        validFilename(TxBxFilename.Value) And _
        Len(TxBxFormat.Value) > 0 And _
        (Not WorkFolder Is Nothing) And _
        (Not ExportRange Is Nothing) _
    ) Then
        BtnExport.Enabled = True
    Else
        BtnExport.Enabled = False
    End If
    
End Sub


' =====  FUNCTIONS  =====

Private Sub writeCSV(dataRg As Range, tStrm As TextStream, nFormat As String, _
        Separator As String)
    
    Dim cel As Range
    Dim idxRow As Long, idxCol As Long
    Dim workStr As String
    
    ' Assume suitable TextStream already opened and dataRg proofed to only
    '  contain one Area.
    
    ' Loop
    For idxRow = 1 To dataRg.Rows.Count
        ' Reset the working string
        workStr = ""
        
        For idxCol = 1 To dataRg.Columns.Count
            ' Tag on the value and a separator
            workStr = workStr & Format(dataRg.Cells(idxRow, idxCol).Value, nFormat)
            workStr = workStr & Separator
        Next idxCol
        
        ' Cull the trailing separator
        workStr = Left(workStr, Len(workStr) - Len(Separator))
        
        ' Write the line
        tStrm.WriteLine workStr
        
    Next idxRow
    
End Sub

Function validFilename(fName As String) As Boolean
    
    Dim rxChrs As New RegExp
    
    With rxChrs
        .Global = True
        .IgnoreCase = True
        .MultiLine = False
        .Pattern = "[\\/:*?""<>|]"
        
        validFilename = (Len(fName) >= 1 And (Not .Test(fName)))
    End With
    
End Function

Private Sub setExportRangeText()
    ' Helper to set the export range info text
    
    Dim workStr As String
    
    If Not TypeOf Selection Is Range Then Exit Sub
    
    workStr = "  Worksheet: " _
        & Selection.Parent.Name _
        & Chr(10) _
        & "  Range: " _
        & getExportRangeAddress
    
    LblExportRg.Caption = workStr
    
End Sub

Private Sub setExportRange()
    ' Proofing of Selection, to see if it's valid -- plus,
    ' implementing the reduction of the export range to
    ' Intersect(UsedRange, Selection) when whole rows/columns
    ' are selected.
    
    If Selection.Areas.Count <> 1 Then
        Set ExportRange = Nothing
    Else
        If Selection.Address = Selection.EntireRow.Address Or _
                Selection.Address = Selection.EntireColumn.Address Then
            Set ExportRange = Intersect(Selection, Selection.Parent.UsedRange)
        Else
            Set ExportRange = Selection
        End If
    End If
    
End Sub

Private Function getExportRangeAddress() As String
    ' Helper to generate the export range address without dollar signs
    
    If ExportRange Is Nothing Then
        getExportRangeAddress = InvalidSelStr
    Else
        getExportRangeAddress = ExportRange.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    End If
    
End Function
