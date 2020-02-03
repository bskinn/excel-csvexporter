Attribute VB_Name = "Exporter"

' # ------------------------------------------------------------------------------
' # Name:        Exporter.bas
' # Purpose:     Helper module for launching the CSV Exporter add-in
' #
' # Author:      Brian Skinn
' #                bskinn@alum.mit.edu
' #
' # Created:     24 Jan 2016
' # Copyright:   (c) Brian Skinn 2016-2020
' # License:     The MIT License; see "LICENSE.txt" for full license terms.
' #
' #       http://www.github.com/bskinn/excel-csvexporter
' #
' # ------------------------------------------------------------------------------

Option Explicit

Sub showForm()
Attribute showForm.VB_Description = "Load the CSVExporter application."
Attribute showForm.VB_ProcData.VB_Invoke_Func = "C\n14"
    UFExporter.Show
End Sub

