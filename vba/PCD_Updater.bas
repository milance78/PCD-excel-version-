Attribute VB_Name = "PCD_Updater"
Option Explicit

' This module is the Excel-side updater.
' It is kept in source control now and will be embedded in the first
' project-specific VBA build. The workbook never depends on GitHub at runtime
' unless the user explicitly checks for an update.

Private Const VERSION_URL As String = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/VERSION.json"
Private Const ARTIFACT_URL As String = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/dist/PCD-Excel-Version-latest.xlsm"

Public Sub CheckForUpdate()
    MsgBox "Update engine source is ready. It will be embedded in the next project-specific VBA build.", vbInformation, "PCD Excel"
End Sub
