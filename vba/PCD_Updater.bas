Attribute VB_Name = "PCD_Updater"
Option Explicit

Private Const VERSION_URL As String = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/VERSION.json"
Private Const ARTIFACT_URL As String = "https://raw.githubusercontent.com/milance78/PCD-excel-version-/main/dist/PCD-Excel-Version-latest.xlsm"

Public Sub CheckForUpdate()
    MsgBox "Le moteur de mise à jour est prêt dans les sources et sera intégré au projet VBA.", vbInformation, "PCD Excel"
End Sub
