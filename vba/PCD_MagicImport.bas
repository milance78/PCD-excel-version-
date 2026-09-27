Attribute VB_Name = "Module1"
Option Explicit

Private Function Clean(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, "<br>", " ")
    s = Replace(s, "<br/>", " ")
    s = Replace(s, "<br />", " ")
    s = Replace(s, "**", "")
    s = Replace(s, ChrW(&HFFFD), "'")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    Clean = Trim$(s)
End Function

Private Function ValOk(ByVal s As String) As String
    s = Clean(s)
    If s = "" Or s = "-" Or s = "--" Then ValOk = "" Else ValOk = s
End Function

Private Function Cells(ByVal line As String) As Collection
    Dim c As New Collection, a() As String, i As Long, s As String
    s = Trim$(line)
    If Left$(s, 1) = "|" Then s = Mid$(s, 2)
    If Right$(s, 1) = "|" Then s = Left$(s, Len(s) - 1)
    a = Split(s, "|")
    For i = LBound(a) To UBound(a): c.Add ValOk(a(i)): Next
    Set Cells = c
End Function

Private Function IsSep(ByVal s As String) As Boolean
    Dim c As Collection, i As Long, x As String
    If InStr(s, "|") = 0 Then Exit Function
    Set c = Cells(s)
    If c.Count = 0 Then IsSep = True: Exit Function
    For i = 1 To c.Count
        x = Replace(Replace(CStr(c(i)), "-", ""), ":", "")
        x = Replace(x, " ", "")
        If x <> "" Then Exit Function
    Next
    IsSep = True
End Function

Private Function FindNextCell(ByVal text As String, ByVal label As String) As String
    Dim a() As String, i As Long, c As Collection, j As Long, x As String
    a = Split(Replace(text, vbCr, ""), vbLf)
    For i = LBound(a) To UBound(a)
        If InStr(1, a(i), "|", vbTextCompare) > 0 Then
            Set c = Cells(a(i))
            For j = 1 To c.Count
                If LCase$(Clean(CStr(c(j)))) = LCase$(Clean(label)) Then
                    If j < c.Count Then
                        x = ValOk(CStr(c(j + 1)))
                        If x <> "" Then FindNextCell = x: Exit Function
                    End If
                End If
            Next
        ElseIf InStr(1, a(i), label, vbTextCompare) > 0 Then
            x = Trim$(Mid$(a(i), InStr(1, a(i), label, vbTextCompare) + Len(label)))
            x = Replace(x, ":", "")
            x = ValOk(x)
            If x <> "" Then FindNextCell = x: Exit Function
        End If
    Next
End Function

Private Function TableValue(ByVal text As String, ByVal header As String) As String
    Dim a() As String, i As Long, h As Collection, d As Collection, j As Long, k As Long
    a = Split(Replace(text, vbCr, ""), vbLf)
    For i = LBound(a) To UBound(a) - 1
        If InStr(1, a(i), "|", vbTextCompare) > 0 And InStr(1, a(i), header, vbTextCompare) > 0 Then
            If Not IsSep(a(i)) Then
                Set h = Cells(a(i))
                j = i + 1
                Do While j <= UBound(a) And (Trim$(a(j)) = "" Or IsSep(a(j))): j = j + 1: Loop
                If j <= UBound(a) Then
                    Set d = Cells(a(j))
                    For k = 1 To h.Count
                        If LCase$(Clean(CStr(h(k)))) = LCase$(Clean(header)) Then
                            If k <= d.Count Then TableValue = ValOk(CStr(d(k))): Exit Function
                        End If
                    Next
                End If
            End If
        End If
    Next
End Function

Private Function Rx(ByVal text As String, ByVal pattern As String, Optional ByVal n As Long = 1) As String
    Dim r As Object, m As Object
    Set r = CreateObject("VBScript.RegExp")
    r.Global = False: r.IgnoreCase = True: r.MultiLine = True: r.Pattern = pattern
    Set m = r.Execute(text)
    If m.Count > 0 Then
        If m(0).SubMatches.Count >= n Then Rx = CStr(m(0).SubMatches(n - 1))
    End If
End Function

Private Function NormalizePhone(ByVal s As String) As String
    s = Replace(Clean(s), " ", "")
    If Left$(s, 5) = "00324" Then NormalizePhone = "04" & Mid$(s, 6): Exit Function
    If Left$(s, 3) = "324" Then NormalizePhone = "04" & Mid$(s, 4): Exit Function
    If Left$(s, 2) = "04" Then NormalizePhone = s
End Function

Private Function SourceType(ByVal t As String) As String
    If InStr(1, t, "SNOW_ID", vbTextCompare) > 0 Or InStr(1, t, "SNOW_TITLE", vbTextCompare) > 0 Then
        SourceType = "SNOW"
    ElseIf InStr(1, t, "WORK ITEM TREATMENT", vbTextCompare) > 0 Or InStr(1, t, "NPS_EXCEPTION_CD", vbTextCompare) > 0 Then
        SourceType = "NPS"
    ElseIf InStr(1, t, "ORDER VIEWER LINKS", vbTextCompare) > 0 Or InStr(1, t, "NOUVELLE ADRESSE", vbTextCompare) > 0 Or InStr(1, t, "MISE À JOUR INTERVENTION", vbTextCompare) > 0 Then
        SourceType = "SAFE"
    ElseIf InStr(1, t, "FISISINTV", vbTextCompare) > 0 Or InStr(1, t, "SERVICE ORDER", vbTextCompare) > 0 Then
        SourceType = "ISIS"
    Else
        SourceType = "UNKNOWN"
    End If
End Function

Private Sub Put(ByVal ws As Worksheet, ByVal cell As String, ByVal v As String)
    ws.Range(cell).Value = v
End Sub

Public Sub ParseMagicImportText(ByVal rawText As String, Optional ByVal showMessage As Boolean = False)
    Dim t As String, infra As String, network As String, iid As String, oag As String
    Dim snow As String, cidClient As String, client As String, phone As String
    Dim desc As String, na As String, cid As String, street As String, house As String
    Dim alpha As String, postal As String, city As String, lom As String, box As String
    Dim floor As String, apt As String, block As String, status As String, addr As String
    Dim ws As Worksheet, src As String
    
    t = Replace(rawText, vbCr, "")
    If Len(Trim$(t)) < 10 Then Exit Sub
    src = SourceType(t)

    infra = TableValue(t, "Fiber")
    If infra = "" Then infra = TableValue(t, "Cuivre")
    If LCase$(infra) = "fiber" Or LCase$(infra) = "fibre" Then infra = "fiber"
    If LCase$(infra) = "copper" Or LCase$(infra) = "cuivre" Then infra = "copper"

    iid = FindNextCell(t, "ID d'intervention")
    If iid = "" Then iid = FindNextCell(t, "ID intervention")
    oag = FindNextCell(t, "Provisioning Order Id")
    If oag = "" Then oag = FindNextCell(t, "OAG_ID")
    If oag = "" Then oag = FindNextCell(t, "OAG ID")
    snow = FindNextCell(t, "SNOW_ID")
    cidClient = FindNextCell(t, "ID client")
    client = FindNextCell(t, "Nom de la personne de contact")
    If client = "" Then client = FindNextCell(t, "Nom du client")
    phone = NormalizePhone(FindNextCell(t, "N° de GSM"))
    If phone = "" Then phone = NormalizePhone(FindNextCell(t, "No de GSM"))
    If phone = "" Then phone = NormalizePhone(FindNextCell(t, "N° de téléphone"))

    desc = FindNextCell(t, "Descriptions")
    If desc = "" Then desc = FindNextCell(t, "Description")
    na = FindNextCell(t, "NA / CID")
    If na = "" Then na = FindNextCell(t, "NA")
    cid = Rx(t, "Service ID\s*=\s*(1\d{11})")
    If cid = "" Then cid = FindNextCell(t, "CID")

    street = TableValue(t, "Nom de la rue")
    house = TableValue(t, "N° de maison")
    alpha = TableValue(t, "N° de maison alphanumérique")
    postal = TableValue(t, "Code postal")
    city = TableValue(t, "Nom de la ville")
    box = TableValue(t, "Mail Box")
    floor = TableValue(t, "Etage")
    apt = TableValue(t, "Appartement")
    block = TableValue(t, "N° de bloc")
    lom = TableValue(t, "LOM Key")
    If street = "" Then street = FindNextCell(t, "Nom de la rue")
    If house = "" Then house = FindNextCell(t, "N° de maison")
    If alpha = "" Then alpha = FindNextCell(t, "Alphanumerique")
    If postal = "" Then postal = FindNextCell(t, "Code postal")
    If city = "" Then city = FindNextCell(t, "Nom de la ville")

    If infra = "fiber" Then
        na = ""
    ElseIf infra = "copper" Then
        cid = ""
    End If

    If street <> "" Then
        addr = street & " " & house & alpha
        If postal <> "" Or city <> "" Then addr = addr & ", " & postal & " " & city
    End If

    network = FindNextCell(t, "Opérateur")
    If network = "" Then network = FindNextCell(t, "Operateur")
    If network = "" Then network = FindNextCell(t, "Operator")
    If network = "" Then network = FindNextCell(t, "Réseau")
    If InStr(1, network, "Mobile Vikings", vbTextCompare) > 0 Then network = "mobileVikings"
    If InStr(1, network, "Scarlet", vbTextCompare) > 0 Then network = "scarlet"
    If InStr(1, network, "OLO", vbTextCompare) > 0 Then network = "otherOlo"
    If InStr(1, network, "Proximus", vbTextCompare) > 0 Or UCase$(network) = "PXS" Then network = "proximus"

    status = FindNextCell(t, "Statut")
    If LCase$(status) = "inprogress" Or InStr(1, status, "pending", vbTextCompare) > 0 Then status = "on hold"
    If LCase$(status) = "done" Or LCase$(status) = "closed" Or LCase$(status) = "resolved" Then status = "completed"

    Set ws = ThisWorkbook.Worksheets("Intervention en cours")
    Put ws, "B4", infra: Put ws, "F4", network
    Put ws, "B6", iid: Put ws, "F6", oag
    Put ws, "B8", snow: Put ws, "F8", cidClient
    Put ws, "B10", desc: Put ws, "B12", na: Put ws, "F12", cid
    Put ws, "B14", addr: Put ws, "F14", lom
    Put ws, "B16", box: Put ws, "C16", floor: Put ws, "D16", apt: Put ws, "E16", block: Put ws, "F16", phone
    Put ws, "B19", client: Put ws, "F25", status
    Put ws, "B28", "Source détectée: " & src
    ws.Activate
    
    If showMessage Then MsgBox "Import terminé." & vbCrLf & "Source: " & src, vbInformation, "PCD Magic Import"
End Sub

Public Sub OpenMagicImport()
    ThisWorkbook.Worksheets("Magic Import").Activate
    ThisWorkbook.Worksheets("Magic Import").Range("B5").Select
End Sub

Public Sub ClearMagicImport()
    ThisWorkbook.Worksheets("Magic Import").Range("B5").ClearContents
    ThisWorkbook.Worksheets("Intervention en cours").Range("B4:F28").ClearContents
End Sub


Public Sub ImportMagicFromSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Magic Import")
    ParseMagicImportText CStr(ws.Range("B5").Value), True
End Sub
