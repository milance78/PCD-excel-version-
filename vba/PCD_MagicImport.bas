Attribute VB_Name = "Module1"
Option Explicit

Private Function Clean(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, ChrW(&HFFFD), "'")
    Do While InStr(s, "  ") > 0: s = Replace(s, "  ", " "): Loop
    Clean = Trim$(s)
End Function

Private Function Meaningful(ByVal s As String) As String
    s = Clean(s)
    If s = "" Or s = "--" Or s = "-" Then
        Meaningful = ""
    Else
        Meaningful = s
    End If
End Function

Private Function IsPlaceholder(ByVal s As String) As Boolean
    s = Clean(s)
    IsPlaceholder = (s Like "preemptive_first_name preemptive_last_name") Or _
                    (LCase$(s) = "nom de famille") Or _
                    (LCase$(s) = "last name") Or _
                    (LCase$(s) = "first name") Or _
                    (LCase$(s) = "action")
End Function

Private Function MeaningfulBusiness(ByVal s As String) As String
    s = Meaningful(s)
    If s <> "" And Not IsPlaceholder(s) Then MeaningfulBusiness = s
End Function

Private Function FirstValue(ParamArray values() As Variant) As String
    Dim i As Long, s As String
    For i = LBound(values) To UBound(values)
        s = Meaningful(CStr(values(i)))
        If s <> "" Then FirstValue = s: Exit Function
    Next
End Function

Private Function NormalizeText(ByVal s As String) As String
    NormalizeText = Replace(Replace(Replace(s, vbCrLf, vbLf), vbCr, ""), ChrW(160), " ")
    NormalizeText = Replace(NormalizeText, ChrW(&HFFFD), "'")
End Function

Private Function Rx(ByVal text As String, ByVal pattern As String, Optional ByVal n As Long = 1) As String
    Dim r As Object, m As Object
    Set r = CreateObject("VBScript.RegExp")
    r.Global = False: r.IgnoreCase = True: r.MultiLine = True: r.Pattern = pattern
    Set m = r.Execute(text)
    If m.Count > 0 And m(0).SubMatches.Count >= n Then Rx = CStr(m(0).SubMatches(n - 1))
End Function

Private Function ExtractAllRx(ByVal text As String, ByVal pattern As String, ByVal occurrence As Long) As String
    Dim r As Object, ms As Object
    Set r = CreateObject("VBScript.RegExp")
    r.Global = True: r.IgnoreCase = True: r.MultiLine = True: r.Pattern = pattern
    Set ms = r.Execute(text)
    If ms.Count >= occurrence Then ExtractAllRx = CStr(ms(occurrence - 1).SubMatches(0))
End Function

Private Function Section(ByVal text As String, ByVal startPattern As String, ParamArray endPatterns() As Variant) As String
    Dim r As Object, m As Object, rest As String, i As Long, p As String, pos As Long, best As Long
    Set r = CreateObject("VBScript.RegExp")
    r.Global = False: r.IgnoreCase = True: r.MultiLine = True: r.Pattern = startPattern
    Set m = r.Execute(text)
    If m.Count = 0 Then Exit Function
    rest = Mid$(text, m(0).FirstIndex + m(0).Length + 1)
    best = Len(rest) + 1
    For i = LBound(endPatterns) To UBound(endPatterns)
        r.Pattern = CStr(endPatterns(i))
        Set m = r.Execute(rest)
        If m.Count > 0 Then
            pos = m(0).FirstIndex + 1
            If pos < best Then best = pos
        End If
    Next
    If best <= Len(rest) Then Section = Left$(rest, best - 1) Else Section = rest
End Function

Private Function SplitLoose(ByVal line As String) As Collection
    Dim c As New Collection, r As Object, ms As Object, m As Object, s As String
    s = Trim$(line)
    Set r = CreateObject("VBScript.RegExp")
    r.Global = True: r.Pattern = "[	]+| {2,}"
    Set ms = r.Split(s)
    For Each m In ms
        If Meaningful(CStr(m)) <> "" Then c.Add Meaningful(CStr(m))
    Next
    Set SplitLoose = c
End Function

Private Function ExtractLabelValue(ByVal text As String, ByVal labels As Variant) As String
    Dim lines() As String, i As Long, j As Long, label As String, line As String, c As Collection
    Dim value As String, pattern As String
    lines = Split(text, vbLf)
    For j = LBound(labels) To UBound(labels)
        label = CStr(labels(j))
        For i = LBound(lines) To UBound(lines)
            line = lines(i)
            pattern = "(^|\s{2,}|\t)" & EscapeRx(label) & "\s*(\t+| {2,}|:)\s*(.+)$"
            value = Meaningful(Rx(line, pattern, 3))
            If value <> "" Then ExtractLabelValue = value: Exit Function
            Set c = SplitLoose(line)
            Dim k As Long
            For k = 1 To c.Count
                If LCase$(Clean(CStr(c(k)))) = LCase$(Clean(label)) Then
                    If k < c.Count Then
                        value = Meaningful(CStr(c(k + 1)))
                        If value <> "" Then ExtractLabelValue = value: Exit Function
                    End If
                End If
            Next
        Next
        pattern = "(^|\n)\s*" & EscapeRx(label) & "\s*\n\s*([^\n]+)"
        value = Meaningful(Rx(text, pattern, 2))
        If value <> "" Then ExtractLabelValue = value: Exit Function
    Next
End Function

Private Function ExtractBusinessLabelValue(ByVal text As String, ByVal labels As Variant) As String
    Dim remaining As String, candidate As String, p As Long
    remaining = text
    Dim attempt As Long
    For attempt = 1 To 12
        candidate = ExtractLabelValue(remaining, labels)
        If candidate = "" Then Exit Function
        If MeaningfulBusiness(candidate) <> "" Then
            ExtractBusinessLabelValue = MeaningfulBusiness(candidate)
            Exit Function
        End If
        p = InStr(1, remaining, candidate, vbTextCompare)
        If p <= 0 Then Exit Function
        remaining = Mid$(remaining, p + Len(candidate))
    Next
End Function

Private Function EscapeRx(ByVal s As String) As String
    Dim chars As Variant, i As Long
    chars = Array("\", ".", "+", "*", "?", "^", "$", "(", ")", "[", "]", "{", "}", "|")
    EscapeRx = s
    For i = LBound(chars) To UBound(chars)
        EscapeRx = Replace(EscapeRx, chars(i), "\" & chars(i))
    Next
End Function

Private Function HeaderData(ByVal block As String, ByVal requiredHeaders As Variant, ByRef headerLine As String, ByRef valueLine As String) As Boolean
    Dim lines() As String, i As Long, h As Variant, ok As Boolean
    lines = Split(block, vbLf)
    For i = LBound(lines) To UBound(lines)
        ok = True
        For Each h In requiredHeaders
            If InStr(1, lines(i), CStr(h), vbTextCompare) = 0 Then ok = False: Exit For
        Next
        If ok Then
            headerLine = lines(i)
            Dim j As Long
            For j = i + 1 To Application.Min(i + 5, UBound(lines))
                If Trim$(lines(j)) <> "" Then valueLine = lines(j): Exit For
            Next
            HeaderData = True
            Exit Function
        End If
    Next
End Function

Private Function NormalizeBelgianMobile(ByVal value As String) As String
    Dim digits As String
    digits = DigitsOnly(value)
    If digits Like "00324########" Then NormalizeBelgianMobile = "0" & Mid$(digits, 5): Exit Function
    If digits Like "324########" Then NormalizeBelgianMobile = "0" & Mid$(digits, 4): Exit Function
    If digits Like "04########" Then NormalizeBelgianMobile = digits
End Function

Private Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then DigitsOnly = DigitsOnly & ch
    Next
End Function

Private Function ExtractPreferredPhone(ByVal block As String) As String
    Dim r As Object, ms As Object, m As Object, n As String, digits As String
    Set r = CreateObject("VBScript.RegExp")
    r.Global = True: r.IgnoreCase = True
    r.Pattern = "(\+32|0032)4\d{8}\b|\b04\d{8}\b"
    Set ms = r.Execute(block)
    For Each m In ms
        n = NormalizeBelgianMobile(m.Value)
        If n <> "" Then ExtractPreferredPhone = n: Exit Function
    Next
    r.Pattern = "(\+32|0032|0)\s*4(?:[\s./()-]*\d){8}"
    Set ms = r.Execute(block)
    For Each m In ms
        n = NormalizeBelgianMobile(m.Value)
        If n <> "" Then ExtractPreferredPhone = n: Exit Function
    Next
    r.Pattern = "(?:\+|00)?\d[\d\s()./-]{7,}\d"
    Set ms = r.Execute(block)
    For Each m In ms
        digits = DigitsOnly(m.Value)
        If Len(digits) >= 9 And digits <> String(Len(digits), "0") Then ExtractPreferredPhone = digits: Exit Function
    Next
End Function

Private Function ExtractFiberServiceId(ByVal text As String) As String
    Dim r As Object, ms As Object, m As Object
    Set r = CreateObject("VBScript.RegExp")
    r.Global = True: r.IgnoreCase = True
    r.Pattern = "Service\s+ID\s*=\s*(1\d{11})\b"
    Set ms = r.Execute(text)
    If ms.Count > 0 Then ExtractFiberServiceId = Meaningful(ms(0).SubMatches(0)): Exit Function
    r.Global = False
    r.Pattern = "Access\s+TYPE\s*=\s*Fiber[\s\S]{0,260}?Service\s+ID\s*=\s*(\d{9,15})"
    Set ms = r.Execute(text)
    If ms.Count > 0 Then ExtractFiberServiceId = Meaningful(ms(0).SubMatches(0))
End Function

Private Function ParseCustomer(ByVal text As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim block As String, h As String, v As String, c As Collection
    block = Section(text, "(^|\n)\s*Information client\b", "(^|\n)\s*Contact Client\b", "(^|\n)\s*Informations d'intervention\b")
    If HeaderData(block, Array("ID client", "Partner Account ID"), h, v) Then
        Set c = SplitLoose(v)
        If c.Count >= 1 Then d("clientID") = Meaningful(CStr(c(1)))
        If c.Count >= 4 Then d("firstName") = Meaningful(CStr(c(4)))
        If c.Count >= 3 Then d("lastName") = Meaningful(CStr(c(3)))
    End If
    Set ParseCustomer = d
End Function

Private Function ParseContact(ByVal text As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim block As String, h As String, v As String, c As Collection, emailPos As Long
    block = Section(text, "(^|\n)\s*Contact Client\b", "(^|\n)\s*Informations d'intervention\b", "(^|\n)\s*Adresse d'installation\b")
    HeaderData block, Array("Nom de la personne de contact", "N° de GSM"), h, v
    Set c = SplitLoose(v)
    Dim phoneRaw As String, beforePhone As String, idx As Long
    phoneRaw = Rx(v, "(?:\+|00)?\d[\d\s()./-]{7,}\d")
    idx = InStr(1, v, phoneRaw, vbTextCompare)
    emailPos = InStr(1, v, "@", vbTextCompare)
    If idx > 0 Then beforePhone = Left$(v, idx - 1) ElseIf emailPos > 0 Then beforePhone = Left$(v, emailPos - 1) Else beforePhone = v
    Set c = SplitLoose(beforePhone)
    If c.Count >= 3 Then d("contactName") = Meaningful(CStr(c(c.Count)))
    d("phone") = ExtractPreferredPhone(block)
    Set ParseContact = d
End Function

Private Function ParseNewAddress(ByVal text As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim block As String, h As String, v As String, headers As Variant, hc() As String, vc() As String
    block = Section(text, "(^|\n)\s*Nouvelle adresse\b", "(^|\n)\s*Ancienne adresse\b", "(^|\n)\s*Manual TSI/Design reason\b", "(^|\n)\s*Stop Servicing Copper Date\b", "(^|\n)\s*Order Viewer Links\b")
    If block = "" Then Set ParseNewAddress = d: Exit Function
    Dim infra As String
    infra = Rx(block, "(^|\n|\s)(Fiber|Fibre|Copper|Cuivre)(?=\s|\n|$)", 2)
    If InStr(1, infra, "fiber", vbTextCompare) > 0 Or InStr(1, infra, "fibre", vbTextCompare) > 0 Then d("infrastructure") = "fiber"
    If InStr(1, infra, "copper", vbTextCompare) > 0 Or InStr(1, infra, "cuivre", vbTextCompare) > 0 Then d("infrastructure") = "copper"
    If Not HeaderData(block, Array("Nom de la rue", "LOM Key"), h, v) Then Set ParseNewAddress = d: Exit Function
    headers = Array("Pays","Code postal","Nom de la ville","Nom de la rue","N° de maison","N° de maison alphanumérique","Mail Box","Etage","Appartement","N° de bloc","LOM Key","subArea","Indicateur MDU/SDU","ZONE:")
    If InStr(h, vbTab) > 0 And InStr(v, vbTab) > 0 Then
        hc = Split(h, vbTab): vc = Split(v, vbTab)
        Dim i As Long, j As Long
        For i = LBound(headers) To UBound(headers)
            For j = LBound(hc) To UBound(hc)
                If LCase$(Clean(hc(j))) = LCase$(Clean(CStr(headers(i)))) Then
                    If j <= UBound(vc) Then d(CStr(headers(i))) = Meaningful(vc(j))
                End If
            Next
        Next
    Else
        Dim row As Object: Set row = ParseCollapsedAddressRow(v)
        Dim key As Variant
        For Each key In row.Keys: d(CStr(key)) = row(key): Next
    End If
    Dim street As String, num As String, alpha As String, zip As String, city As String, house As String
    street = Meaningful(DictGet(d,"Nom de la rue")): num = Meaningful(DictGet(d,"N° de maison"))
    alpha = Meaningful(DictGet(d,"N° de maison alphanumérique")): zip = Meaningful(DictGet(d,"Code postal"))
    city = Meaningful(DictGet(d,"Nom de la ville")): house = num & alpha
    d("streetName")=street: d("streetNumber")=num: d("streetAlpha")=alpha: d("postalCode")=zip: d("city")=city
    d("mainAddress") = Clean(street & " " & house)
    If zip <> "" Or city <> "" Then d("mainAddress") = d("mainAddress") & ", " & Clean(zip & " " & city)
    d("mailbox")=Meaningful(DictGet(d,"Mail Box")): d("floor")=Meaningful(DictGet(d,"Etage"))
    d("apartment")=Meaningful(DictGet(d,"Appartement")): d("blockNumber")=Meaningful(DictGet(d,"N° de bloc"))
    d("LOMKey")=Meaningful(DictGet(d,"LOM Key"))
    Set ParseNewAddress = d
End Function

Private Function ParseCollapsedAddressRow(ByVal line As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim zip As String, p As Long, afterZip As String, streetPos As Long, fromStreet As String
    zip = Rx(line, "\b(\d{4})\b")
    If zip = "" Then Set ParseCollapsedAddressRow = d: Exit Function
    p = InStr(1, line, zip, vbTextCompare): afterZip = Trim$(Mid$(line, p + Len(zip)))
    streetPos = RxIndex(afterZip, "\b(Rue|Avenue|Boulevard|Chaussée|Chaussee|Square|Clos|Place|Quai|Route|Chemin|Allée|Allee|Drève|Dreve|Sentier|Impasse|Laan|Straat|Steenweg|Weg|Plein)\b")
    If streetPos < 1 Then Set ParseCollapsedAddressRow = d: Exit Function
    Dim city As String: city = Meaningful(Left$(afterZip, streetPos - 1)): fromStreet = Trim$(Mid$(afterZip, streetPos))
    Dim r As String: r = Rx(fromStreet, "^(.*?)\s+(\d+[A-Za-z]?)\s+(?:(.*?)\s+)?(\d{5,})\s+([A-Z0-9]+)\s+(SDU|MDU)\s+(.+)$")
    If r = "" Then Set ParseCollapsedAddressRow = d: Exit Function
    d("Code postal")=zip: d("Nom de la ville")=city
    d("Nom de la rue")=Rx(fromStreet, "^(.*?)\s+(\d+[A-Za-z]?)\s+(?:(.*?)\s+)?(\d{5,})\s+([A-Z0-9]+)\s+(SDU|MDU)\s+(.+)$",1)
    d("N° de maison")=Rx(fromStreet, "^(.*?)\s+(\d+[A-Za-z]?)\s+(?:(.*?)\s+)?(\d{5,})\s+([A-Z0-9]+)\s+(SDU|MDU)\s+(.+)$",2)
    d("N° de maison alphanumérique")=Replace(DictGet(d,"N° de maison"), DigitsOnly(DictGet(d,"N° de maison")), "")
    Dim detail As String: detail=Rx(fromStreet, "^(.*?)\s+(\d+[A-Za-z]?)\s+(?:(.*?)\s+)?(\d{5,})\s+([A-Z0-9]+)\s+(SDU|MDU)\s+(.+)$",3)
    Dim c As Collection: Set c=SplitLoose(detail)
    If c.Count>=1 Then d("Mail Box")=c(1)
    If c.Count>=2 Then d("Etage")=c(2)
    If c.Count>=3 Then d("Appartement")=c(3)
    If c.Count>=4 Then d("N° de bloc")=c(4)
    d("LOM Key")=Rx(fromStreet, "^(.*?)\s+(\d+[A-Za-z]?)\s+(?:(.*?)\s+)?(\d{5,})\s+([A-Z0-9]+)\s+(SDU|MDU)\s+(.+)$",4)
    Set ParseCollapsedAddressRow = d
End Function

Private Function RxIndex(ByVal text As String, ByVal pattern As String) As Long
    Dim r As Object, ms As Object
    Set r=CreateObject("VBScript.RegExp"): r.Global=False: r.IgnoreCase=True: r.Pattern=pattern
    Set ms=r.Execute(text)
    If ms.Count>0 Then RxIndex=ms(0).FirstIndex+1
End Function

Private Function DictGet(ByVal d As Object, ByVal key As String) As String
    If d.Exists(key) Then DictGet=CStr(d(key))
End Function

Private Function ParseServiceOrderAddress(ByVal text As String) As Object
    Dim d As Object: Set d=CreateObject("Scripting.Dictionary")
    Dim block As String
    block=Section(text,"(^|\n)\s*Informations ['’]Service Order['’]","(^|\n)\s*Stop Servicing Copper Date\b","(^|\n)\s*Actions d'ordre\b","(^|\n)\s*Order Viewer Links\b")
    If block="" Then Set ParseServiceOrderAddress=d: Exit Function
    Dim street As String, num As String, zip As String, city As String
    street=Rx(block,"Nouvelle adresse\s+Nom de la rue\s+(.+?)\s*,\s*N[°o] de maison\s+([^\n,]+)[\s\S]{0,160}?Code postal\s+(\d{4})\s*,\s*Nom de la ville\s+([^\n,]+)",1)
    num=Rx(block,"Nouvelle adresse\s+Nom de la rue\s+(.+?)\s*,\s*N[°o] de maison\s+([^\n,]+)[\s\S]{0,160}?Code postal\s+(\d{4})\s*,\s*Nom de la ville\s+([^\n,]+)",2)
    zip=Rx(block,"Nouvelle adresse\s+Nom de la rue\s+(.+?)\s*,\s*N[°o] de maison\s+([^\n,]+)[\s\S]{0,160}?Code postal\s+(\d{4})\s*,\s*Nom de la ville\s+([^\n,]+)",3)
    city=Rx(block,"Nouvelle adresse\s+Nom de la rue\s+(.+?)\s*,\s*N[°o] de maison\s+([^\n,]+)[\s\S]{0,160}?Code postal\s+(\d{4})\s*,\s*Nom de la ville\s+([^\n,]+)",4)
    If street="" Then Set ParseServiceOrderAddress=d: Exit Function
    d("infrastructure")="fiber": d("streetName")=street: d("streetNumber")=Rx(num,"^(\d+)")
    d("streetAlpha")=Replace(num,d("streetNumber"),""): d("postalCode")=zip: d("city")=Replace(city,", Pays","")
    d("mainAddress")=Clean(street & " " & d("streetNumber") & d("streetAlpha")) & ", " & Clean(zip & " " & d("city"))
    d("LOMKey")=Meaningful(Rx(block,"LOM Key\s*:\s*([^\n]+)"))
    Set ParseServiceOrderAddress=d
End Function

Private Function ExtractSafeDescription(ByVal text As String) As String
    Dim block As String, lines() As String, i As Long, c As Collection, j As Long, v As String
    block=Section(text,"(^|\n)\s*Informations d'intervention\b","(^|\n)\s*Adresse d'installation\b","(^|\n)\s*Order Viewer Links\b","(^|\n)\s*Envoi Notification\b","(^|\n)\s*Mise à jour intervention\b")
    If block<>"" Then
        lines=Split(block,vbLf)
        For i=LBound(lines) To UBound(lines)
            Set c=SplitLoose(lines(i))
            For j=1 To c.Count
                If LCase$(CStr(c(j)))="descriptions" Or LCase$(CStr(c(j)))="description d'intervention" Then
                    If j<c.Count Then
                        v=Meaningful(CStr(c(j+1)))
                        If v<>"" Then ExtractSafeDescription=v: Exit Function
                    End If
                End If
            Next
        Next
        v=Meaningful(Rx(block,"(?:Descriptions|Description d'intervention)\s*(?:\t+| {2,}|:)\s*([^\n\t]+?)(?=(?:\t+| {2,})(?:Date de création|Date souhaitée|Statut|Source|Priorité|Remarques)\b|$)"))
        If v<>"" Then ExtractSafeDescription=v: Exit Function
    End If
    If Rx(text,"\bDescriptions\b")<>"" Then
        lines=Split(text,vbLf)
        For i=LBound(lines) To UBound(lines)
            v=Meaningful(lines(i))
            If v<>"" And LCase$(v)<>"powered by" And v<>"Information client" And v<>"Contact Client" And v<>"Informations d'intervention" And v<>"Adresse d'installation" Then
                If Not v Like "[A-Z0-9_-][A-Z0-9_-][A-Z0-9_-][A-Z0-9_-][A-Z0-9_-][A-Z0-9_-]*" Then ExtractSafeDescription=v: Exit Function
            End If
        Next
    End If
End Function

Private Function ParseSafe(ByVal text As String) As Object
    Dim d As Object: Set d=CreateObject("Scripting.Dictionary")
    Dim cust As Object, contact As Object, addr As Object, soAddr As Object
    Set cust=ParseCustomer(text): Set contact=ParseContact(text): Set addr=ParseNewAddress(text): Set soAddr=ParseServiceOrderAddress(text)
    Dim a As Object: If addr.Exists("mainAddress") Then Set a=addr Else Set a=soAddr
    d("interventionId")=ExtractLabelValue(text,Array("ID d'intervention"))
    d("oagID")=FirstValue(ExtractLabelValue(text,Array("OAG_ID","OAG ID","Provisioning Order Id")),ExtractLabelValue(text,Array("ORDER_NUM")))
    d("clientID")=FirstValue(MeaningfulBusiness(DictGet(cust,"clientID")),ExtractBusinessLabelValue(text,Array("ID client")))
    d("na")=ExtractLabelValue(text,Array("NA / CID"))
    d("cid")=ExtractFiberServiceId(text)
    d("clientName")=FirstValue(Meaningful(contact("contactName")),Clean(DictGet(cust,"firstName") & " " & DictGet(cust,"lastName")))
    d("phone")=Meaningful(DictGet(contact,"phone"))
    If a.Exists("infrastructure") Then d("infrastructure")=a("infrastructure")
    Dim key As Variant
    For Each key In Array("streetName","streetNumber","streetAlpha","postalCode","city","mainAddress","mailbox","floor","apartment","blockNumber","LOMKey")
        If a.Exists(key) Then d(key)=a(key)
    Next
    d("descriptionFr")=ExtractSafeDescription(text)
    Dim updateBlock As String, remarks As String
    updateBlock=Section(text,"(^|\n)\s*Mise à jour intervention\b","(^|\n)\s*Retour\b","(^|\n)\s*Ajouter Tâche\b")
    remarks=Meaningful(Rx(updateBlock,"(^|\n)\s*Remarques\s*(?:\t+| {2,}|\n)\s*([\s\S]*?)(?=(?:\t+| {2,}|\n)\s*(?:Action|Route vers|A la clôture|Client en ligne)\b|$)",2))
    If remarks<>"" And Not LCase$(Left$(remarks,6))="action" Then d("comment")="Remarque préexistante:" & vbLf & """" & remarks & """"
    Set ParseSafe=d
End Function

Private Function ParseWorkItem(ByVal text As String) As Object
    Dim d As Object: Set d=CreateObject("Scripting.Dictionary")
    d("snowMentioned")=FirstValue(ExtractLabelValue(text,Array("SNOW_ID")),ExtractLabelValue(text,Array("TICKET_NUM")))
    d("interventionId")=ExtractLabelValue(text,Array("INTERVENTION_ID"))
    d("oagID")=FirstValue(ExtractLabelValue(text,Array("OAG_ID")),ExtractLabelValue(text,Array("ORDER_NUM")))
    d("clientID")=FirstValue(ExtractBusinessLabelValue(text,Array("CUSTOMER_ID")),ExtractBusinessLabelValue(text,Array("CUST_NUM")))
    d("na")=ExtractLabelValue(text,Array("NA"))
    d("descriptionEn")=FirstValue(ExtractLabelValue(text,Array("INTERVENTION_DESCRIPTION")),ExtractLabelValue(text,Array("NPS_EXC_DESCRIPTION")))
    Dim tech As String: tech=ExtractLabelValue(text,Array("TECHNOLOGY"))
    If InStr(1,tech,"fiber",vbTextCompare)>0 Or InStr(1,tech,"fibre",vbTextCompare)>0 Then d("infrastructure")="fiber"
    If InStr(1,tech,"copper",vbTextCompare)>0 Or InStr(1,tech,"cuivre",vbTextCompare)>0 Then d("infrastructure")="copper"
    If InStr(1,text,"MOBILE VIKINGS",vbTextCompare)>0 Then d("clientName")=ExtractBusinessLabelValue(text,Array("SCOPE"))
    Set ParseWorkItem=d
End Function

Private Function NormalizeNetwork(ByVal text As String) As String
    Dim op As String
    op=UCase$(FirstValue(ExtractLabelValue(text,Array("Opérateur","Operateur","Operator","Réseau","Reseau")),ExtractLabelValue(text,Array("LIST_FILTER"))))
    If InStr(op,"MOBILE VIKINGS")>0 Then NormalizeNetwork="mobileVikings": Exit Function
    If InStr(op,"SCARLET")>0 Then NormalizeNetwork="scarlet": Exit Function
    If InStr(op,"OLO")>0 Then NormalizeNetwork="otherOlo": Exit Function
    If op="PXS" Or InStr(op,"PROXIMUS")>0 Then NormalizeNetwork="proximus"
End Function

Private Function NormalizeStatus(ByVal text As String) As String
    Dim raw As String
    raw=UCase$(FirstValue(ExtractLabelValue(text,Array("Status")),ExtractLabelValue(text,Array("Statut")),ExtractLabelValue(text,Array("NPS_STATUS"))))
    If InStr(raw,"DONE")>0 Or InStr(raw,"TERMIN")>0 Or InStr(raw,"CLOSED")>0 Or InStr(raw,"RESOLVED")>0 Then NormalizeStatus="completed": Exit Function
    If InStr(raw,"WAIT")>0 Or InStr(raw,"PENDING")>0 Or InStr(raw,"HOLD")>0 Or InStr(raw,"CURE CONTACT")>0 Or InStr(raw,"INPROGRESS")>0 Or InStr(raw,"IN PROGRESS")>0 Then NormalizeStatus="on hold": Exit Function
    If InStr(raw,"ROUTE")>0 Or InStr(raw,"TRANSFER")>0 Or InStr(raw,"TRANSMIS")>0 Then NormalizeStatus="transferred"
End Function

Private Function DictFirst(ByVal a As Object, ByVal b As Object, ByVal key As String) As String
    DictFirst=FirstValue(DictGet(a,key),DictGet(b,key))
End Function

Private Function ParseStatusValue(ByVal safe As Object, ByVal work As Object, ByVal text As String) As Object
    Dim d As Object: Set d=CreateObject("Scripting.Dictionary")
    d("interventionId")=FirstValue(DictGet(work,"interventionId"),DictGet(safe,"interventionId"))
    d("snowMentioned")=DictGet(work,"snowMentioned")
    d("oagID")=FirstValue(DictGet(work,"oagID"),DictGet(safe,"oagID"))
    d("clientID")=FirstValue(MeaningfulBusiness(DictGet(work,"clientID")),MeaningfulBusiness(DictGet(safe,"clientID")))
    d("clientName")=FirstValue(DictGet(safe,"clientName"),DictGet(work,"clientName"))
    d("infrastructure")=FirstValue(DictGet(safe,"infrastructure"),DictGet(work,"infrastructure"))
    d("na")=FirstValue(DictGet(safe,"na"),DictGet(work,"na"))
    d("cid")=FirstValue(DictGet(safe,"cid"),DictGet(work,"cid"))
    Dim k As Variant
    For Each k In Array("mainAddress","streetName","streetNumber","streetAlpha","postalCode","city","mailbox","floor","apartment","blockNumber","LOMKey","phone","descriptionFr","descriptionEn","comment")
        d(CStr(k))=FirstValue(DictGet(safe,CStr(k)),DictGet(work,CStr(k)))
    Next
    If d("infrastructure")="fiber" Then d("na")=""
    If d("infrastructure")="copper" Then d("cid")=""
    d("network")=NormalizeNetwork(text)
    d("status")=NormalizeStatus(text)
    If d("descriptionFr")<>"" Then d("interventionDescription")=d("descriptionFr") Else d("interventionDescription")=d("descriptionEn")
    Set ParseStatusValue=d
End Function

Private Sub WriteValue(ByVal ws As Worksheet, ByVal labelCell As String, ByVal valueCell As String, ByVal value As String)
    ws.Range(valueCell).Value = value
End Sub

Public Sub ParseMagicImportText(ByVal rawText As String, Optional ByVal showMessage As Boolean = False)
    Dim text As String, safe As Object, work As Object, result As Object, ws As Worksheet
    text=NormalizeText(rawText)
    If Len(Trim$(text))<10 Then Exit Sub
    Set safe=ParseSafe(text): Set work=ParseWorkItem(text): Set result=ParseStatusValue(safe,work,text)
    Set ws=ThisWorkbook.Worksheets("Intervention en cours")

    WriteValue ws,"B4","C4",DictGet(result,"interventionId")
    WriteValue ws,"F4","G4",DictGet(result,"oagID")
    WriteValue ws,"B5","C5",DictGet(result,"clientID")
    WriteValue ws,"F5","G5",DictGet(result,"snowMentioned")
    WriteValue ws,"B6","C6",DictGet(result,"clientName")
    WriteValue ws,"F6","G6",DictGet(result,"phone")
    WriteValue ws,"B7","C7",DictGet(result,"infrastructure")
    WriteValue ws,"F7","G7",DictGet(result,"network")
    WriteValue ws,"B8","C8",DictGet(result,"status")

    WriteValue ws,"B11","C11",DictGet(result,"streetName")
    WriteValue ws,"F11","G11",DictGet(result,"streetNumber")
    WriteValue ws,"B12","C12",DictGet(result,"streetAlpha")
    WriteValue ws,"F12","G12",DictGet(result,"postalCode")
    WriteValue ws,"B13","C13",DictGet(result,"city")
    WriteValue ws,"F13","G13",DictGet(result,"mailbox")
    WriteValue ws,"B14","C14",DictGet(result,"floor")
    WriteValue ws,"F14","G14",DictGet(result,"apartment")
    WriteValue ws,"B15","C15",DictGet(result,"blockNumber")
    WriteValue ws,"F15","G15",DictGet(result,"LOMKey")
    WriteValue ws,"B17","C17",DictGet(result,"cid")
    WriteValue ws,"F17","G17",DictGet(result,"na")
    ws.Range("B20").Value=DictGet(result,"interventionDescription")
    ws.Range("B25").Value=DictGet(result,"comment")
    ws.Range("B30").Value=ws.Range("B30").Value
    ws.Activate
    If showMessage Then MsgBox "Import terminé." & vbCrLf & "Les champs ont été remplis selon la logique du Smart Import de l'application.", vbInformation, "PCD Smart Import"
End Sub

Public Sub OpenMagicImport()
    ThisWorkbook.Worksheets("Magic Import").Activate
    ThisWorkbook.Worksheets("Magic Import").Range("B5").Select
End Sub

Public Sub ClearMagicImport()
    ThisWorkbook.Worksheets("Magic Import").Range("B5").ClearContents
    ThisWorkbook.Worksheets("Intervention en cours").Range("C4:C8,G4:G8,C11:C15,G11:G15,C17,G17,B20,B25").ClearContents
End Sub

Public Sub ImportMagicFromSheet()
    Dim ws As Worksheet
    Set ws=ThisWorkbook.Worksheets("Magic Import")
    ParseMagicImportText CStr(ws.Range("B5").Value), True
End Sub
