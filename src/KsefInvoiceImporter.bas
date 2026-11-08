Attribute VB_Name = "KsefInvoiceImporter"
Option Explicit

'==================================================================
'  KSeF Invoice Importer
'
'  Imports purchase invoices downloaded from KSeF (Krajowy System
'  e-Faktur, the Polish National e-Invoicing System) as XML files
'  (logical structure FA(2) / FA(3)) into an Excel list.
'
'  - regular invoices -> sheet "Invoices"
'  - corrections (KOR, KOR_ZAL, KOR_ROZ) -> sheet "Corrections"
'  - duplicates are detected by: seller NIP + invoice number
'  - new invoices are APPENDED; existing rows and any extra columns
'    you add on the right-hand side are never modified
'  - every processed file is recorded in sheet "Log"
'  - missing sheets (Start / Invoices / Corrections / Log) are
'    created automatically, so the module also works when imported
'    into an empty workbook
'
'  Requirements: Excel for Windows (uses MSXML2.DOMDocument.6.0,
'  Scripting.FileSystemObject, Scripting.Dictionary, VBScript.RegExp)
'  Entry point: ImportKsefInvoices
'==================================================================

Private Const SH_START As String = "Start"
Private Const SH_INV As String = "Invoices"
Private Const SH_COR As String = "Corrections"
Private Const SH_LOG As String = "Log"

Private Const ADR_FOLDER As String = "C6"
Private Const ADR_MYNIP As String = "C7"
Private Const ADR_LAST As String = "C9"
Private Const DEFAULT_FOLDER As String = "C:\xml_FAKTURY"

Private Const N_BASE As Long = 42      ' columns shared by both sheets
Private Const N_EXTRA As Long = 5      ' extra correction columns
Private Const EXTRA_POS As Long = 5    ' correction columns are inserted after this column

Private Const C_NUMBER As Long = 1
Private Const C_NIP As Long = 4
Private Const C_FILE As Long = 41

'------------------------------------------------------------------
'  Column layout
'------------------------------------------------------------------
Private Function BaseHeaders() As Variant
    BaseHeaders = Array("Invoice number", "KSeF number", "Invoice type", "Seller NIP", "Seller name", _
        "Issue date", "Sale / delivery date", "Service period from", "Service period to", "Payment due date", _
        "Settlement method", "Bank account country", "Bank account number", "Payment status", "Payment date", _
        "Currency", "Net total", "VAT total", "Gross total", "Amount outstanding", _
        "Net 23%", "VAT 23%", "Net 8%", "VAT 8%", "Net 5%", "VAT 5%", "Net 4% (taxi)", "VAT 4% (taxi)", _
        "Net OSS", "VAT OSS", "0% domestic", "0% intra-EU supply", "0% export", "Exempt", _
        "Not subject to VAT", "Not subject - art. 100(1)(4)", "Reverse charge", "Margin scheme", _
        "VAT in PLN (from invoice)", "XML created at", "File name", "Imported at")
End Function

Private Function ExtraHeaders() As Variant
    ExtraHeaders = Array("Correction reason", "Correction type", "Corrected invoice number", _
        "Corrected invoice issue date", "Corrected invoice KSeF number")
End Function

Private Function LogHeaders() As Variant
    LogHeaders = Array("Run at", "File name", "Status", "Details")
End Function

'------------------------------------------------------------------
'  MAIN PROCEDURE (assigned to the "Import invoices" button)
'------------------------------------------------------------------
Public Sub ImportKsefInvoices()
    EnsureWorkbookLayout

    Dim wsS As Worksheet, wsI As Worksheet, wsC As Worksheet, wsL As Worksheet
    Set wsS = ThisWorkbook.Worksheets(SH_START)
    Set wsI = ThisWorkbook.Worksheets(SH_INV)
    Set wsC = ThisWorkbook.Worksheets(SH_COR)
    Set wsL = ThisWorkbook.Worksheets(SH_LOG)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim folder As String
    folder = Trim$(CStr(wsS.Range(ADR_FOLDER).Value))
    If Len(folder) = 0 Or Not fso.FolderExists(folder) Then
        folder = PickFolder(folder)
        If Len(folder) = 0 Then Exit Sub
        wsS.Range(ADR_FOLDER).Value = folder
    End If

    Dim myNip As String
    myNip = OnlyDigits(CStr(wsS.Range(ADR_MYNIP).Value))

    Dim keys As Object, files As Object
    Set keys = CreateObject("Scripting.Dictionary")    ' NIP|NUMBER -> file name
    Set files = CreateObject("Scripting.Dictionary")   ' files already imported
    files.CompareMode = vbTextCompare

    LoadExisting wsI, False, keys, files
    LoadExisting wsC, True, keys, files

    Dim nInv As Long, nCor As Long, nDup As Long, nOld As Long, nSales As Long, nErr As Long
    Dim runTime As Date
    runTime = Now

    Application.ScreenUpdating = False
    Application.Cursor = xlWait
    On Error GoTo Finish

    Dim f As Object
    For Each f In fso.GetFolder(folder).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "xml" Then
            If files.Exists(f.Name) Then
                nOld = nOld + 1
            Else
                Select Case ProcessFile(f.Path, f.Name, myNip, keys, files, wsI, wsC, wsL, runTime)
                    Case 1: nInv = nInv + 1
                    Case 2: nCor = nCor + 1
                    Case 3: nDup = nDup + 1
                    Case 4: nSales = nSales + 1
                    Case Else: nErr = nErr + 1
                End Select
            End If
        End If
    Next f

Finish:
    Dim errMsg As String
    If Err.Number <> 0 Then errMsg = Err.Description
    On Error GoTo 0
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True

    If Len(errMsg) > 0 Then
        MsgBox "Import stopped: " & errMsg, vbCritical, "KSeF import"
        Exit Sub
    End If

    wsS.Range(ADR_LAST).Value = Format$(runTime, "yyyy-mm-dd hh:nn") & " | invoices: " & nInv & _
        " | corrections: " & nCor & " | duplicates: " & nDup & " | errors: " & nErr

    Dim msg As String
    msg = "Import finished." & vbLf & vbLf & _
          "Invoices added:  " & nInv & vbLf & _
          "Corrections added:  " & nCor & vbLf & _
          "Duplicates skipped:  " & nDup & vbLf & _
          "Files imported earlier:  " & nOld & vbLf
    If nSales > 0 Then msg = msg & "Sales invoices skipped:  " & nSales & vbLf
    msg = msg & "Errors:  " & nErr
    If nDup + nErr + nSales > 0 Then msg = msg & vbLf & vbLf & "See the Log sheet for details."
    MsgBox msg, IIf(nErr > 0, vbExclamation, vbInformation), "KSeF import"
End Sub

'------------------------------------------------------------------
'  Process one XML file
'  returns: 1 = invoice, 2 = correction, 3 = duplicate, 4 = sales invoice, 0 = error
'------------------------------------------------------------------
Private Function ProcessFile(ByVal path As String, ByVal fname As String, ByVal myNip As String, _
        ByVal keys As Object, ByVal files As Object, ByVal wsI As Worksheet, ByVal wsC As Worksheet, _
        ByVal wsL As Worksheet, ByVal runTime As Date) As Long

    On Error GoTo Failed
    ProcessFile = 0

    Dim doc As Object, r As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False
    If Not doc.Load(path) Then
        WriteLog wsL, runTime, fname, "Error", "Invalid XML file: " & Trim$(doc.parseError.reason)
        Exit Function
    End If
    Set r = doc.documentElement
    If r.baseName <> "Faktura" Then
        WriteLog wsL, runTime, fname, "Error", "Not a KSeF invoice (root element: " & r.baseName & ")"
        Exit Function
    End If

    Dim invNo As String, nip As String, kind As String
    invNo = X(r, "Fa/P_2")
    nip = X(r, "Podmiot1/DaneIdentyfikacyjne/NIP")
    kind = UCase$(X(r, "Fa/RodzajFaktury"))

    If Len(invNo) = 0 Then
        WriteLog wsL, runTime, fname, "Error", "Invoice number (P_2) is missing"
        Exit Function
    End If

    If Len(myNip) > 0 And OnlyDigits(nip) = myNip Then
        WriteLog wsL, runTime, fname, "Skipped", "Sales invoice (seller = my NIP): " & invNo
        ProcessFile = 4
        Exit Function
    End If

    Dim key As String
    key = MakeKey(nip, invNo)
    If keys.Exists(key) Then
        WriteLog wsL, runTime, fname, "Duplicate", _
            "Invoice " & invNo & " (NIP " & nip & ") is already on the list - file: " & keys(key)
        ProcessFile = 3
        Exit Function
    End If

    Dim isCor As Boolean
    isCor = (kind = "KOR" Or kind = "KOR_ZAL" Or kind = "KOR_ROZ")

    '--- basic data ----------------------------------------------------
    Dim v(1 To N_BASE) As Variant
    v(1) = invNo
    v(2) = KsefFromFileName(fname)
    v(3) = InvoiceTypeText(kind)
    v(4) = nip
    v(5) = X(r, "Podmiot1/DaneIdentyfikacyjne/Nazwa")
    v(6) = ToDate(X(r, "Fa/P_1"))
    v(7) = ToDate(X(r, "Fa/P_6"))
    v(8) = ToDate(X(r, "Fa/OkresFa/P_6_Od"))
    v(9) = ToDate(X(r, "Fa/OkresFa/P_6_Do"))
    v(10) = DueDate(r)
    v(11) = SettlementMethod(X(r, "Fa/Platnosc/FormaPlatnosci"), X(r, "Fa/Platnosc/PlatnoscInna"), _
                             X(r, "Fa/Platnosc/OpisPlatnosci"))

    Dim acc As String
    acc = X(r, "Fa/Platnosc/RachunekBankowy/NrRB")
    If Len(acc) = 0 Then acc = X(r, "Fa/Platnosc/RachunekBankowyFaktora/NrRB")
    v(12) = AccountCountry(acc)
    v(13) = AccountNumber(acc)

    '--- payment status ------------------------------------------------
    Dim gross As Double, paidMark As String, partMark As String
    Dim paidParts As Double, lastPartDate As String, s As String, nd As Object
    gross = Val(X(r, "Fa/P_15"))
    paidMark = X(r, "Fa/Platnosc/Zaplacono")
    partMark = X(r, "Fa/Platnosc/ZnacznikZaplatyCzesciowej")
    For Each nd In r.selectNodes(LP("Fa/Platnosc/ZaplataCzesciowa"))
        paidParts = paidParts + Val(X(nd, "KwotaZaplatyCzesciowej"))
        s = X(nd, "DataZaplatyCzesciowej")
        If s > lastPartDate Then lastPartDate = s      ' ISO dates compare as text
    Next nd
    v(14) = PaymentStatus(paidMark, partMark, gross, paidParts)
    If paidMark = "1" Then
        v(15) = ToDate(X(r, "Fa/Platnosc/DataZaplaty"))
    Else
        v(15) = ToDate(lastPartDate)
    End If

    v(16) = X(r, "Fa/KodWaluty")

    '--- amounts by VAT rate -------------------------------------------
    Dim tags As Variant, i As Long, sumN As Double, sumV As Double
    tags = Array("P_13_1", "P_14_1", "P_13_2", "P_14_2", "P_13_3", "P_14_3", _
                 "P_13_4", "P_14_4", "P_13_5", "P_14_5", "P_13_6_1", "P_13_6_2", _
                 "P_13_6_3", "P_13_7", "P_13_8", "P_13_9", "P_13_10", "P_13_11")
    For i = 0 To UBound(tags)
        s = X(r, "Fa/" & tags(i))
        v(21 + i) = ToNum(s)
        If Len(s) > 0 Then
            If Left$(tags(i), 4) = "P_13" Then sumN = sumN + Val(s) Else sumV = sumV + Val(s)
        End If
    Next i
    v(17) = Round(sumN, 2)
    v(18) = Round(sumV, 2)
    v(19) = ToNum(X(r, "Fa/P_15"))
    v(20) = AmountOutstanding(paidMark, partMark, gross, paidParts)

    Dim sumW As Double, anyW As Boolean
    For i = 1 To 5
        s = X(r, "Fa/P_14_" & i & "W")
        If Len(s) > 0 Then sumW = sumW + Val(s): anyW = True
    Next i
    If anyW Then v(39) = Round(sumW, 2) Else v(39) = Empty

    v(40) = ToDate(X(r, "Naglowek/DataWytworzeniaFa"))
    v(41) = fname
    v(42) = runTime

    '--- correction data -----------------------------------------------
    Dim e(1 To N_EXTRA) As Variant, cnt As Long
    If isCor Then
        e(1) = X(r, "Fa/PrzyczynaKorekty")
        e(2) = CorrectionTypeText(X(r, "Fa/TypKorekty"))
        Dim nos As String, dts As String, ksefs As String
        For Each nd In r.selectNodes(LP("Fa/DaneFaKorygowanej"))
            cnt = cnt + 1
            nos = JoinText(nos, X(nd, "NrFaKorygowanej"))
            dts = JoinText(dts, X(nd, "DataWystFaKorygowanej"))
            ksefs = JoinText(ksefs, X(nd, "NrKSeFFaKorygowanej"))
        Next nd
        e(3) = nos
        If cnt = 1 Then e(4) = ToDate(dts) Else e(4) = dts
        e(5) = ksefs
    End If

    '--- write to sheet ------------------------------------------------
    Dim ws As Worksheet, rowN As Long, c As Long
    If isCor Then Set ws = wsC Else Set ws = wsI
    rowN = ws.Cells(ws.Rows.Count, C_NUMBER).End(xlUp).Row + 1
    If rowN < 2 Then rowN = 2

    For c = 1 To N_BASE
        WriteCell ws.Cells(rowN, SheetCol(c, isCor)), v(c), BaseFormat(c)
    Next c
    If isCor Then
        For c = 1 To N_EXTRA
            If c = 4 Then
                WriteCell ws.Cells(rowN, EXTRA_POS + c), e(c), "yyyy-mm-dd"
            Else
                WriteCell ws.Cells(rowN, EXTRA_POS + c), e(c), "@"
            End If
        Next c
    End If

    keys.Add key, fname
    files(fname) = True

    If isCor Then
        WriteLog wsL, runTime, fname, "Added", "Corrections: " & invNo
        ProcessFile = 2
    Else
        WriteLog wsL, runTime, fname, "Added", "Invoices: " & invNo
        ProcessFile = 1
    End If
    Exit Function

Failed:
    WriteLog wsL, runTime, fname, "Error", Err.Description
    ProcessFile = 0
End Function

'------------------------------------------------------------------
'  Business rules
'------------------------------------------------------------------

' Settlement method: bank transfer / cash / offset / other
Private Function SettlementMethod(ByVal code As String, ByVal otherFlag As String, ByVal descr As String) As String
    Dim d As String
    d = LCase$(Trim$(descr))
    Select Case Trim$(code)
        Case "6": SettlementMethod = "bank transfer"
        Case "1": SettlementMethod = "cash"
        Case "2", "3", "4", "5", "7": SettlementMethod = "other"    ' card, voucher, cheque, credit, mobile
        Case Else
            If otherFlag = "1" Or Len(d) > 0 Then
                If HasAny(d, Array("kompens", "potr" & ChrW(261) & "c", "potrac", "offset", "netting", "wzajemn")) Then
                    SettlementMethod = "offset"
                ElseIf HasAny(d, Array("przelew", "transfer")) Then
                    SettlementMethod = "bank transfer"
                ElseIf HasAny(d, Array("got" & ChrW(243) & "wk", "gotowk", "cash")) Then
                    SettlementMethod = "cash"
                Else
                    SettlementMethod = "other"
                End If
            Else
                SettlementMethod = "bank transfer"   ' no payment method on the invoice
            End If
    End Select
End Function

' Payment status: paid / partial / pending
Private Function PaymentStatus(ByVal paidMark As String, ByVal partMark As String, _
        ByVal gross As Double, ByVal paidParts As Double) As String
    If paidMark = "1" Then
        PaymentStatus = "paid"
    ElseIf partMark = "1" Or partMark = "2" Or paidParts <> 0 Then
        If Abs(AmountOutstanding(paidMark, partMark, gross, paidParts)) < 0.005 Then
            PaymentStatus = "paid"
        Else
            PaymentStatus = "partial"
        End If
    Else
        PaymentStatus = "pending"
    End If
End Function

' Amount still to be paid (gross total minus partial payments)
Private Function AmountOutstanding(ByVal paidMark As String, ByVal partMark As String, _
        ByVal gross As Double, ByVal paidParts As Double) As Double
    If paidMark = "1" Or partMark = "2" Then          ' 2 = paid in full in instalments
        AmountOutstanding = 0
    ElseIf partMark = "1" Or paidParts <> 0 Then
        AmountOutstanding = Round(gross - paidParts, 2)
    Else
        AmountOutstanding = Round(gross, 2)
    End If
End Function

' Bank account: letters and digits only, upper case, optional "IBAN" prefix removed
Private Function CleanAccount(ByVal raw As String) As String
    Dim i As Long, ch As String, out As String
    raw = UCase$(Trim$(raw))
    For i = 1 To Len(raw)
        ch = Mid$(raw, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Then out = out & ch
    Next i
    If Left$(out, 4) = "IBAN" Then out = Mid$(out, 5)
    CleanAccount = out
End Function

' Country code of the account; PL when the number has no country prefix (or is missing)
Private Function AccountCountry(ByVal raw As String) As String
    Dim a As String
    a = CleanAccount(raw)
    If a Like "[A-Z][A-Z]*" Then AccountCountry = Left$(a, 2) Else AccountCountry = "PL"
End Function

' Account number without spaces and without the country prefix
Private Function AccountNumber(ByVal raw As String) As String
    Dim a As String
    a = CleanAccount(raw)
    If a Like "[A-Z][A-Z]*" Then AccountNumber = Mid$(a, 3) Else AccountNumber = a
End Function

Private Function DueDate(ByVal r As Object) As Variant
    Dim s As String
    s = X(r, "Fa/Platnosc/TerminPlatnosci/Termin")
    If Len(s) > 0 Then
        DueDate = ToDate(s)
        Exit Function
    End If
    ' descriptive due date (e.g. "14 dni od doreczenia")
    Dim n As Object, ch As Object, txt As String
    Set n = r.selectSingleNode(LP("Fa/Platnosc/TerminPlatnosci/TerminOpis"))
    If n Is Nothing Then
        DueDate = Empty
        Exit Function
    End If
    For Each ch In n.childNodes
        If Len(Trim$(ch.Text)) > 0 Then txt = txt & " " & Trim$(ch.Text)
    Next ch
    DueDate = Trim$(txt)
End Function

Private Function InvoiceTypeText(ByVal code As String) As String
    Select Case code
        Case "VAT": InvoiceTypeText = "VAT - standard"
        Case "KOR": InvoiceTypeText = "KOR - correction"
        Case "ZAL": InvoiceTypeText = "ZAL - advance"
        Case "ROZ": InvoiceTypeText = "ROZ - final settlement"
        Case "UPR": InvoiceTypeText = "UPR - simplified"
        Case "KOR_ZAL": InvoiceTypeText = "KOR_ZAL - advance correction"
        Case "KOR_ROZ": InvoiceTypeText = "KOR_ROZ - settlement correction"
        Case Else: InvoiceTypeText = code
    End Select
End Function

Private Function CorrectionTypeText(ByVal code As String) As String
    Select Case code
        Case "1": CorrectionTypeText = "1 - effective at original invoice date"
        Case "2": CorrectionTypeText = "2 - effective at correction issue date"
        Case "3": CorrectionTypeText = "3 - effective at another date"
        Case Else: CorrectionTypeText = code
    End Select
End Function

'------------------------------------------------------------------
'  Workbook layout (creates missing sheets)
'------------------------------------------------------------------
Private Sub EnsureWorkbookLayout()
    Dim ws As Worksheet, active As Object
    Set active = ActiveSheet

    If Not SheetExists(SH_START) Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        ws.Name = SH_START
        ws.Range("B2").Value = "KSeF Invoice Importer"
        ws.Range("B2").Font.Bold = True
        ws.Range("B2").Font.Size = 16
        ws.Range("B6").Value = "XML folder:"
        ws.Range("C6").NumberFormat = "@"
        ws.Range("C6").Value = DEFAULT_FOLDER
        ws.Range("B7").Value = "My NIP (optional):"
        ws.Range("C7").NumberFormat = "@"
        ws.Range("D7").Value = "If filled in, invoices where you are the seller are skipped."
        ws.Range("B9").Value = "Last import:"
        ws.Range("B6:B9").Font.Bold = True
        ws.Range("B6:B9").HorizontalAlignment = xlRight
        ws.Range("C6:C7").Interior.Color = RGB(255, 242, 204)
        ws.Columns("B").ColumnWidth = 24
        ws.Columns("C").ColumnWidth = 60
        Dim btn As Object
        Set btn = ws.Buttons.Add(ws.Range("C11").Left, ws.Range("C11").Top, 180, 36)
        btn.Caption = "Import invoices"
        btn.OnAction = "ImportKsefInvoices"
    End If

    EnsureDataSheet SH_INV, False
    EnsureDataSheet SH_COR, True

    If Not SheetExists(SH_LOG) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_LOG
        Dim h As Variant, i As Long
        h = LogHeaders()
        For i = 0 To UBound(h)
            ws.Cells(1, i + 1).Value = h(i)
        Next i
        StyleHeader ws, UBound(h) + 1
    End If

    active.Activate
End Sub

Private Sub EnsureDataSheet(ByVal sheetName As String, ByVal isCor As Boolean)
    If SheetExists(sheetName) Then Exit Sub
    Dim ws As Worksheet, h As Variant, c As Long, lastCol As Long
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = sheetName
    h = BaseHeaders()
    For c = 1 To N_BASE
        ws.Cells(1, SheetCol(c, isCor)).Value = h(c - 1)
    Next c
    lastCol = N_BASE
    If isCor Then
        h = ExtraHeaders()
        For c = 1 To N_EXTRA
            ws.Cells(1, EXTRA_POS + c).Value = h(c - 1)
        Next c
        lastCol = N_BASE + N_EXTRA
    End If
    StyleHeader ws, lastCol
End Sub

Private Sub StyleHeader(ByVal ws As Worksheet, ByVal lastCol As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120)
        .WrapText = True
        .VerticalAlignment = xlCenter
        .ColumnWidth = 14
        .AutoFilter
    End With
    ws.Rows(1).RowHeight = 45
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .SplitColumn = 1
        .SplitRow = 1
        .FreezePanes = True
    End With
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'------------------------------------------------------------------
'  Helpers
'------------------------------------------------------------------
Private Sub LoadExisting(ByVal ws As Worksheet, ByVal isCor As Boolean, ByVal keys As Object, ByVal files As Object)
    Dim last As Long, r As Long, invNo As String, k As String, fn As String
    last = ws.Cells(ws.Rows.Count, C_NUMBER).End(xlUp).Row
    For r = 2 To last
        invNo = Trim$(CStr(ws.Cells(r, C_NUMBER).Value))
        If Len(invNo) > 0 Then
            fn = Trim$(CStr(ws.Cells(r, SheetCol(C_FILE, isCor)).Value))
            k = MakeKey(CStr(ws.Cells(r, C_NIP).Value), invNo)
            If Not keys.Exists(k) Then keys.Add k, fn
            If Len(fn) > 0 Then files(fn) = True
        End If
    Next r
End Sub

Private Sub WriteCell(ByVal cell As Range, ByVal content As Variant, ByVal fmt As String)
    If IsEmpty(content) Then Exit Sub
    If VarType(content) = vbString Then
        If Len(content) = 0 Then Exit Sub
        cell.NumberFormat = "@"
    Else
        cell.NumberFormat = fmt
    End If
    cell.Value = content
End Sub

Private Function SheetCol(ByVal baseCol As Long, ByVal isCor As Boolean) As Long
    If isCor And baseCol > EXTRA_POS Then SheetCol = baseCol + N_EXTRA Else SheetCol = baseCol
End Function

Private Function BaseFormat(ByVal c As Long) As String
    Select Case c
        Case 6 To 10, 15: BaseFormat = "yyyy-mm-dd"
        Case 17 To 39: BaseFormat = "#,##0.00"
        Case 40, 42: BaseFormat = "yyyy-mm-dd hh:mm"
        Case Else: BaseFormat = "@"
    End Select
End Function

Private Sub WriteLog(ByVal ws As Worksheet, ByVal runTime As Date, ByVal fileName As String, _
        ByVal status As String, ByVal details As String)
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm"
    ws.Cells(r, 1).Value = runTime
    ws.Cells(r, 2).NumberFormat = "@"
    ws.Cells(r, 2).Value = fileName
    ws.Cells(r, 3).Value = status
    ws.Cells(r, 4).NumberFormat = "@"
    ws.Cells(r, 4).Value = details
End Sub

' Namespace-independent XPath: "Fa/P_2" -> *[local-name()='Fa']/*[local-name()='P_2']
Private Function LP(ByVal path As String) As String
    Dim parts() As String, i As Long, s As String
    parts = Split(path, "/")
    For i = 0 To UBound(parts)
        If i > 0 Then s = s & "/"
        s = s & "*[local-name()='" & parts(i) & "']"
    Next i
    LP = s
End Function

Private Function X(ByVal ctx As Object, ByVal path As String) As String
    Dim n As Object
    Set n = ctx.selectSingleNode(LP(path))
    If n Is Nothing Then X = "" Else X = Trim$(n.Text)
End Function

Private Function KsefFromFileName(ByVal fname As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\d{10}-\d{8}-[0-9A-F]{6,16}-[0-9A-F]{2}"
    re.IgnoreCase = True
    Set m = re.Execute(fname)
    If m.Count > 0 Then KsefFromFileName = UCase$(m(0).Value) Else KsefFromFileName = ""
End Function

Private Function MakeKey(ByVal nip As String, ByVal invNo As String) As String
    Dim n As String
    n = OnlyDigits(nip)
    If Len(n) = 0 Then n = UCase$(Trim$(nip))
    invNo = UCase$(Trim$(invNo))
    Do While InStr(invNo, "  ") > 0
        invNo = Replace(invNo, "  ", " ")
    Loop
    MakeKey = n & "|" & invNo
End Function

Private Function OnlyDigits(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    OnlyDigits = out
End Function

Private Function HasAny(ByVal haystack As String, ByVal words As Variant) As Boolean
    Dim i As Long
    For i = 0 To UBound(words)
        If InStr(1, haystack, words(i), vbBinaryCompare) > 0 Then HasAny = True: Exit Function
    Next i
End Function

Private Function ToDate(ByVal s As String) As Variant
    s = Trim$(s)
    If Len(s) < 10 Then ToDate = Empty: Exit Function
    If Mid$(s, 5, 1) <> "-" Or Mid$(s, 8, 1) <> "-" Then ToDate = s: Exit Function
    Dim d As Date
    d = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
    If Len(s) >= 19 And Mid$(s, 11, 1) = "T" Then
        d = d + TimeSerial(CInt(Mid$(s, 12, 2)), CInt(Mid$(s, 15, 2)), CInt(Mid$(s, 18, 2)))
    End If
    ToDate = d
End Function

Private Function ToNum(ByVal s As String) As Variant
    s = Trim$(s)
    If Len(s) = 0 Then ToNum = Empty Else ToNum = Val(s)
End Function

Private Function JoinText(ByVal a As String, ByVal b As String) As String
    If Len(b) = 0 Then
        JoinText = a
    ElseIf Len(a) = 0 Then
        JoinText = b
    Else
        JoinText = a & "; " & b
    End If
End Function

Private Function PickFolder(ByVal startPath As String) As String
    With Application.FileDialog(4)   ' msoFileDialogFolderPicker
        .Title = "Select the folder with KSeF XML files"
        If Len(startPath) > 0 Then .InitialFileName = startPath
        If .Show = -1 Then PickFolder = .SelectedItems(1) Else PickFolder = ""
    End With
End Function
