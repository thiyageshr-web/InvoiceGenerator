Sub Generate_Invoice_PDF()

    Dim ws As Worksheet
    Dim InvoiceNo As String
    Dim SavePath As String
    Dim FileName As String
    Dim FullPath As String
    
    '--- Set your invoice sheet name ---
    Set ws = ThisWorkbook.Sheets("INVOICE")
    
    '--- Get invoice number from a cell (e.g. Q9) ---
    InvoiceNo = ws.Range("Q9").Value
    
    '--- Folder where PDF will be saved ---
    SavePath = ThisWorkbook.Path & "\Invoices\"   ' Creates "Invoices" folder next to your file
    
    '--- Create folder if not existing ---
    If Dir(SavePath, vbDirectory) = "" Then MkDir SavePath
    
    '--- File name ---
    FileName = "Invoice_" & InvoiceNo & "_" & Format(Date, "ddmmyy") & ".pdf"
    
    '--- Full file path ---
    FullPath = SavePath & FileName
    
    '--- Export as PDF ---
    ws.ExportAsFixedFormat Type:=xlTypePDF, _
        FileName:=FullPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True   'change to False if you don’t want to open automatically
    
    '--- Confirmation message ---
    MsgBox "Invoice saved as PDF:" & vbCrLf & FullPath, vbInformation, "PDF Generated Successfully"

End Sub


