Private Function ApplyBackgroundToPDF(BasePDF As String, BackgroundPDF As String)
    Dim pdDoc As Acrobat.CAcroPDDoc
    Dim pdTemplate As Acrobat.CAcroPDDoc
    Dim template As Variant
    Dim lngPage As Long
    
    'Open base document
        Set pdDoc = CreateObject("AcroExch.PDDoc")
        pdDoc.Open BasePDF
        DoEvents
    
    'Open background document
        Set pdTemplate = CreateObject("AcroExch.PDDoc")
        pdTemplate.Open BackgroundPDF
        DoEvents
    
    'Add background document to base document
        pdDoc.InsertPages pdDoc.GetNumPages - 1, pdTemplate, 0, 1, 0
    
    'Create a template from the inserted background document
        Set template = pdDoc.GetJSObject.CreateTemplate("background", pdDoc.GetNumPages - 1)
    
    'Place the template as a background to all pages
        For lngPage = 0 To pdDoc.GetNumPages - 2
            template.Spawn lngPage, True, True
        Next
    
    'Delete last page (used for template creation purposes only)
        pdDoc.DeletePages pdDoc.GetNumPages - 1, pdDoc.GetNumPages - 1
    
    'Save
        pdDoc.Save 1, BasePDF
    
    'Close & Destroy Objects
        pdDoc.Close
        Set pdDoc = Nothing
        
        pdTemplate.Close
        Set pdTemplate = Nothing
End Function
