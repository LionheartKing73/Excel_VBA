Sub IE_Autiomation()
  Dim i As Long
    Dim IE As Object
    Dim objElement As Object
    Dim objCollection As Object
 
    ' Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
 
    ' You can uncoment Next line To see form results
    IE.Visible = False
 
    ' Send the form data To URL As POST binary request
    IE.Navigate "https://web.edumaximizer.com/14196.cmp?IPAddress=50.253.143.46"
 
    ' Statusbar
    Application.StatusBar = "www.excely.com is loading. Please wait..."
    
 
    ' Wait while IE loading...
    Do While IE.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop
    
      
    newInvoiceScript = "var scope = angular.element(document.getElementsByClassName('form-group')).scope();scope.lead={FirstName:'" + Sheet1.Cells(2, 1) + "',LastName:'" + Sheet1.Cells(2, 2) + "',Address1:'" + Sheet1.Cells(2, 3) + "',Address2:'" + Sheet1.Cells(2, 3) + "',City:'" + Sheet1.Cells(2, 4) + "',State:'" + Sheet1.Cells(2, 5) + "'};"
    IE.document.parentWindow.execScript newInvoiceScript
    

    Do While IE.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop
 
    ' Show IE
    IE.Visible = True
 
    ' Clean up
    Set IE = Nothing
    Set objElement = Nothing
    Set objCollection = Nothing
 
    Application.StatusBar = ""

End Sub
