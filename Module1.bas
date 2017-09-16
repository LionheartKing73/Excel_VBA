Attribute VB_Name = "Module1"

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
    IE.Navigate "http://www.excely.com/"
 
    ' Statusbar
    Application.StatusBar = "www.excely.com is loading. Please wait..."
 
    ' Wait while IE loading...
    Do While IE.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop
 
    ' Find 2 input tags:
    '   1. Text field
    '   <input type="text" class="textfield" name="s" size="24" value="" />
    '
    '   2. Button
    '   <input type="submit" class="button" value="" />
    
    Application.StatusBar = "Search form submission. Please wait..."
 
    Set objCollection = IE.document.getElementsByTagName("input")
 
    i = 0
    While i < objCollection.Length
        If objCollection(i).Name = "s" Then
 
            ' Set text for search
            objCollection(i).Value = "excel vba"
 
        Else
            If objCollection(i).Type = "submit" And _
               objCollection(i).Name = "" Then
 
                ' "Search" button is found
                Set objElement = objCollection(i)
 
            End If
        End If
        i = i + 1
    Wend
    objElement.Click    ' click button to search
    
    ' Wait while IE re-loading...
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
