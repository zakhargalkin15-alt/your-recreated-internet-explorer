' Create an Internet Explorer object
Set IE = CreateObject("InternetExplorer.Application")

' Make the IE window visible
IE.Visible = True

' Navigate to a website
IE.Navigate "https://www.example.com"

' Optional: Wait for the page to load (ReadyState 4 indicates complete)
Do While IE.ReadyState <> 4 Or IE.Busy
    WScript.Sleep 100 ' Wait for 100 milliseconds
Loop

' You can now interact with the page, for example, by accessing elements
' Set objElement = IE.Document.getElementById("someElementId")
' If IsObject(objElement) Then
'     objElement.Click
' End If

' Clean up the object
Set IE = Nothing