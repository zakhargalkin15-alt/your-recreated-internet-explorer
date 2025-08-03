# Recreated Internet Explorer (VBScript)

Internet Explorer is a classic web browser, but sadly it has been discontinued by Microsoft. However, with this project, you can relive the experience!

This repository features a fully recreated version of Internet Explorer using VBScript. This script can open real websites and offers an interface that closely resembles the original Internet Explorer—bringing nostalgia and functionality together!

## Features

- Runs on modern Windows (Windows 10 and 11), as well as earlier versions (Windows 98 to Windows 8)
- Opens real websites
- Faithful Internet Explorer look and feel

## Is this a virus?

**Nope, it's not a virus!**  
This VBScript is safe to use. You can review the code yourself (see the script below). It simply creates an Internet Explorer window, navigates to a website, and optionally interacts with elements on the page. There are no malicious instructions—its only purpose is to recreate the Internet Explorer experience and allow you to browse the web. If you have any concerns, feel free to check the script before running it!

## Example Code

```vbscript
' Create an Internet Explorer object
Set IE = CreateObject("InternetExplorer.Application")

' Make the IE window visible
IE.Visible = True

' Navigate to a website
IE.Navigate "https://www.google.com"

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
```

## How to Use

1. Download the VBScript file from this repository.
2. Double-click the script to launch your recreated Internet Explorer.
3. Surf the web—just like old times!

## Why?

Internet Explorer has been discontinued, but many users feel nostalgic for its classic interface and experience. This project gives you a way to bring Internet Explorer back to life on modern systems.

---

Enjoy your trip down memory lane!
