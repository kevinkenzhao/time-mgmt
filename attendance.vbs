Option Explicit

Dim qrcode

Do While True
    qrcode = GetPassword( "Scan your QR code | 請掃描您的二維碼: " )
    if (qrcode = "nycs") then
        Exit Do
    else
        OpenWithChrome qrcode
	exitChrome()
    end if
Loop

Function GetPassword( myPrompt )
	Dim blnFavoritesBar, blnLinksExplorer, objIE, strHTML, strRegValFB, strRegValLE, wshShell

	blnFavoritesBar  = False
	strRegValFB = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\MINIE\LinksBandEnabled"
	blnLinksExplorer = False
	strRegValLE = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\LinksExplorer\Docked"

	Set wshShell = CreateObject( "WScript.Shell" )

	On Error Resume Next
	' Temporarily hide IE's Favorites Bar if it is visible
	If wshShell.RegRead( strRegValFB ) = 1 Then
		blnFavoritesBar = True
		wshShell.RegWrite strRegValFB, 0, "REG_DWORD"
	End If
	If wshShell.RegRead( strRegValLE ) = 1 Then
		blnLinksExplorer = True
		wshShell.RegWrite strRegValLE, 0, "REG_DWORD"
	End If
	On Error Goto 0
	
	Set objIE = CreateObject( "InternetExplorer.Application" )
	objIE.Visible        = True
	CreateObject("WScript.Shell").AppActivate "Internet Explorer"
	objIE.Navigate "about:blank"
	objIE.Document.title = "New York Chinese School TA Sign-In"
	objIE.AddressBar     = False
	objIE.Resizable      = False
	objIE.StatusBar      = False
	objIE.ToolBar        = False
	objIE.Width          = 1200
	objIE.Height         = 800
	objIE.FullScreen     = True
	objIE.document.focus()
	' Center the dialog window on the screen
	With objIE.Document.parentWindow.screen
		objIE.Left = (.availWidth  - objIE.Width ) \ 2
		objIE.Top  = (.availheight - objIE.Height) \ 2
	End With
	' Wait till IE is ready
	Do While objIE.Busy
		WScript.Sleep 200
	Loop

	' Insert the HTML code to prompt for a password
	strHTML = "<div style=""text-align: center;"">" _
		& "<h1 style = font-family:Arial,Helvetica,sans-serif;font-size:32px;font-style:bold;>" & "New York Chinese School TA Sign-In | 紐約華僑學校打卡系統" & "</h1>" _
		& "<p style = font-family:Arial,Helvetica,sans-serif;font-size:22px;font-style:bold;>" & myPrompt & "</p>" _
	        & "<p><input type=""password"" size=""20"" id=""Password"" oninput=""document.all.OKButton.click()""" _
	        & """if(event.keyCode==13){document.all.OKButton.click();}"" /></p>" _
	        & "<p><input type=""hidden"" id=""OK"" name=""OK"" value=""0"" />" _
	        & "<input type=""submit"" value="" OK | 輸入 "" id=""OKButton"" " _
	        & "onclick=""document.all.OK.value=1"" /></p>" _
		& "<img src='C:\Users\Kevin\Desktop\2018 NYCS LOGO.png'>" _
	        & "</div>"
	objIE.Document.body.innerHTML = strHTML

	' Hide the scrollbars
	objIE.Document.body.style.overflow = "auto"
	objIE.Document.body.style.backgroundColor = "#FFFFF0"
	' Make the window visible
	objIE.Visible = True
	' Set focus on password input field
	objIE.Document.all.Password.focus

	' Wait till the OK button has been clicked
	On Error Resume Next
	Do While objIE.Document.all.OK.value = 0 
		WScript.Sleep 200
		If Err Then	' User closes Window
			GetPassword = ""
			objIE.Quit
			Set objIE = Nothing
			' Restore IE's Favorites Bar if applicable
			If blnFavoritesBar Then wshShell.RegWrite strRegValFB, 1, "REG_DWORD"
			' Restore IE's Links Explorer if applicable
			If blnLinksExplorer Then wshShell.RegWrite strRegValLE, 1, "REG_DWORD"

			Exit Function
		End if
	Loop
	On Error Goto 0

	' Read the password from the dialog window
	GetPassword = objIE.Document.all.Password.value

	' Terminate the IE object
	objIE.Quit
	Set objIE = Nothing

	On Error Resume Next
	' Restore IE's Favorites Bar if applicable
	If blnFavoritesBar Then wshShell.RegWrite strRegValFB, 1, "REG_DWORD"
	' Restore IE's Links Explorer if applicable
	If blnLinksExplorer Then wshShell.RegWrite strRegValLE, 1, "REG_DWORD"
	On Error Goto 0

	Set wshShell = Nothing
End Function

function readFromRegistry (strRegistryKey, strDefault)
    Dim WSHShell, value
    On Error Resume Next
    Set WSHShell = CreateObject ("WScript.Shell")
    value = WSHShell.RegRead (strRegistryKey)

    if err.number <> 0 then
        readFromRegistry= strDefault
    else
        readFromRegistry=value
    end if

    set WSHShell = nothing
end function

function OpenWithChrome(strURL)
    Dim strChrome
    Dim WShellChrome
    strChrome = readFromRegistry ( "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe\Path", "")

    if (strChrome = "") then
        strChrome = "chrome.exe"
    else
        strChrome = strChrome & "\chrome.exe"
    end if

    if (strURL = "") then
        Exit Function
    else
        Set WShellChrome = CreateObject("WScript.Shell")
    	strChrome = """" & strChrome & """" & " " & "--kiosk " & strURL
   	WShellChrome.Run strChrome, 1, false
    end if

end function

Function exitChrome()
	WScript.Sleep .5*60*1000 'convert from minutes to milliseconds

	Dim objExec, objShell

	Set objShell = CreateObject("WScript.Shell")
	Set objExec = objShell.Exec("tasklist /fi " & Chr(34) & "imagename eq chrome.exe" & Chr(34))
	If Not InStr(1, objExec.StdOut.ReadAll(), "INFO: No tasks", vbTextCompare) Then
	    objShell.Run "taskkill /f /t /im chrome.exe", 0, True
	End If

	Set objExec = Nothing : Set objShell = Nothing

End Function

WScript.Quit
