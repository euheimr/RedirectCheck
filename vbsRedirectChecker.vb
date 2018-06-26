
Function main()
    Dim userUrl, mbTitle, mbDefault, urlIsValid, inputUrl
    mbTitle = "Redirection checker"

    Try
        'grab info from clipboard
        userUrl = %windir%\system32\clip.exe
	    'make sure it isn't empty
	    'if the length of userUrl is 0, quit the script with a message
	    If Len(userUrl) = 0 Then
            MsgBox("Copy a link first before executing this script! You must include 'http:' OR 'https:'", 1, "ERROR")

        End If
        'check it to make sure it's a link
        'this basically uses InStr to see if http: or https: occurs in userUrl's input

        'isValid acts like a boolean. validateUrl will return 1 for GOOD LINK 
        ' returns 0 for BAD / ERROR
        isValid = validateUrl(userURL)

        'if "http:" or "https:" isn't in the userUrl, the script will quit
        If urlVar0 = 1 Or urlVar1 = 1 Then
            InputBox("Pasted from your clipboard:", mbTitle, userUrl)
        Else

            InputBox("Enter or paste your url here:", mbTitle, inputUrl)
        End If
    End Try
End Function


'call this to recursively check url redirects		
Function validateUrl(userUrl)
	On Error Resume Next
	Const WHR_EnableRedirects = 6
	Dim oHttp,Target
	Set oHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
	oHttp.Option(WHR_EnableRedirects) = False
	oHttp.Open "HEAD", userUrl, False
	oHttp.send
	If Err.Number = 0 Then
	    Target = "There is no redirection " &_
            oHttp.Status & " " & oHttp.statusText & vbcrlf &_
            "for this URL : " & chr(34) & strUrl & chr(34)
			return 0
        ElseIf oHttp.Status = 301 Or oHttp.Status = 302 Then
            Target = "This URL is redirected to : " & vbCrlf &_
            chr(34) & oHttp.getResponseHeader("Location") & chr(34)
			return 1
		End If
    Else
        GetHeaderLocation = "Error " & Err.Number & vbCrlf &_
        Err.Source & " " & Err.Description
    End If
    GetHeaderLocation = Target
End Function
		
	

%>


strName = Request.ServerVariables("SERVER_NAME")

sScriptLoc = Request.ServerVariables(strName)
'variables for http request handling
'Dim strProtocol, strDomain, strPath, strQueryURL, strFullURL

'vars for getting url from user
'Dim userURL
'userUrl = document.referrer
'get user input using a dialog box

'assign the user input info to


'If lcase(Request.ServerVariables("HTTPS")) = "on" Then
'    strProtocol = "https"
'	Else
'	strProtocol = "http"
'End If

'strDomain = 













