 strLink = "https://github.com/444c43/Excel-Tools-AddIn/blob/master/GFCTools.xla?raw=true"
	' Before using, change the below directory
 	 strSaveTo = "C:\Users\JaneDoe\Desktop\file.xla"
 	 
 	 WScript.Echo "GFC Tools Downloader"
 	 WScript.Echo "-------------"
 	 WScript.Echo "File To Download: " & strLink
 	 WScript.Echo "Save As: " & strSaveTo
	
     ' Create an HTTP object
     Set objHTTP = CreateObject( "WinHttp.WinHttpRequest.5.1" )
 
     ' Download the specified URL
     objHTTP.Open "GET", strLink, False
     ' Use HTTPREQUEST_SETCREDENTIALS_FOR_PROXY if user and password is for proxy, not for download the file.
     ' objHTTP.SetCredentials "User", "Password", HTTPREQUEST_SETCREDENTIALS_FOR_SERVER
     objHTTP.Send
     
          Set objFSO = CreateObject("Scripting.FileSystemObject")
	  If objFSO.FileExists(strSaveTo) Then
	  	objFSO.DeleteFile(strSaveTo)
	  End If
 
      If objHTTP.Status = 200 Then
    	Dim objStream
	    Set objStream = CreateObject("ADODB.Stream")
	    With objStream
		    .Type = 1 'adTypeBinary
		    .Open
		    .Write objHTTP.ResponseBody
		    .SaveToFile strSaveTo
		    .Close
	    End With
	    set objStream = Nothing
	  End If
	  
	  If objFSO.FileExists(strSaveTo) Then
	  	WScript.Echo "Download {" & strSaveTo & "} completed successfuly."
	  End If 