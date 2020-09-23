Attribute VB_Name = "html_data"
'This is all the html data for the webserver. some is imbedded
' into the program, others are external files.

Public Function html_404error()

'This is the 404 error page that is showed when the page
' that is requested isn't found.  I embedded this into the
' program so that it is always able to be displayed.



Dim x As String
x = ""

x = x & "<html>" & vbCrLf
x = x & "" & vbCrLf
x = x & "<head>" & vbCrLf
x = x & "<style>" & vbCrLf
x = x & "a:link          {font:8pt/11pt verdana; color:red; text-decoration:none}" & vbCrLf
x = x & "a:visited       {font:8pt/11pt verdana; color:red}" & vbCrLf
x = x & "a:hover          {font:8pt/11pt verdana; color:red; text-decoration:underline}" & vbCrLf
x = x & "</style>" & vbCrLf
x = x & "<meta HTTP-EQUIV=""Content-Type"" Content=""text-html; charset=Windows-1252"">" & vbCrLf
x = x & "<title>HTTP 404 Not Found</title>" & vbCrLf
x = x & "</head>" & vbCrLf
x = x & "" & vbCrLf
x = x & "<body bgcolor=""#FFFFFF"">" & vbCrLf
x = x & "<p><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""2""><b><font color=""#FF0000"">The" & vbCrLf
x = x & "  page cannot be found </font></b></font></p>" & vbCrLf
x = x & "<p>&nbsp;</p>" & vbCrLf
x = x & "<p><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1"">The page you are" & vbCrLf
x = x & "  looking for might have been removed, had its name changed, or is temporarily" & vbCrLf
x = x & "  unavailable. </font></p>" & vbCrLf
x = x & "<p align=""center"">&nbsp;</p>" & vbCrLf
x = x & "<p align=""center""><font size=""1"" face=""Verdana, Arial, Helvetica, sans-serif"" color=""#0000FF""><i><font color=""#000000"">HTTP" & vbCrLf
x = x & "  404 - File not found</font></i></font></p>" & vbCrLf
x = x & "</body>" & vbCrLf
x = x & "</html>" & vbCrLf & vbCrLf & vbCrLf
html_404error = x
End Function


