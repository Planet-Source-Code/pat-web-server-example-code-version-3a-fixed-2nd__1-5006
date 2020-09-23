Attribute VB_Name = "http_cmds"

Global http_port As Long            'Port we will listen on (http port is 80)
Global ttlConnections As Long    'Total # of connections we have had.
Global maxConnections As Long 'Max. # of connections allowed.
Global numConnections As Long  'Number of connections at the time.
Global htmlPageDir As String     'The directory where the
                                               'HTML pages are being stored.

Global html_404 As String         'This is the 404 error page HTML.
Global htmlIndexPage As String

Sub load_defaults()
Dim tport As String

'This simply loads up the server.
'Example use:  Call load_defaults

http_port = 80 'This is the port we are listening on :)
maxConnections = 100 'Maximium number of connections
'you can have at one time.  After we send the html data,
'the connection is CLOSED.  So, you probably could set this
'to 5 and it would work fine. :)

ttlConnections = 0 'Total number of connections = 0
numConnections = 0 'Total number of connections at the moment is zero :)

htmlPageDir$ = App.Path  'Set the html directory to wherever the app is.
htmlIndexPage$ = "index.html"

html_404$ = html_404error() 'Set the html 404 error page HTML.
                           'This could also be htmlPageDir$ & "\404.html"

tport$ = ""
If http_port = 80 Then tport$ = "" Else: tport$ = ":" & http_port ' this makes the
                                'string tport ':port'.  the format is http://ip:port.
                                 'if the port is 80 you can just leave it out.(http://ip)

With frmMain
    .sckWS(0).Close
    .sckWS(0).LocalPort = http_port
    .sckWS(0).Listen
    
    .Text1.Text = "http://" & .sckWS(0).LocalIP & tport$
    .Command1.Enabled = False
    .Command2.Enabled = True
End With

End Sub


Public Sub retrieveHeader(tPage As String, sckWSC)
'This won't be used since this is a web server, not client.
' but i thought i might as well add it. :)

'This is the data sent to a server when the client is requesting
'a page.

'tPage$ is the requested page, e.g., about.html
'sckWSC is a MS winsock control. e.g., Winsock1

'Example use:  Call retrieveHeader(downloads.html, Winsock1)

With sckWSC
    .SendData "GET /" & tPage$ & " HTTP/1.1" & vbCrLf
    .SendData "Accept: text/plain" & vbCrLf
    .SendData "Accept-Language: en-us" & vbCrLf
    .SendData "Accept-Encoding: gzip, deflate" & vbCrLf
    .SendData "User-Agent: Mozilla/4.0 (compatible; MSIE 5.0; Windows 98; DigExt)" & vbCrLf
    .SendData "Host: " & sckWSC.LocalIP & vbCrLf
    .SendData "Connection: Keep-Alive" & vbCrLf & vbCrLf
End With

End Sub

Public Sub stop_server()
'This sub shuts down the server

With frmMain
.Command1.Enabled = True  'Enableds the start button
.Command2.Enabled = False 'Disables the stop button
.List1.Clear 'Clears the log
.Text1.Text = ""
.sckWS(0).Close 'Closes the port
End With

Call unloadControls
Call unset_vars

End Sub

Public Sub unloadControls()
'This unloads all the winsock controls we loaded

With frmMain
For i = 1 To ttlConnections
Unload .sckWS(i)
Next i
End With

End Sub


Public Sub unset_vars()
'This clears out all of the varibles

http_port = 0
ttlConnections = 0
maxConnections = 0
numConnections = 0
htmlPageDir = 0
html_404 = ""
htmlIndexPage = ""
End Sub


