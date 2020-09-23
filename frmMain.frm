VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "web server"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   4755
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "stop server"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start server"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock sckWS 
      Index           =   0
      Left            =   3360
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private requestedPage As String
' Web Server code .3a - last updated 12/19/99
'
' E-mail me at nymainst@nais.com
'
'This code is an example of how to create
' a basic webserver.  connect to
' http://your.ip  and it will send the requested
' page. (Requesting "/" will send the index.htm
' page.  if it is not found it sends the 404 error :)
' you could do plenty with this, i.e. when someone
' sends a form ('POST') have it get the data, etc.
' you could create a program that runs a guestbook
' off of your computer. :)
'
'There is room for improvement on this, i'm sure.  If you
' improve it, please let me know and send me a copy :)
'
'v .2a fixed:  some loading problems
'v .3a fixed: fixed the $ip linking problem, everything works now. :)
' also fixed: i made the project compatible with vb5 by removing the replace
' function.
'  Creating a link to another page
' To link to another page, link to 'http://$ip/page_name.html'
'The webserver will replace $ip with your ip. The webserver also
'supports files in other directories in the html dir.  You could link
'to test.html in the directory 'links' by linking to:
' http://$ip/links/test.html
'
'                                  PLEASE REPORT ANY BUGS!
'                                   pat - nymainst@nais.com


Private Sub Command1_Click()
load_defaults 'Start up the server.
End Sub

Private Sub Command2_Click()
Call stop_server
End Sub

Private Sub Form_Load()
ttlConnections = 0 'Set the ttlConnections varible
'to zero. :)
End Sub


Private Sub sckWS_Close(Index As Integer)

End Sub

Private Sub sckWS_ConnectionRequest(Index As Integer, ByVal requestID As Long)
   If Index = 0 Then
      ttlConnections = ttlConnections + 1  'add 1 to the total # of connections
      numConnections = numConnections + 1 'number of connected clients + 1
    
      If numConnections = maxConnections Then GoTo done 'if we've reached the max # of connections, exit sub.
      Load sckWS(ttlConnections) 'load a new instance of sckWS.
      sckWS(ttlConnections).LocalPort = 0 'set its local port to 0
      sckWS(ttlConnections).Accept requestID 'Accept the connection request.
      
      List1.AddItem sckWS(ttlConnections).RemoteHostIP & " connected" ' just add this to the listbox.
      
StartOver:
      
      DoEvents 'DoEvents so it doesn't freeze while we wait.
      If requestedPage$ = "" Then GoTo StartOver 'if we havent gotten the page request yet, go back to startOver.
      List1.AddItem "requested page: " & requestedPage$
      If requestedPage$ = "/" Then requestedPage$ = htmlIndexPage$ ' if the page '/' was requested, set requested page to the index html page.
      
      
      If FileExists(htmlPageDir & "\" & requestedPage$) Then 'if the requested page exists, then..
      
      htmldata$ = text_read(htmlPageDir & "\" & requestedPage$) 'This reads the file and stores it's contents in htmldata$
      htmldata$ = ReplaceStr(htmldata$, "$ip", sckWS(0).LocalIP) 'Oops, i didn't use the replace function right.  Now it's fixed at replaces $ip with your IP.
      sckWS(ttlConnections).SendData htmldata$ & vbCrLf  'open and read the requested HTML page.
      Else 'if it doesn't exist, then...
       
       If requestedPage$ = htmlIndexPage$ Then 'If the requested page is the index page and it doesn't exist, print this.
       sckWS(ttlConnections).SendData "<html><font face=""Verdana, Arial, Helvetica, sans-serif"" size=""1""><b>Please create an index html page.  It was not found.</font></html>" & vbCrLf ' If the requested page is the index and it doesn't exist, it tells you.
       requestedPage$ = ""
       End If
      
      requestedPage$ = "/a"
      sckWS(ttlConnections).SendData html_404$ & vbCrLf 'Send the 404 Error HTML
      End If
   End If
   
done:
      numConnections = numConnections - 1 'number of connections at the moment - 1
End Sub

Private Sub sckWS_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strdata As String

sckWS(Index).GetData strdata$ 'Get any data sent to us


If Mid$(strdata$, 1, 3) = "GET" Then 'If it is trying to get a site, find out
findget = InStr(strdata$, "GET ")      ' the site they want then set requestedPage$
spc2 = InStr(findget + 5, strdata$, " ") ' to it.
pagetoget$ = Mid$(strdata$, findget + 4, spc2 - (findget + 4))
requestedPage$ = pagetoget$
End If
End Sub


Private Sub sckWS_SendComplete(Index As Integer)
'This was a bug that was fixed from v.2a.


If requestedPage$ <> "" Then 'f the requested page doesn't = nothing then...
      requestedPage$ = "" 'clear the requestedPage varible.
      sckWS(ttlConnections).Close 'Close the connection.
End If
End Sub


