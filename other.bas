Attribute VB_Name = "other"
'Other functions and subs needed :)

Public Function ReplaceStr(ByVal strMain As String, strFind As String, strReplace As String) As String
'Thsi is the same thing as the Replace function in vb6.  I added this
'for those of you using vb5.  This was NOT written by me, it was written by
' someone named 'dos'.  He's a great programmer, visit his webpage @
' http://hider.com/dos

    Dim lngSpot As Long, lngNewSpot As Long, strLeft As String
    Dim strRight As String, strNew As String
    lngSpot& = InStr(LCase(strMain$), LCase(strFind$))
    lngNewSpot& = lngSpot&
    Do
        If lngNewSpot& > 0& Then
            strLeft$ = Left(strMain$, lngNewSpot& - 1)
            If lngSpot& + Len(strFind$) <= Len(strMain$) Then
                strRight$ = Right(strMain$, Len(strMain$) - lngNewSpot& - Len(strFind$) + 1)
            Else
                strRight = ""
            End If
            strNew$ = strLeft$ & strReplace$ & strRight$
            strMain$ = strNew$
        Else
            strNew$ = strMain$
        End If
        lngSpot& = lngNewSpot& + Len(strReplace$)
        If lngSpot& > 0 Then
            lngNewSpot& = InStr(lngSpot&, LCase(strMain$), LCase(strFind$))
        End If
    Loop Until lngNewSpot& < 1
    ReplaceStr$ = strNew$
End Function
Public Function text_read(filename)
'This function reads a file and spits out the text in it.

Dim f
Dim textda
Dim cha

On Error Resume Next
f = FreeFile
textda = ""
If FileExists(filename) Then
    If Len(filename) Then
        Open filename For Input As #f   ' Open file.
        Do While Not EOF(f)
            cha = Input(1, #f) ' Get one character.
             textda = "" & textda & cha
        Loop    ' Loop if not end of file.
        Close #f
    End If
text_read = textda
Else
text_read = ""
End If

End Function
Public Function FileExists(ByVal sFileName As String) As Integer
'Checks if the given file exists.

Dim i As Integer
On Error Resume Next

    i = Len(Dir$(sFileName))
    
    If Err Or i = 0 Then
        FileExists = False
        Else
            FileExists = True
    End If
End Function
Public Sub timeout(ByVal nSecond As Single)
'Pauses for x seconds.

   Dim t0 As Single
   t0 = Timer
   Do While Timer - t0 < nSecond
      Dim dummy As Integer

      dummy = DoEvents()
      If Timer < t0 Then
         t0 = t0 - CLng(24) * CLng(60) * CLng(60)
      End If
   Loop

End Sub
