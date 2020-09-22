Attribute VB_Name = "Rtf2Html"

Option Explicit
'----------------------------------------------------------------------------------
Dim sCodes() As String                      'Code array for end codes for bold,
                                            'italic, and underline. as well as
                                            'bullet

Dim sRTFLine As String                      'Holds the current RTF line we are
                                            'working on
Dim sRTFWord As String                      'holds the current RTF command word
                                            'we are working on.
Dim sRTFLeft As String                      'holds the left character of the
                                            'command we are working on.
Dim sRTFRight As String                     'holds any text that may follow the
                                            'rtf command.
Dim sHoldHtml As String                     'Holds the html text as we build the
                                            'current line.
'----------------------------------------------------------------------------------

Dim bBullet As Boolean                      'Boolean used in processing unordered
                                            'list. (Bullet). If this is true, we
                                            'are processing a bulleted list.
'-----------------------------------------------------------------------------------

Dim iBracketLocation As Integer             'Counter for brackets
                                            'header is discarded by finding double
                                            'brackets }}
Dim iFindBackSlash As Integer               'Holds location of last backslash
                                            'each RTF command begins with a \ and
                                            'ends with another \ or a space
Dim iFindNextSlash As Integer               'See if there is another backslash
Dim iFindRightBracket As Integer            'Holds loc of last right bracket
                                            'used to discard RTF header
Dim iFindSpace As Integer                   'if an RTF command doesn't end with
                                            'another \ then it ends with a space,
                                            'and there is ascii text following it.
Dim iFindEOL As Integer                     'Find the end of the current line we
                                            'want to process
Dim iCodeCounter As Integer                 'This counter tells the routine popcode
                                            'when to remove code from sCodes() and
                                            'add it to sHoldHtml
Dim I As Integer                            'Used in For/Next loop

'-----------------------------------------------------------------------------------
' This is the main function that is called to convert RTF to HTML
' You pass it the RTF text from your Richtext box, and it will return HTML
' In a future addition I may want to add options, so I've set the function up
' to accept optional string information that will be used to flag options.
'-----------------------------------------------------------------------------------

Function RTFtoHTML(sRTFText As String, Optional sOptions As String) As String


    Dim sRTFRemaining As String             'This is what remains of the
                                            'RTF text being processed
                                            
    sRTFRemaining = sRTFText                'get the rtftext to work with
    subClearScodes                          'Redimension sCodes to clear it of
                                            'any previous information.
    
'----------------------------------------------------------------------------------
' Strip RTF header from the text
' Because we only want to process Bold, Underline, Italic, and Bullet, we can
' savely discard all the header information.
'----------------------------------------------------------------------------------
    iFindRightBracket = 1
    While iFindRightBracket > 0
        iFindRightBracket = InStr(iFindRightBracket + 1, sRTFRemaining, "}}")
        If iFindRightBracket > 0 Then iBracketLocation = iFindRightBracket
    Wend
'-----------------------------------------------------------------------------------
' Put remaining text after header removed back into sRTFRemaining
' A little known fact about the Mid function; if you don't give it a length to
' process, then it assumes the length of the string from the starting point to the
' end of the string.
'-----------------------------------------------------------------------------------
    sRTFRemaining = Mid(sRTFRemaining, iBracketLocation + 2)
    
        While sRTFRemaining <> ""
'-----------------------------------------------------------------------------------
' I like to name my functions and subs using fun or sub for the first three letters
' of the routine. This makes it easier in the code to see if you're calling a sub
' or a function.
'-----------------------------------------------------------------------------------
            sRTFLine = funGetNextLine(sRTFRemaining)
            sRTFRemaining = Mid(sRTFRemaining, Len(sRTFLine) + 1)
            subProcessRTFLine
        Wend
        RTFtoHTML = sHoldHtml
        
End Function

'-----------------------------------------------------------------------------------
' This subroutine takes the RTF line that we are working on, and looks for RTF
' commands. RTF commands always start with a backslash, and will end with a
' backslash or a space
'-----------------------------------------------------------------------------------

Sub subProcessRTFLine()

        iFindBackSlash = 1
        

        While sRTFLine > ""
        
'-----------------------------------------------------------------------------------
' Check to see if sRTFline has a backslash
'-----------------------------------------------------------------------------------
            While iFindBackSlash > 0
       
                iFindBackSlash = InStr(sRTFLine, "\")
                If iFindBackSlash > 1 Then
                    sHoldHtml = sHoldHtml & Left(sRTFLine, iFindBackSlash - 1)
                    
                End If
                Select Case iFindBackSlash
                    
                    
'-----------------------------------------------------------------------------------
' Yes, this RTF line has at least one backslash. Now check for a second one
'-----------------------------------------------------------------------------------
                    Case Is > 0
                        iFindNextSlash = InStr(iFindBackSlash + 1, sRTFLine, "\")
                        
                        Select Case iFindNextSlash
                        
'-----------------------------------------------------------------------------------
' Found a second backslash, store the RTF command in sRTFWord
' Remove that word from the beginning of sRTFLine
'-----------------------------------------------------------------------------------
                            Case Is > 0
                                sRTFWord = Left(sRTFLine, iFindNextSlash - 1)
                                sRTFLine = Mid(sRTFLine, Len(sRTFWord) + 1)
                                iFindSpace = InStr(sRTFWord, " ")
                                
                               If iFindSpace > 0 Then
                                    sRTFRight = Mid(sRTFWord, iFindSpace + 1)
                                    sRTFWord = Left(sRTFWord, iFindSpace - 1)
                               End If
                           
                                
                            Case 0
'-----------------------------------------------------------------------------------
' No second back slash was found. We now need to see if there is a space.
'-----------------------------------------------------------------------------------
                                iFindNextSlash = InStr(iFindBackSlash + 1, sRTFLine, " ")
                                Select Case iFindNextSlash
                                    Case Is > 0
                                        sRTFWord = Left(sRTFLine, iFindNextSlash - 1)
                                        sRTFLine = Mid(sRTFLine, Len(sRTFWord) + 1)
                                        
                                    Case 0
                                    
                                End Select
                               
                        End Select
                        
                    Case 0
'-----------------------------------------------------------------------------------
' No backslash found at all. We've reached the end of SRTFLine
'-----------------------------------------------------------------------------------
                    
                        sRTFWord = sRTFLine
                        sRTFLine = ""
                        
                End Select
                
                subProcessRTFWord
            Wend
    Wend
                
   ' subClearScodes
                                                
End Sub
Sub subProcessRTFWord()
'-----------------------------------------------------------------------------------
'Okay, here is the nitty and the gritty.
'We're going to check each word, and decide if it's an RTF control word we want
'to deal with, a vbcrlf that we want to deal with, plain-old-text, or
'plain-old-text that ends with a vbcrlf. The RTF codes we will deal with are:
' \i which is the start of italic text
' \i0 which is the end of italic text
' \b which is the start of bold text
' \b0 which is the end of the bold text
' \ul which is the start of underlined text
' \ulnone which is the end of underlined text
' \'b7 which we will use to indicate the start of a bulleted line
' \par which is RTF paragraph
' All other RTF control words are skipped
'----------------------------------------------------------------------------------


    Select Case sRTFWord
        
        Case "\i"
            sHoldHtml = sHoldHtml & "<i>"
            subPushCode "</i>"
            
        Case "\i0"
                subPopCode
                
        Case "\b"
            sHoldHtml = sHoldHtml & "<b>"
            subPushCode "</b>"
           
        Case "\b0"
               subPopCode
            
        Case "\ul"
            sHoldHtml = sHoldHtml & "<u>"
            subPushCode "</u>"
            
        Case "\ulnone"
           subPopCode
             
        Case "\'b7"
            
           ' sHoldHtml = sHoldHtml & "*"
            If Not bBullet Then
                bBullet = True
                subPushCode "</ul>"
                subPushCode "</li>"
                sHoldHtml = sHoldHtml & "<ul><li>"
            Else
                sHoldHtml = sHoldHtml & "</li><li>"
            End If
            
        Case "\par"
            If bBullet And (InStr(sRTFLine, "\'") = 0) Then
                bBullet = False
                iCodeCounter = 1
                subPopCode
            End If
            
           
        Case vbCrLf
            sHoldHtml = sHoldHtml & "<br>"
            
        Case Else
            
            sRTFLeft = Left(sRTFWord, 1)
            Select Case sRTFLeft
'-----------------------------------------------------------------------------------

'If the left character is a backslash then we have an RTF command word that we
'don't recognize and don't want to deal with. We will ignore it.
'-----------------------------------------------------------------------------------
                Case "\"
                   'Ignore unknown RTF command
                
                Case Else
'-----------------------------------------------------------------------------------
' The last character in the RTF text that we are processing should be a }
' When we reach this character we are done. We can set sRTFWord to "" so we don't
' include the } in our converted html.
'-----------------------------------------------------------------------------------

                    If InStr(sRTFWord, "}") Then
                        sRTFWord = ""
                        
                    ElseIf Right(sRTFWord, 2) = vbCrLf Then
                        sRTFWord = sRTFWord & "<BR>"
                        
                    End If
                    
                    sHoldHtml = sHoldHtml & sRTFWord
                    
            End Select
    End Select
    
    If Len(sRTFRight) > 0 Then
        sHoldHtml = sHoldHtml & sRTFRight
        sRTFRight = ""
        
    End If
    
End Sub
Function funGetNextLine(sRTF As String) As String

'-----------------------------------------------------------------------------------
' When the RTF text is loaded into sRTFRemaining, it's loaded as one long string of
' text, even though each line of that text is seperated by a carriage return/
' linefeed. This function finds the carriage return / linefeed, and removes one
' line of text from sRTFRemaining, and puts it in sHoldLine for further processing.
'-----------------------------------------------------------------------------------
    
    Dim sHoldLine As String
    
    iFindEOL = InStr(sRTF, vbCrLf)
    
    If iFindEOL > 0 Then
        sHoldLine = Left(sRTF, iFindEOL + 1)
    Else
        sHoldLine = sRTF
    End If
    
    funGetNextLine = sHoldLine
    
End Function

Sub subPushCode(sRTFString As String)

'-----------------------------------------------------------------------------------
' This subroutine is something that I borrowed and modified for my own use from
' Brady Hegberg at http://www2.bitstream.net/~bradyh/downloads/rtf2htmlrm.html
' He has a far more impressive module to convert RTF to html, but it didn't do
' bulleted lists, and I had difficulty following to code. Sometimes no matter
' how well the code is written, you have to go "Huh?"
' This routine stores html code in the array sCodes, and increments
' iCodeCounter.
'-----------------------------------------------------------------------------------

    Dim lUbound As Long
    lUbound = UBound(sCodes)
    
    ReDim Preserve sCodes(UBound(sCodes) + 1)
    sCodes(UBound(sCodes)) = sRTFString
    iCodeCounter = iCodeCounter + 1
    
End Sub
Sub subPopCode()
    
'-----------------------------------------------------------------------------------
' Another routine I borrowed and modified from Barry. This routine subtracts one
' from iCodeCounter. If iCodeCounter is then zero, it means there is html code
' in the sCodes array that needs to be added to sHoldHTML. This is kind of a
' First-in/Last-out stack for those familiar with assembly language.
'-----------------------------------------------------------------------------------
    
    
    iCodeCounter = iCodeCounter - 1
    
    If iCodeCounter = 0 Then
     
        For I = UBound(sCodes) To 1 Step -1
            sHoldHtml = sHoldHtml & sCodes(I)
            
            
        Next I
        subClearScodes
    End If
    
End Sub
Sub subClearScodes()

'-----------------------------------------------------------------------------------
' Clear any remaining code out of the sCodes array by re-dimensioning the array.
'-----------------------------------------------------------------------------------

    ReDim sCodes(0)
    
End Sub
