Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Const WAIT_TIMEOUT = 300
Const ERR_TIMEOUT = 1000
Const READYSTATE_COMPLETE = 4


Sub InactivateAddressesCases()


Dim ie As InternetExplorer
Dim html As HTMLDocument
Dim sfdcArray() As String
Dim i As Integer
Dim inner As Range
Set ie = New InternetExplorer


region = Range("M2") 'SELECT THE SFDC CASE REGION (NA/EU)

ie.Visible = True
ie.Navigate "https://login.salesforce.com/"


Call ieBusy(ie)

ie.Document.getElementById("Login").Click   'clicks on the Login SFDC button

Call ieBusy(ie)

If region = "EU" Then
    ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VW5q"     'navigates to the reports page for EU
ElseIf region = "NA" Then
    ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VcF4"     'navigates to the reports page for NA
ElseIf region = "NA-CustReg" Then
    ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B2A000008sZ1l"     'navigates to the reports page for NA CUST REG
End If


Call ieBusy(ie)
Call ieBusy(ie)
Call ieBusy(ie)
Call ieBusy(ie)
ie.Document.querySelector(".x-grid3-body div:first-child .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the first case from the list


firstRow = True
w = 1
howManyCases = Range("N2")  'how many cases to work on, based on the N2 column
howManyCases = howManyCases + 1

For x = 2 To howManyCases

Call ieBusy(ie)
Call ieBusy(ie)
Call ieBusy(ie)

        ie.Document.querySelector(".x-grid3-body div:first-child .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the first case from the list
        Call ieBusy(ie)
        ie.Document.querySelector(".oRight .bPageBlock .pbHeader table:first-child tbody:first-child tr:first-child .pbButton input:nth-child(4)").Click      'clicks on "Close Case"
        Call ieBusy(ie)
                                Set dropOptions = ie.Document.getElementById("cas7")    'selects the closed status from the dropdown
                                    For Each o In dropOptions.Options
                                        If o.Value = "Closed" Then
                                            o.Selected = True
                                            Exit For
                                        End If
                                    Next
                                    
                                Set dropOptions = ie.Document.getElementById("cas6")    'selects the "Obsolete" reason from the dropdown
                                    For Each o In dropOptions.Options
                                        If o.Value = "Obsolete" Then
                                            o.Selected = True
                                            Exit For
                                        End If
                                    Next
                                    
                                 ie.Document.getElementById("00NA00000045ZfG").Value = "n/a"    'leave the closure details note
                                 
                                 Set dropOptions = ie.Document.getElementById("00N5A00000LmTdJ")    'selects the "Manual fix in systems internally" reason from the dropdown
                                    For Each o In dropOptions.Options
                                        If o.Value = "Manual fix in systems internally" Then
                                            o.Selected = True
                                            Exit For
                                        End If
                                    Next
                                 
                                 ie.Document.querySelector(".pbButtonb input:first-child").Click 'clicks on Save button
                                 Call ieBusy(ie)
                                 Range("E" & x) = "Yes"
                                 Range("K" & x) = ie.Document.getElementById("cas7_ileinner").innerHTML
                                 Sleep 4000
                                 GoTo NextIteration

NextIteration:
     
            If region = "EU" Then
                ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VW5q"     'navigates to the reports page for EU
            ElseIf region = "NA" Then
                ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VcF4"     'navigates to the reports page for NA
            ElseIf region = "NA-CustReg" Then
                ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B2A000008sZ1l"     'navigates to the reports page for NA CUST REG
            End If


Next x


End Sub

Function onlyDigits(s As String) As String 'returns only the digits from a string with multiple characters
    Dim retval As String
    Dim i As Integer
    retval = ""
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            retval = retval + Mid(s, i, 1)
        End If
    Next
    onlyDigits = retval
End Function


Sub ieBusy(ie As Object)
On Error GoTo nexxt
    Dim i, j, ready
    Dim fr As MSHTML.HTMLWindow2, allframes As HTMLIFrame
    ' wait for page to connect
    i = 0
    Do Until ie.READYSTATE = READYSTATE_COMPLETE
        Sleep 100
        i = i + 1
        If i > WAIT_TIMEOUT Then
            Err.Raise ERR_TIMEOUT, , "Timeout"
        End If
    Loop

    ' wait for document to load
    Do Until ie.Document.READYSTATE = "complete"
        Sleep 100
        i = i + 1
        If i > WAIT_TIMEOUT Then
            Err.Raise ERR_TIMEOUT, , "Timeout"
        End If
    Loop

    ' wait for frames to load
    Do
        ready = True
        Set allframes = ie.Document.frames
        
        For j = 0 To allframes.Length - 1
            Set fr = allframes.frames(j)
            If fr.Document.READYSTATE <> "complete" Then
                ready = False
                Sleep 100
                i = i + 1
                If i > WAIT_TIMEOUT Then
                    Err.Raise ERR_TIMEOUT, , "Timeout"
                End If
            End If
        Next
    Loop Until ready
    
nexxt:
End Sub
