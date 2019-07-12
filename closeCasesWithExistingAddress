Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)

Const WAIT_TIMEOUT = 300
Const ERR_TIMEOUT = 1000
Const READYSTATE_COMPLETE = 4

Sub InactivateAddressesCases()

Dim ie As InternetExplorer
Dim html As HTMLDocument
Dim sfdcArray() As String
Dim i As Integer
Dim w As Integer
Dim x As Integer
Dim inner As Range
Dim region As Range
Dim firstRow As Boolean
Set ie = New InternetExplorer
Set region = Range("M2") 'SFDC CASE REGION (NA/EU)
w = 1   'counter for selecting any other row than the first row
howManyCases = Range("N2")  'how many cases to work on, based on the N2 spreadsheet column
howManyCases = howManyCases + 1
Call ClearCells
firstRow = True
ie.Visible = True
ie.Navigate "https://login.salesforce.com/"
            Call ieBusy(ie)
ie.Document.getElementById("Login").Click   'clicks on the Login SFDC button
            Sleep 1000
            Call ieBusy(ie)
            Sleep 1000
Call RegionChecking(region, ie)
            Call ieBusy(ie)
ie.Document.querySelector(".x-grid3-body div:first-child .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the first case from the list
            Call ieBusy(ie)

For x = 2 To howManyCases

    'gets the location# and assigns it to the A & x cell (A2/A3/etc)
    Range("A" & x) = ie.Document.querySelector(".dataRow .dataCell:nth-child(3)").innerHTML
    sfdcArray = Split(Range("A" & x), "Address:")
    Range("A" & x) = sfdcArray(1)
    sfdcArray = Split(Range("A" & x), ":")
    Range("A" & x) = sfdcArray(0)
    Range("A" & x) = onlyDigits(Range("A" & x))
    'gets the case# and assigns it to the B & x cell (B2/B3/etc)
    Range("B" & x) = ie.Document.querySelector(".bPageTitle .ptBody .content .pageDescription").innerHTML
    sfdcArray = Split(Range("B" & x), "<")
    Range("B" & x) = sfdcArray(0)
            Call ieBusy(ie)
    'enters the location# in the search bar
    ie.Document.querySelector(".searchBoxClearContainer input:first-child").Value = Range("A" & x)
    'clicks on the search button
    ie.Document.querySelector("#phSearchForm .headerSearchContainer .headerSearchLeftRoundedCorner .headerSearchRightRoundedCorner input:first-child").Click
            Call ieBusy(ie)
    'how many addresses have been found in SFDC
    Range("C" & x) = onlyDigits(ie.Document.querySelector(".searchEntityList .itemLink:nth-child(2) .item .linkSelector .resultCount").innerHTML)
            Call ieBusy(ie)
    
    If Range("C" & x) = 0 Then      'if no addresses have been found, goes to the next iteration
        w = w + 1
        Range("D" & x) = "address not found"
        Range("E" & x) = "No"
        firstRow = False
        GoTo NextIteration
        Exit For
        
    ElseIf Range("C" & x) > 1 Then      'if more than one address has been found
        'checks the addresses results and compares them to the site# from the case
        Set resultOptions = ie.Document.getElementById("Business_Address__c_body").getElementsByClassName("dataRow")
            For Each a In resultOptions
                Range("D" & x) = a.innerHTML
                sfdcArray = Split(Range("D" & x), "<th")
                Range("D" & x) = sfdcArray(1)
                sfdcArray = Split(Range("D" & x), "</th>")
                Range("D" & x) = sfdcArray(0)
                Set inner = Range("D" & x)
                    If InStr(inner.Value, Range("A" & x)) > 0 Then
                        a.querySelector("th a").Click
                        Exit For
                    ElseIf innerHTML <> Range("A" & x) Then
                        w = w + 1
                        Range("D" & x) = "address not found"
                        Range("E" & x) = "No"
                        firstRow = False
                        GoTo NextIteration
                        Exit For
                    End If
            Next
            
    ElseIf Range("C" & x) = 1 Then      'if only one address has been found
    
        Set resultOptions = ie.Document.getElementById("Business_Address__c_body").getElementsByClassName("dataRow")    'checks the results address# vs the address# from the case
                        For Each a In resultOptions
                            Range("D" & x) = a.innerHTML
                            Set inner = Range("D" & x)
                                If InStr(inner.Value, Range("A" & x)) > 0 Then
                                    a.querySelector("th a").Click
                                    Exit For
                                End If
                        Next
    End If
    
                Call ieBusy(ie)
    Range("D" & x) = ie.Document.getElementById("00NF0000008W7z8_ileinner").innerHTML   'retrieves the business address number again, to double check for correct entry
    
            If Range("D" & x) = Range("A" & x) Then
                    Call FlagChecking(sfdcArray, ie, x)
            End If
            
    If Range("F" & x).Interior.ColorIndex = 4 Then
        If Range("G" & x).Interior.ColorIndex = 4 Then
            If Range("H" & x).Interior.ColorIndex = 4 Then
                If Range("I" & x).Interior.ColorIndex = 3 Then
                    If Range("J" & x).Interior.ColorIndex = 4 Then
                        Range("D" & x) = ie.Document.getElementById("00NF0000008W7z8_ileinner").innerHTML   'retrieves the business address number again, to double check for correct entry
                            If Range("D" & x) = Range("A" & x) Then
                            
                                Call RegionChecking(region, ie)
                                Call ieBusy(ie)
                                Call SelectRow(ie, firstRow, w)
                                Call ieBusy(ie)
                                Call CaseClosing(ie)
                                Call ieBusy(ie)
                                Range("D" & x) = "address found"
                                Range("E" & x) = "Yes"
                                Range("K" & x) = ie.Document.getElementById("cas7_ileinner").innerHTML
                                Call ieBusy(ie)
                                GoTo NextIteration
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    Range("D" & x) = ie.Document.getElementById("00NF0000008W7z8_ileinner").innerHTML   'retrieves the business address number again, to double check for correct entry
    
    If Range("D" & x) = Range("A" & x) Then     'if the address number is correct:
            ie.Document.querySelector("#topButtonRow input:nth-child(3)").Click     'click on edit on the business address' page
        ElseIf Range("D" & x) <> Range("A" & x) Then    'if the address number does not match
            w = w + 1
            Range("D" & x) = "address not found"
            Range("E" & x) = "No"
            firstRow = False
            GoTo NextIteration
        End If
        
                Call ieBusy(ie)
    ie.Document.querySelector(".bPageBlock .pbBody .pbSubsection .detailList tbody:first-child tr:nth-child(4) td:nth-child(4) textarea:first-child").Value = "Inactivated per case# " & Range("B" & x) 'SFDC comment
    Call FlagClick(ie, x)
    ie.Document.querySelector(".pbBottomButtons .pbButtonb input:first-child").Click       'save button
                Call ieBusy(ie)
    Call FlagChecking(sfdcArray, ie, x)
    Call RegionChecking(region, ie)
                Call ieBusy(ie)
    Call SelectRow(ie, firstRow, w)
                Call ieBusy(ie)
    Range("K" & x) = ie.Document.querySelector(".bPageTitle .ptBody .content .pageDescription").innerHTML       'gets the case# and checks it agains the ID from the B cell
    sfdcArray = Split(Range("K" & x), "<")
    Range("K" & x) = sfdcArray(0)
    
            If Range("B" & x) = Range("K" & x) Then
                    Call ieBusy(ie)
                    Call CaseClosing(ie)
                    Range("D" & x) = "address found"
                    Range("E" & x) = "Yes"
                    Call ieBusy(ie)
                    Range("K" & x) = ie.Document.getElementById("cas7_ileinner").innerHTML
                    Call ieBusy(ie)
            ElseIf Range("B" & x) <> Range("K" & x) Then
                    Range("K" & x) = "Not Closed"
                    w = w + 1
                    Range("E" & x) = "No"
                    firstRow = False
            End If
     
NextIteration:

    Call RegionChecking(region, ie)
                    Call ieBusy(ie)
    Call SelectRow(ie, firstRow, w)
                    Call ieBusy(ie)
Next x

MsgBox "Charlie Oscar Mike"

End Sub
Function ClearCells() 'clears the cells' content

    Range("A2:K99").Clear

End Function
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

Function RegionChecking(region As Range, ie As Object)

    If region = "EU" Then
        ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VW5q"     'navigates to the reports page for EU
    ElseIf region = "NA" Then
        ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VcF4"     'navigates to the reports page for NA
    End If

End Function

Function SelectRow(ie As Object, firstRow As Boolean, w As Integer)

     If firstRow = True Then
        ie.Document.querySelector(".x-grid3-body div:first-child .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the first case from the list
     ElseIf firstRow = False Then
        ie.Document.querySelector(".x-grid3-body div:nth-child(" & w & ") .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the nth case from the list
     End If

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

Function FlagChecking(sfdcArray() As String, ie As Object, x As Integer)

                    Range("D" & x) = ie.Document.getElementById("00NF0000008W7z9_ileinner").innerHTML   'checking the Bill-to flag
                    sfdcArray = Split(Range("D" & x), "alt=")
                    Range("D" & x) = sfdcArray(1)
                    Set inner = Range("D" & x)
                        If InStr(inner.Value, "Not Checked") > 0 Then
                            Range("F" & x).Interior.ColorIndex = 4
                        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
                            Range("F" & x).Interior.ColorIndex = 3
                        End If
                        
                    Range("D" & x) = ie.Document.getElementById("00NF0000008W7zc_ileinner").innerHTML   'checking the Ship-to flag
                    sfdcArray = Split(Range("D" & x), "alt=")
                    Range("D" & x) = sfdcArray(1)
                    Set inner = Range("D" & x)
                        If InStr(inner.Value, "Not Checked") > 0 Then
                            Range("G" & x).Interior.ColorIndex = 4
                        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
                            Range("G" & x).Interior.ColorIndex = 3
                        End If
                        
                    Range("D" & x) = ie.Document.getElementById("00NF0000008W7zH_ileinner").innerHTML   'checking the Deliver-to flag
                    sfdcArray = Split(Range("D" & x), "alt=")
                    Range("D" & x) = sfdcArray(1)
                    Set inner = Range("D" & x)
                        If InStr(inner.Value, "Not Checked") > 0 Then
                            Range("H" & x).Interior.ColorIndex = 4
                        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
                            Range("H" & x).Interior.ColorIndex = 3
                        End If
                        
                    Range("D" & x) = ie.Document.getElementById("00NF0000008W7zL_ileinner").innerHTML   'checking the Inactive flag
                    sfdcArray = Split(Range("D" & x), "alt=")
                    Range("D" & x) = sfdcArray(1)
                    Set inner = Range("D" & x)
                        If InStr(inner.Value, "Not Checked") > 0 Then
                            Range("I" & x).Interior.ColorIndex = 4
                        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
                            Range("I" & x).Interior.ColorIndex = 3
                        End If
                        
                    Range("D" & x) = ie.Document.getElementById("00N2A00000DSnY0_ileinner").innerHTML   'checking the Invoice-to flag
                    sfdcArray = Split(Range("D" & x), "alt=")
                    Range("D" & x) = sfdcArray(1)
                    Set inner = Range("D" & x)
                        If InStr(inner.Value, "Not Checked") > 0 Then
                            Range("J" & x).Interior.ColorIndex = 4
                        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
                            Range("J" & x).Interior.ColorIndex = 3
                        End If

End Function

Function FlagClick(ie As Object, x As Integer)

    Range("D" & x) = ie.Document.querySelector(".bPageBlock .pbBody div:nth-child(9) .detailList tbody:first-child tr:first-child td:nth-child(4)").innerHTML 'checking for checked flags
    Set inner = Range("D" & x)
    If InStr(inner.Value, "checked") > 0 Then
        ie.Document.getElementById("00NF0000008W7z9").Click     'bill-to flag
    End If
    
    Range("D" & x) = ie.Document.querySelector(".bPageBlock .pbBody div:nth-child(9) .detailList tbody:first-child tr:nth-child(2) td:nth-child(4)").innerHTML 'checking for checked flags
    Set inner = Range("D" & x)
        If InStr(inner.Value, "checked") > 0 Then
            ie.Document.getElementById("00NF0000008W7zc").Click     'ship-to flag
        End If
        
    Range("D" & x) = ie.Document.querySelector(".bPageBlock .pbBody div:nth-child(9) .detailList tbody:first-child tr:nth-child(3) td:nth-child(4)").innerHTML 'checking for checked flags
    Set inner = Range("D" & x)
        If InStr(inner.Value, "checked") > 0 Then
            ie.Document.getElementById("00NF0000008W7zH").Click     'deliver-to flag
        End If
    
    Range("D" & x) = ie.Document.querySelector(".bPageBlock .pbBody div:nth-child(9) .detailList tbody:first-child tr:nth-child(4) td:nth-child(4)").innerHTML 'checking for checked flags
    Set inner = Range("D" & x)
        If InStr(inner.Value, "checked") = 0 Then
            ie.Document.getElementById("00NF0000008W7zL").Click     'inactive flag
        End If
    
    Range("D" & x) = ie.Document.querySelector(".bPageBlock .pbBody div:nth-child(9) .detailList tbody:first-child tr:nth-child(5) td:nth-child(4)").innerHTML 'checking for checked flags
    Set inner = Range("D" & x)
        If InStr(inner.Value, "checked") > 0 Then
            ie.Document.getElementById("00N2A00000DSnY0").Click     'invoice-to flag
        End If

End Function

Function CaseClosing(ie As Object)

ie.Document.querySelector(".oRight .bPageBlock .pbHeader table:first-child tbody:first-child tr:first-child .pbButton input:nth-child(4)").Click      'clicks on "Close Case"

                                            Call ieBusy(ie)
                                            Set dropOptions = ie.Document.getElementById("cas7")    'selects the closed status from the dropdown
                                                For Each o In dropOptions.Options
                                                    If o.Value = "Closed" Then
                                                        o.Selected = True
                                                        Exit For
                                                    End If
                                                Next
                                                
                                            Set dropOptions = ie.Document.getElementById("cas6")    'selects the "Solution Delivered" reason from the dropdown
                                                For Each o In dropOptions.Options
                                                    If o.Value = "Solution Delivered" Then
                                                        o.Selected = True
                                                        Exit For
                                                    End If
                                                Next
                                                
                                            ie.Document.getElementById("00NA00000045ZfG").Value = "case completed"  'leave the closure details note
                                            
                                            Set dropOptions = ie.Document.getElementById("00N5A00000LmTdJ")    'selects the "Manual fix in systems internally" reason from the dropdown
                                                For Each o In dropOptions.Options
                                                    If o.Value = "Manual fix in systems internally" Then
                                                        o.Selected = True
                                                        Exit For
                                                    End If
                                                Next
                                                
                                            ie.Document.querySelector(".pbButtonb input:first-child").Click 'clicks on Save button
Call ieBusy(ie)

End Function
