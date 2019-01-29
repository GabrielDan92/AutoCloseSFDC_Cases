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

Do While ie.readyState <> READYSTATE_COMPLETE
    DoEvents
Loop

'ieBusy ie

Sleep 2000

ie.document.getElementById("Login").Click   'clicks on the Login SFDC button
Do While ie.readyState <> READYSTATE_COMPLETE
    DoEvents
Loop

'ieBusy ie
Sleep 5000

If region = "EU" Then
    ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VW5q"     'navigates to the reports page for EU
ElseIf region = "NA" Then
    ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VcF4"     'navigates to the reports page for NA
End If


Sleep 8000
'ieBusy ie
ie.document.querySelector(".x-grid3-body div:first-child .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the first case from the list
Sleep 3500
'ieBusy ie

firstRow = True
w = 106
howManyCases = Range("N2")  'how many cases to work on, based on the N2 column
howManyCases = howManyCases + 1

For x = 2 To howManyCases
    
    
    Range("A" & x) = ie.document.querySelector(".dataRow .dataCell:nth-child(3)").innerHTML     'gets the location# and assigns it to the A & x cell (A2/A3/etc)
    sfdcArray = Split(Range("A" & x), "Address:")
    Range("A" & x) = sfdcArray(1)
    sfdcArray = Split(Range("A" & x), ":")
    Range("A" & x) = sfdcArray(0)
    Range("A" & x) = onlyDigits(Range("A" & x))
    Range("B" & x) = ie.document.querySelector(".bPageTitle .ptBody .content .pageDescription").innerHTML       'gets the case# and assigns it to the B & x cell (B2/B3/etc)
    sfdcArray = Split(Range("B" & x), "<")
    Range("B" & x) = sfdcArray(0)
    Sleep 3000
    ie.document.querySelector(".searchBoxClearContainer input:first-child").Value = Range("A" & x)     'entering the location# in the search bar
    ie.document.querySelector("#phSearchForm .headerSearchContainer .headerSearchLeftRoundedCorner .headerSearchRightRoundedCorner input:first-child").Click    'clicks on the search button
    Sleep 5000
    Range("C" & x) = onlyDigits(ie.document.querySelector(".searchEntityList .itemLink:nth-child(2) .item .linkSelector .resultCount").innerHTML)    'how many addresses have been found in SFDC
    Sleep 500
    
    If Range("C" & x) = 0 Then      'if no addresses have been found, goes to the next iteration
        w = w + 1
        Range("D" & x) = "address not found"
        Range("E" & x) = "No"
        firstRow = False
        GoTo nextiteration
        Exit For
        
    ElseIf Range("C" & x) > 1 Then
    
        Set resultOptions = ie.document.getElementById("Business_Address__c_body").getElementsByClassName("dataRow")    'checks the addresses results and compares them to the site# from the case
        
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
                        GoTo nextiteration
                        Exit For
                    End If
            Next
            
    ElseIf Range("C" & x) = 1 Then
    
        Set resultOptions = ie.document.getElementById("Business_Address__c_body").getElementsByClassName("dataRow")
                        For Each a In resultOptions
                            Range("D" & x) = a.innerHTML
                            Set inner = Range("D" & x)
                                If InStr(inner.Value, Range("A" & x)) > 0 Then
                                    a.querySelector("th a").Click
                                    Exit For
                                End If
                        Next
    End If
    
    
    Sleep 5000
    Range("D" & x) = ie.document.getElementById("00NF0000008W7z8_ileinner").innerHTML   'retrieves the business address number again, to double check for correct entry
    If Range("D" & x) = Range("A" & x) Then
    
                    Range("D" & x) = ie.document.getElementById("00NF0000008W7z9_ileinner").innerHTML   'checking the Bill-to flag
                    sfdcArray = Split(Range("D" & x), "alt=")
                    Range("D" & x) = sfdcArray(1)
                    Set inner = Range("D" & x)
                        If InStr(inner.Value, "Not Checked") > 0 Then
                            Range("F" & x).Interior.ColorIndex = 4
                        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
                            Range("F" & x).Interior.ColorIndex = 3
                        End If
                        
                    Range("D" & x) = ie.document.getElementById("00NF0000008W7zc_ileinner").innerHTML   'checking the Ship-to flag
                    sfdcArray = Split(Range("D" & x), "alt=")
                    Range("D" & x) = sfdcArray(1)
                    Set inner = Range("D" & x)
                        If InStr(inner.Value, "Not Checked") > 0 Then
                            Range("G" & x).Interior.ColorIndex = 4
                        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
                            Range("G" & x).Interior.ColorIndex = 3
                        End If
                        
                    Range("D" & x) = ie.document.getElementById("00NF0000008W7zH_ileinner").innerHTML   'checking the Deliver-to flag
                    sfdcArray = Split(Range("D" & x), "alt=")
                    Range("D" & x) = sfdcArray(1)
                    Set inner = Range("D" & x)
                        If InStr(inner.Value, "Not Checked") > 0 Then
                            Range("H" & x).Interior.ColorIndex = 4
                        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
                            Range("H" & x).Interior.ColorIndex = 3
                        End If
                        
                    Range("D" & x) = ie.document.getElementById("00NF0000008W7zL_ileinner").innerHTML   'checking the Inactive flag
                    sfdcArray = Split(Range("D" & x), "alt=")
                    Range("D" & x) = sfdcArray(1)
                    Set inner = Range("D" & x)
                        If InStr(inner.Value, "Not Checked") > 0 Then
                            Range("I" & x).Interior.ColorIndex = 4
                        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
                            Range("I" & x).Interior.ColorIndex = 3
                        End If
                        
                    Range("D" & x) = ie.document.getElementById("00N2A00000DSnY0_ileinner").innerHTML   'checking the Invoice-to flag
                    sfdcArray = Split(Range("D" & x), "alt=")
                    Range("D" & x) = sfdcArray(1)
                    Set inner = Range("D" & x)
                        If InStr(inner.Value, "Not Checked") > 0 Then
                            Range("J" & x).Interior.ColorIndex = 4
                        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
                            Range("J" & x).Interior.ColorIndex = 3
                        End If
            End If
        
        
        
    If Range("F" & x).Interior.ColorIndex = 4 Then
        If Range("G" & x).Interior.ColorIndex = 4 Then
            If Range("H" & x).Interior.ColorIndex = 4 Then
                If Range("I" & x).Interior.ColorIndex = 3 Then
                    If Range("J" & x).Interior.ColorIndex = 4 Then
                            Range("D" & x) = ie.document.getElementById("00NF0000008W7z8_ileinner").innerHTML   'retrieves the business address number again, to double check for correct entry
                                If Range("D" & x) = Range("A" & x) Then
                                
                                            If region = "EU" Then
                                                ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VW5q"     'navigates to the reports page for EU
                                            ElseIf region = "NA" Then
                                                ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VcF4"     'navigates to the reports page for NA
                                            End If

                        
                                            Do While ie.readyState <> READYSTATE_COMPLETE
                                                DoEvents
                                            Loop
                                            Sleep 3000
                                            
                                             If firstRow = True Then
                                                ie.document.querySelector(".x-grid3-body div:first-child .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the first case from the list
                                             ElseIf firstRow = False Then
                                                ie.document.querySelector(".x-grid3-body div:nth-child(" & w & ") .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the nth case from the list
                                             End If
                                             
                                            Sleep 5000
                                            ie.document.querySelector(".oRight .bPageBlock .pbHeader table:first-child tbody:first-child tr:first-child .pbButton input:nth-child(4)").Click      'clicks on "Close Case"
                                            Sleep 3500
                                            
                                            Set dropOptions = ie.document.getElementById("cas7")    'selects the closed status from the dropdown
                                            
                                                For Each o In dropOptions.Options
                                                    If o.Value = "Closed" Then
                                                        o.Selected = True
                                                        Exit For
                                                    End If
                                                Next
                                                
                                            Set dropOptions = ie.document.getElementById("cas6")    'selects the reason from the dropdown
                                            
                                                For Each o In dropOptions.Options
                                                    If o.Value = "Solution Delivered" Then
                                                        o.Selected = True
                                                        Exit For
                                                    End If
                                                Next
                                                    
                                             ie.document.getElementById("00NA00000045ZfG").Value = "case completed"
                                             ie.document.querySelector(".pbButtonb input:first-child").Click 'clicks on Save button
                                             Sleep 4000
                                             Range("D" & x) = "address found"
                                             Range("E" & x) = "Yes"
                                             Range("K" & x) = ie.document.getElementById("cas7_ileinner").innerHTML
                                             Sleep 4000
                                             GoTo nextiteration
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                            
     

        
        
    Range("D" & x) = ie.document.getElementById("00NF0000008W7z8_ileinner").innerHTML   'retrieves the business address number again, to double check for correct entry
    If Range("D" & x) = Range("A" & x) Then
            ie.document.querySelector("#topButtonRow input:nth-child(3)").Click     'click on edit on the business address' page
        ElseIf Range("D" & x) <> Range("A" & x) Then
            w = w + 1
            Range("D" & x) = "address not found"
            Range("E" & x) = "No"
            firstRow = False
            GoTo nextiteration
        End If
        
        
    
     Sleep 3000
    ie.document.querySelector(".bPageBlock .pbBody .pbSubsection .detailList tbody:first-child tr:nth-child(4) td:nth-child(4) textarea:first-child").Value = "Inactivated per case# " & Range("B" & x)
    Range("D" & x) = ie.document.querySelector(".bPageBlock .pbBody div:nth-child(9) .detailList tbody:first-child tr:first-child td:nth-child(4)").innerHTML 'checking for checked flags
    Set inner = Range("D" & x)
    If InStr(inner.Value, "checked") > 0 Then
        ie.document.getElementById("00NF0000008W7z9").Click     'bill-to flag
    End If
    
    Range("D" & x) = ie.document.querySelector(".bPageBlock .pbBody div:nth-child(9) .detailList tbody:first-child tr:nth-child(2) td:nth-child(4)").innerHTML 'checking for checked flags
    Set inner = Range("D" & x)
        If InStr(inner.Value, "checked") > 0 Then
            ie.document.getElementById("00NF0000008W7zc").Click     'ship-to flag
        End If
        
    Range("D" & x) = ie.document.querySelector(".bPageBlock .pbBody div:nth-child(9) .detailList tbody:first-child tr:nth-child(3) td:nth-child(4)").innerHTML 'checking for checked flags
    Set inner = Range("D" & x)
        If InStr(inner.Value, "checked") > 0 Then
            ie.document.getElementById("00NF0000008W7zH").Click     'deliver-to flag
        End If
    
    Range("D" & x) = ie.document.querySelector(".bPageBlock .pbBody div:nth-child(9) .detailList tbody:first-child tr:nth-child(4) td:nth-child(4)").innerHTML 'checking for checked flags
    Set inner = Range("D" & x)
        If InStr(inner.Value, "checked") = 0 Then
            ie.document.getElementById("00NF0000008W7zL").Click     'inactive flag
        End If
    
    Range("D" & x) = ie.document.querySelector(".bPageBlock .pbBody div:nth-child(9) .detailList tbody:first-child tr:nth-child(5) td:nth-child(4)").innerHTML 'checking for checked flags
    Set inner = Range("D" & x)
        If InStr(inner.Value, "checked") > 0 Then
            ie.document.getElementById("00N2A00000DSnY0").Click     'invoice-to flag
        End If
        
    ie.document.querySelector(".pbBottomButtons .pbButtonb input:first-child").Click       'save button
    
    Sleep 3000
    Range("D" & x) = ie.document.getElementById("00NF0000008W7z9_ileinner").innerHTML   'checking the Bill-to flag
    sfdcArray = Split(Range("D" & x), "alt=")
    Range("D" & x) = sfdcArray(1)
    Set inner = Range("D" & x)
        If InStr(inner.Value, "Not Checked") > 0 Then
            Range("F" & x).Interior.ColorIndex = 4
        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
            Range("F" & x).Interior.ColorIndex = 3
        End If
        
    Range("D" & x) = ie.document.getElementById("00NF0000008W7zc_ileinner").innerHTML   'checking the Ship-to flag
    sfdcArray = Split(Range("D" & x), "alt=")
    Range("D" & x) = sfdcArray(1)
    Set inner = Range("D" & x)
        If InStr(inner.Value, "Not Checked") > 0 Then
            Range("G" & x).Interior.ColorIndex = 4
        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
            Range("G" & x).Interior.ColorIndex = 3
        End If
        
    Range("D" & x) = ie.document.getElementById("00NF0000008W7zH_ileinner").innerHTML   'checking the Deliver-to flag
    sfdcArray = Split(Range("D" & x), "alt=")
    Range("D" & x) = sfdcArray(1)
    Set inner = Range("D" & x)
        If InStr(inner.Value, "Not Checked") > 0 Then
            Range("H" & x).Interior.ColorIndex = 4
        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
            Range("H" & x).Interior.ColorIndex = 3
        End If
        
    Range("D" & x) = ie.document.getElementById("00NF0000008W7zL_ileinner").innerHTML   'checking the Inactive flag
    sfdcArray = Split(Range("D" & x), "alt=")
    Range("D" & x) = sfdcArray(1)
    Set inner = Range("D" & x)
        If InStr(inner.Value, "Not Checked") > 0 Then
            Range("I" & x).Interior.ColorIndex = 4
        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
            Range("I" & x).Interior.ColorIndex = 3
        End If
        
    Range("D" & x) = ie.document.getElementById("00N2A00000DSnY0_ileinner").innerHTML   'checking the Invoice-to flag
    sfdcArray = Split(Range("D" & x), "alt=")
    Range("D" & x) = sfdcArray(1)
    Set inner = Range("D" & x)
        If InStr(inner.Value, "Not Checked") > 0 Then
            Range("J" & x).Interior.ColorIndex = 4
        ElseIf InStr(inner.Value, "Not Checked") = 0 Then
            Range("J" & x).Interior.ColorIndex = 3
        End If
            
            
        If region = "EU" Then
            ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VW5q"     'navigates to the reports page for EU
        ElseIf region = "NA" Then
            ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VcF4"     'navigates to the reports page for NA
        End If

    
    
    Do While ie.readyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    Sleep 1500
    
     If firstRow = True Then
        ie.document.querySelector(".x-grid3-body div:first-child .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the first case from the list
     ElseIf firstRow = False Then
        ie.document.querySelector(".x-grid3-body div:nth-child(" & w & ") .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the nth case from the list
     End If
     
    
    Do While ie.readyState <> READYSTATE_COMPLETE
        DoEvents
    Loop
    Sleep 3000
    Range("K" & x) = ie.document.querySelector(".bPageTitle .ptBody .content .pageDescription").innerHTML       'gets the case# and checks it agains the ID from the B cell
    sfdcArray = Split(Range("K" & x), "<")
    Range("K" & x) = sfdcArray(0)
    
            If Range("B" & x) = Range("K" & x) Then
            
                    Sleep 2000
                    Do While ie.readyState <> READYSTATE_COMPLETE
                        DoEvents
                    Loop
                    ie.document.querySelector(".oRight .bPageBlock .pbHeader table:first-child tbody:first-child tr:first-child .pbButton input:nth-child(4)").Click      'clicks on "Close Case"
                    
                    
                    Do While ie.readyState <> READYSTATE_COMPLETE
                        DoEvents
                    Loop
                    Sleep 2000
                    
                    Set dropOptions = ie.document.getElementById("cas7")
                    
                        For Each o In dropOptions.Options
                            If o.Value = "Closed" Then
                                o.Selected = True
                                Exit For
                            End If
                        Next
                        
                    Set dropOptions = ie.document.getElementById("cas6")
                    
                        For Each o In dropOptions.Options
                            If o.Value = "Solution Delivered" Then
                                o.Selected = True
                                Exit For
                            End If
                        Next
                            
                     ie.document.getElementById("00NA00000045ZfG").Value = "case completed"
                     ie.document.querySelector(".pbButtonb input:first-child").Click 'clicks on Save button
                     Sleep 4000
                     Range("D" & x) = "address found"
                     Range("E" & x) = "Yes"
                     Sleep 3000
                     Range("K" & x) = ie.document.getElementById("cas7_ileinner").innerHTML
                     Sleep 2000
                     Do While ie.readyState <> READYSTATE_COMPLETE
                        DoEvents
                    Loop
                     
            ElseIf Range("B" & x) <> Range("K" & x) Then
            
                    Range("K" & x) = "Not Closed"
                    w = w + 1
                    Range("E" & x) = "No"
                    firstRow = False
                    
            End If
     
nextiteration:

     
        If region = "EU" Then
            ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VW5q"     'navigates to the reports page for EU
        ElseIf region = "NA" Then
            ie.Navigate "https://ptc.my.salesforce.com/500?fcf=00B5A000009VcF4"     'navigates to the reports page for NA
        End If

Do While ie.readyState <> READYSTATE_COMPLETE
    DoEvents
Loop
    
     Sleep 3000
     If firstRow = True Then
        ie.document.querySelector(".x-grid3-body div:first-child .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the first case from the list
     ElseIf firstRow = False Then
        ie.document.querySelector(".x-grid3-body div:nth-child(" & w & ") .x-grid3-row-table .x-grid3-col:nth-child(4) .x-grid3-cell-inner a:first-child").Click       'selects the nth case from the list
     End If
     Sleep 4000
     
Do While ie.readyState <> READYSTATE_COMPLETE
    DoEvents
Loop

Next x

MsgBox "Charlie Oscar Mike"

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
    Dim i, j, ready
    Dim fr As MSHTML.HTMLWindow2, allframes As HTMLIFrame
    ' wait for page to connect
    i = 0
    Do Until ie.readyState = READYSTATE_COMPLETE
        Sleep 100
        i = i + 1
        If i > WAIT_TIMEOUT Then
            Err.Raise ERR_TIMEOUT, , "Timeout"
        End If
    Loop

    ' wait for document to load
    Do Until ie.document.readyState = "complete"
        Sleep 100
        i = i + 1
        If i > WAIT_TIMEOUT Then
            Err.Raise ERR_TIMEOUT, , "Timeout"
        End If
    Loop

    ' wait for frames to load
    Do
        ready = True
        Set allframes = ie.document.frames
        
        For j = 0 To allframes.Length - 1
            Set fr = allframes.frames(j)
            If fr.document.readyState <> "complete" Then
                ready = False
                Sleep 100
                i = i + 1
                If i > WAIT_TIMEOUT Then
                    Err.Raise ERR_TIMEOUT, , "Timeout"
                End If
            End If
        Next
    Loop Until ready
End Sub

