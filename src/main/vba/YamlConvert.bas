Function toYaml(ByVal inDict As Object)
    If Not (TypeOf inDict Is Object ) Then
        toYaml = ""
    End If
    
    parentKey = ""
    startingTags = ""
    If inDict.Count > 0 Then
        If InStr(1, inDict.keys()(0), ".") > 0 Then
            parentKey = Split(inDict.keys()(0), ".")(0)
            startingTags = parentKey & ": {"
            endingTags = " }"
        Else
            startingTags = "{ "
            endingTags = " }"
        End If
    End If
    
    keyValuePairs = ""
    For Each sKey In inDict.keys()
        sItem = inDict.Item(sKey)
        If InStr(1, sKey, ".") > 0 Then
            sKey = Split(sKey, ".")(1)
        End If
        keyValuePairs = keyValuePairs & " " & Trim(sKey) & ": '" & Trim(sItem) & "' ,"
    Next
    keyValuePairs = Left(keyValuePairs, Len(keyValuePairs) - 1)  'Remove last comma
    toYaml = startingTags & keyValuePairs & endingTags
End Function

Function toKeyValuePairs(inputString As String)
    'Source: https://stackoverflow.com/questions/38738162/yaml-parser-for-excel-vba
    
    'Replace crlf with cr
    inputString = Replace(inputString, vbCrLf, vbCr)
    'Replace lf with cr
    inputString = Replace(inputString, vbLf, vbCr)
    
    bracesOpen = False
    quoteOpen = False
    keyNameParent = ""
    keyName = ""
    KeyValue = ""
    currentSegment = ""
    Errors = ""
    commentOpen = False
    Set dictOfValues = CreateObject("Scripting.Dictionary")
    
    For curIndex = 1 To Len(inputString)
        curCharacter = Mid(inputString, curIndex, 1)
        If commentOpen And Not curCharacter = vbCr Then
            'discard all comment text.
            curCharacter = ""
        Else
            commentOpen = False
        End If
    
        Select Case curCharacter
            Case "#"
                commentOpen = True
            Case "{"
                'Ignore braces
                If Not keyName = "" Then
                    keyNameParent = Trim(keyName)
                End If
                keyName = ""
                KeyValue = ""
                currentSegment = ""
            Case "}"
                'Ignore braces
                keyNameParent = ""
            Case "'"
                If quoteOpen Then
                    KeyValue = currentSegment
                    If dictOfValues.exists(keyName) Then
                        Errors = Errors & vbCrLf & "Cannot overwrite existing key."
                    Else
                        If keyNameParent = "" Then
                            Call dictOfValues.Add(keyName, KeyValue)
                        Else
                            Call dictOfValues.Add(keyNameParent & "." & keyName, KeyValue)
                        End If
                        currentSegment = ""
                        KeyValue = ""
                        keyName = ""
                    End If
                End If
                quoteOpen = Not quoteOpen
            Case ","
                If Not keyName = "" Then
                    KeyValue = Trim(currentSegment)
                    If keyNameParent = "" Then
                        Call dictOfValues.Add(keyName, KeyValue)
                    Else
                        Call dictOfValues.Add(keyNameParent & "." & keyName, KeyValue)
                    End If
                    keyName = ""
                    KeyValue = ""
                    currentSegment = ""
                End If
                currentSegment = ""
            Case vbCr
                If quoteOpen Then
                    Errors = Errors & vbCrLf & "New line not allowed inside value"
                Else
                    If Not keyName = "" Then
                        KeyValue = Trim(currentSegment)
                        If keyNameParent = "" Then
                            Call dictOfValues.Add(keyName, KeyValue)
                        Else
                            Call dictOfValues.Add(keyNameParent & "." & keyName, KeyValue)
                        End If
                        keyName = ""
                        KeyValue = ""
                        currentSegment = ""
                    End If
                End If
                currentSegment = ""
            Case vbLf
                'ignore linefeed
            Case ":"
                If quoteOpen Then
                    'Do nothing
                Else
                    keyName = Trim(currentSegment)
                    currentSegment = ""
                End If
            Case Else
                currentSegment = currentSegment & curCharacter
        End Select
    Next
    
    If Not keyName = "" And Not currentSegment = "" Then
        KeyValue = Trim(currentSegment)
        If keyNameParent = "" Then
            Call dictOfValues.Add(keyName, KeyValue)
        Else
            Call dictOfValues.Add(keyNameParent & "." & keyName, KeyValue)
        End If
        keyName = ""
        KeyValue = ""
        currentSegment = ""
    End If
    
    If Not Errors = "" Then
        Call dictOfValues.Add("Errors", Errors)
    End If
    Set toKeyValuePairs = dictOfValues
End Function






