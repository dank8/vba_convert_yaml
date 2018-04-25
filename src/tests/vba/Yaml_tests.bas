
Sub testParserYAML()
    Set newDict = YamlConvert.toKeyValuePairs("{ key1 : 'value1' " & vbCrLf & " key2 : 'value2' }")
    Call assertBoolean(newDict.Count = 2, "Braces, vbcrlf and quotes")
    Debug.Print YamlConvert.toYaml(newDict)
    Set newDict = YamlConvert.toKeyValuePairs("{ key1 : 'value1' , key2 : 'value2' }")
    Call assertBoolean(newDict.Count = 2, "Braces, comma and quotes")
    Debug.Print YamlConvert.toYaml(newDict)
    Set newDict = YamlConvert.toKeyValuePairs("testParent: { key1 : 'value1' , key2 : 'value2' }")
    Call assertBoolean(newDict.Count = 2, "Parent, comma, quoted")
    Call assertBoolean((InStr(1, newDict.keys()(0), ".") > 0), "Key has Parent")
    Debug.Print YamlConvert.toYaml(newDict)
    Set newDict = YamlConvert.toKeyValuePairs(" key1 : value1 , key2 : value2 ") 'Not a true valid key pair.
    Call assertBoolean(newDict.Count = 2, "comma unquoted")
    Debug.Print YamlConvert.toYaml(newDict)
    Set newDict = YamlConvert.toKeyValuePairs(" key1 : value1 # comment" & vbCrLf & " key2 : value2 ")
    Call assertBoolean(newDict.Count = 2, "vbcrlf, unquoted, and comment ")
    Debug.Print YamlConvert.toYaml(newDict)
    
    For Each sKey In newDict.keys()
        sItem = newDict.Item(sKey)
        Debug.Print "Result, " & sKey & "(" & sItem & ") "
    Next
    Debug.Print "----"
End Sub

Sub assertBoolean(result As Boolean, msg As String)
    If result Then
        Debug.Print msg & ". Pass"
    Else
        Debug.Print msg & "Expected (True), Actual(" & result & ")"
    End If
End Sub
