Public Function ExtractJSONValue(strFieldName As String, strJSON As String, Optional StrFieldName2 As String, Optional outputAsNumeric As Boolean, Optional eventName As String)

    Dim strWrkStr As String
    Dim LenStrFieldNm As Long, LenStrJSON As Long, PosStrFldNm As Long
    
    strWrkStr = ""
    
    On Error GoTo Errhandler
    
    '----------------------- Event message based parsing -----------------------------
    
    If Len(eventName) > 0 Then
        '---- For LOGIN events request message will contain only Email ----
        If eventName = "LOGIN" And StrFieldName2 = "Email" Then
            ExtractJSONValue = strJSON
            Exit Function
        End If
        
        '----- For IDENTITY_CHECK failures the response message will contian the error message ---
        If eventName = "IDENTITY_CHECK" And StrFieldName2 = "errorDescription" Then
            ExtractJSONValue = strJSON
            Exit Function
        End If

    End If
    
    '---------------------------------------------------------------------------------
    LenStrFieldNm = Len(strFieldName)
    LenStrJSON = Len(strJSON)
    PosStrFldNm = InStr(strJSON, strFieldName & Chr(34))
    
    '--- If 1st Field name search not found then search for 2nd Field name if provided in the input ---
    If PosStrFldNm = 0 Then
        If Len(StrFieldName2) > 0 Then
            LenStrFieldNm = Len(StrFieldName2)
            PosStrFldNm = InStr(strJSON, StrFieldName2 & Chr(34))
        End If
        
        If PosStrFldNm = 0 Then
            If outputAsNumeric = True Then
                ExtractJSONValue = 0
                Exit Function
            Else
                ExtractJSONValue = ""                                                       'ExtractJSONValue = "Field not present"
                Exit Function
            End If
        End If
    End If
    
    strWrkStr = Right(strJSON, LenStrJSON - (PosStrFldNm + LenStrFieldNm + 1))
    
    Dim chr34Pos As Integer         ' Char34 is "
    Dim chr125Pos As Integer        ' Char125 is }
    
    If Left(strWrkStr, 1) = Chr(34) Then                                                '   Determines if value is a number or nullvalue
    
        chr34Pos = InStr(strWrkStr, Chr(34) & ",")
        chr125Pos = InStr(strWrkStr, Chr(34) & "}")
        
        If (chr34Pos < chr125Pos And chr34Pos > 0) Or (chr34Pos > 0 And chr125Pos = 0) Then       ' If "," delimiter comes first as delimiter
            strWrkStr = Left(strWrkStr, InStr(strWrkStr, Chr(34) & ",") - 1)
            strWrkStr = Right(strWrkStr, Len(strWrkStr) - 1)
        Else                                                                                      ' Else use "}" as the delimiter
            strWrkStr = Left(strWrkStr, InStr(strWrkStr, Chr(34) & "}") - 1)
            strWrkStr = Right(strWrkStr, Len(strWrkStr) - 1)
        End If

    Else
        If InStr(strWrkStr, ",") = 0 Then
            strWrkStr = Left(strWrkStr, InStr(strWrkStr, "}") - 1)
        Else
            strWrkStr = Left(strWrkStr, InStr(strWrkStr, ",") - 1)
        End If
    End If
    
    If outputAsNumeric = True Then
       If Len(strWrkStr) = 0 Or strWrkStr = "null" Then
          ExtractJSONValue = 0
          Exit Function
       End If
       ExtractJSONValue = strWrkStr * 1
    Else
        ExtractJSONValue = strWrkStr
    End If
    Exit Function
    
    '---- Error Handler-----------
Errhandler:
    If outputAsNumeric = True Then
       ExtractJSONValue = 0
    Else
       ExtractJSONValue = ""
    End If

End Function



'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------


Public Function ExtractXMLValue(strFieldName As String, strXML As String, Optional strFieldName2 As String, Optional outputAsNumeric As Boolean)
    
    Dim strWrk As String
    Dim LenStrFieldNm As Long, LenStrXML As Long, startPos As Long, endPos As Long, dataLen As Integer
    
    strWrk = ""
    
    On Error GoTo ErrHandler
    
    '---- 1. Search if first field exists in the xml data -----
    LenStrXML = Len(strXML)
    LenStrFieldNm = Len(strFieldName)
    startPos = InStr(strXML, "<" & strFieldName & ">")
    endPos = InStr(strXML, "</" & strFieldName & ">")
    '---- 2. If first field does not exist check if 2nd field exists in the xml data----
    If startPos = 0 Then
        If Len(strFieldName2) > 0 Then
            LenStrFieldNm = Len(strFieldName2)
            startPos = InStr(strXML, "<" & strFieldName2 & ">")
            endPos = InStr(strXML, "</" & strFieldName2 & ">")
        End If
        
        If startPos = 0 Then                 '-- Both fields not present
            If outputAsNumeric = True Then
                ExtractXMLValue = 0
                Exit Function
            Else
                ExtractXMLValue = ""
                Exit Function
            End If
        End If
    End If

    '----3. Find the length of the XML field data --------
    startPos = startPos + LenStrFieldNm + 2
    dataLen = endPos - startPos
    strWrk = Mid(strXML, startPos, dataLen)
    
    If outputAsNumeric = True Then
       If Len(strWrk) = 0 Or strWrk = "null" Then
          ExtractXMLValue = 0
       Else
          ExtractXMLValue = strWrk * 1
       End If
    Else
        ExtractXMLValue = strWrk
    End If
    
    Exit Function

ErrHandler:

    If outputAsNumeric = True Then
       ExtractXMLValue = 0
    Else
       ExtractXMLValue = ""
    End If
    
End Function