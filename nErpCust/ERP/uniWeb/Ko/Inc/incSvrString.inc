<%

'========================================================================================
' Trim string and set string to space if string length is zero
' pData   : target data
' pStrALT : alternative string if space
' pOpt    :  S is for String
'            D is for Digit
' History : Appended in 2002/08/07 (lee jin soo)
'========================================================================================
Function FilterVar(ByVal pData, ByVal pStrALT, ByVal pOpt)

     If IsNull(pData) Then
        pData = "" 
     Else   
        pData = Trim(pData)
     End If       
     
     pOpt = UCase(pOpt)
     
     Select Case VarType(pData)
        Case vbEmpty                                           '0    Empty (uninitialized)
                 FilterVar = pStrALT
                 Exit Function
        Case vbNull                                            '1    Null (no valid data)
                 FilterVar = "Null"
                 Exit Function
        Case vbInteger, vbLong, vbSingle, vbDouble             '2(Integer),3(Long integer),4(Single-precision floating-point number),5(Double-precision floating-point number)
                 FilterVar = pData
                 Exit Function
        Case vbCurrency, vbBoolean, vbByte                     '6(Currency),11(Boolean),17(Byte)
                 FilterVar = pData
                 Exit Function
        Case Else
        
                 If pData = "" Then
                    
                    If pOpt = "S" And Trim(pStrALT) = "" Then
                       pStrALT = "''"
                    End If
                    
                    If pOpt = "S2" And Trim(pStrALT) = "" Then
                       pStrALT = "''''"
                    End If
                    
                    If gCharSQLSet = "U" Then
                       If Len(pStrALT) > 1 Then
                          If Mid(pStrALT, 1, 2) = "N'" Then
                             pStrALT = Mid(pStrALT, 2)
                          End If
                       End If
                       
                       If pOpt = "S" Then
                       
                          If IsNull(pStrALT) Or UCase(Trim(pStrALT)) = "NULL" Then
                          Else
                             pStrALT = "N" & pStrALT
                          End If
                       
                       End If
                    
                    End If
                    
                    FilterVar = pStrALT
                    
                    Exit Function
                 End If
     
                 Select Case pOpt
                     Case "S"
                                pData = Replace(pData, "'", "''")
                                If gCharSQLSet = "U" Then
                                   FilterVar = "N'" & pData & "'"
                                Else
                                   FilterVar = "'" & pData & "'"
                                End If
                     Case "S2"
                                pData = Replace(pData, "'", "''")
                                If gCharSQLSet = "U" Then
                                   FilterVar = "N''" & pData & "''"
                                Else
                                   FilterVar = "''" & pData & "''"
                                End If
                     Case "SNM"
                                FilterVar = Replace(pData, "'", "''")
                     Case Else
                                FilterVar = pData
                 End Select
     End Select
     
End Function


'========================================================================================
' Function Name : ConvSPChars
' Function Desc : 문자열안의 "를 ""로 바꾼다.
'========================================================================================
Function ConvSPChars(strVal)
	ConvSPChars = Replace("" & strVal, """", """""")
End Function 

'========================================================================================
' Function Name : ConvertTypChaStr
' Function Desc : 
'========================================================================================
function ConvertTypChaStr(strText) 
  Dim strInput,stroutput,strTmp
  Dim nCount, i 
    
  strInput = strText
  nCount = Len(strInput)
  strOutput = ""
  
  For i = 0 To nCount
  
     strTmp = Mid(strInput, i + 1, 1)
     Select Case (strTmp)
     
          Case "<":
               strOutput = strOutput + "&lt;"
          Case ">":
               strOutput = strOutput + "&gt;"
          Case "&":
               strOutput = strOutput + "&amp;"
          Case Chr(34):
               strOutput = strOutput + "&quot;"
          Case Chr(39):
               strOutput = strOutput + "&#39;"
          Case Chr(13):
               strOutput = strOutput + ""
          Case Else:
               strOutput = strOutput + strTmp
    End Select
     
  Next
  Convert = strOutput
  
end function


%>
