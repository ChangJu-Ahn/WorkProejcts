<% session.CodePage=949 %>

<%
'======================================================================================================================
' Name : CheckSYSTEMErrorAcct()
' Desc : Display SYSTEM error message
'======================================================================================================================

Function CheckSYSTEMErrorAcct(objError, pBool)

	Dim iDesc
	Dim iArrayCnt
	Dim pArg1,pArg2     
	pArg1=""
	pArg2=""
		
	CheckSYSTEMErrorAcct = False
	    
	If objError.Number = 0 Then
	   Exit Function
	End If

	CheckSYSTEMErrorAcct = True
	    
	If objError.Number = vbObjectError Then
	   If InStr(UCase(objError.Description), "B_MESSAGE") > 0 Then

	   		      
		    iDesc = Split(objError.Description, Chr(11))
		    iArrayCnt= Ubound(iDesc)
		    
		    Select Case iArrayCnt		  
			Case	2
		        If Mid(iDesc(1),1,5) = "REF00" Then  
					iDesc(1) =Mid(iDesc(1),6)
				End If		  
					pArg1=iDesc(2)

			Case	3
		        If Mid(iDesc(1),1,5) = "REF00" Then  
					iDesc(1) =Mid(iDesc(1),6)
				End If		  
				pArg1=iDesc(2)
				pArg2=iDesc(3)	  
			End Select  
			          			
			Call DisplayMsgBox(iDesc(1), vbOKOnly, pArg1, pArg2, I_MKSCRIPT)
			   Exit Function
	   End If
	End If

	If InStr(UCase(objError.Description), "B_CASE") > 0 Then
	   If HandleBCaseError(objError.Number, objError.Description, pArg1, pArg2) = True Then
	      Exit Function
	   End If
	End If	    
	       
	If ShowODBCErrorCode(objError.Number, objError.Description) = False Then
	   If pBool = True Then
	      Call SvrMsgBox(objError.Description & vbCrLf & "Error Code : " & objError.Number, vbCritical, I_MKSCRIPT)
	   End If
	End If
	    
	objError.Clear
    
End Function

'======================================================================================================================
' Name : CheckSYSTEMErrorAcct2()
' Desc : Display SYSTEM error message 
'======================================================================================================================
Function CheckSYSTEMErrorAcct2(objError, ByVal pBool, ByVal pArg1, ByVal pArg2, ByVal Opt1, ByVal Opt2, ByVal Opt3)
    Dim iDesc
    Dim iRet
	Dim iArrayCnt
	Dim iArg1,iArg2     
	iArg1=""
	iArg2=""

    CheckSYSTEMErrorAcct2 = False
    
    If objError.Number = 0 Then
       Exit Function
    End If
    
    CheckSYSTEMErrorAcct2 = True
    
    If InStr(UCase(objError.Description), "B_CASE") > 0 Then
       If HandleBCaseError(objError.Number, objError.Description, pArg1, pArg2) = True Then
          Exit Function
       End If
    End If
    
    If objError.Number = vbObjectError Then
       If InStr(UCase(objError.Description), "B_MESSAGE") > 0 Then
       		iDesc = Split(objError.Description, Chr(11))
		    iArrayCnt= Ubound(iDesc)
		    
		    Select Case iArrayCnt		  

			Case	2
		        If Mid(iDesc(1),1,5) = "REF00" Then  
					iDesc(1) =Mid(iDesc(1),6)
				End If		  
					iArg1=iDesc(2)

			Case	3
		        If Mid(iDesc(1),1,5) = "REF00" Then  
					iDesc(1) =Mid(iDesc(1),6)
				End If		  
				iArg1=iDesc(2)
				iArg2=iDesc(3)	  
			End Select  

			Call DisplayMsgBox(iDesc(1), vbOKOnly, pArg1 & "  " & iArg1 , pArg2 & "  " & iArg2 , I_MKSCRIPT)
			   Exit Function
			   
       End If
    End If
    
    CheckSYSTEMErrorAcct2 = CheckSYSTEMError(objError, pBool)
    
End Function

%>
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   