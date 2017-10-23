<%@ LANGUAGE=VBSCript%>
<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->

<%
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("Q","A", "COOKIE", "QB")
 
    Dim lgFiscStart
	Dim txtStartDt
	Dim txtEndDt
	Dim txtClassType
	Dim txtBizArea
	Dim txtPrintOpt
	Dim strZeroFg
	Dim strHqBrchFg
	Dim strUserId
	Dim strSPID
    
	Dim lgPreDayAmt
	Dim lgNowDayAmt
	
	Dim yyyy,mm,dd

	' ���Ѱ��� �߰� 
	Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
	Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
	Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
	Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

	Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					
	
    
    Call HideStatusWnd                                                               '��: Hide Processing message
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '��: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '��: Read Operation Mode (CRUD)
    
	txtStartDt = uniconvdate(Request("txtStartDt"))
	txtEndDt = uniconvdate(Request("txtEndDt"))
	txtClassType	= Trim(Request("txtClassType"))
	txtBizArea		= Trim(Request("txtBizArea"))
	strUserId		= Trim(Request("strUserId"))
	strSpid	    	= Trim(Request("strSpid"))
	strZeroFg		= Trim(Request("strZeroFg"))
	
	
	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))		
	lgInternalCd		= Trim(Request("lgInternalCd"))	
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))	
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))


	Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
    Call SubBizBatch()   
    Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection

'============================================================================================================
' Name : SubBizbatch
' Desc : Batch
'============================================================================================================
Sub SubBizBatch()
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call SubCreateCommandObject(lgObjComm)
    Call SubBizBatchMulti()                            '��: Run Batch
    Call SubCloseCommandObject(lgObjComm)


    If lgErrorStatus    = "YES" Then
  '     lgErrorPos = lgErrorPos & arrColVal(1) & gColSep         
    End If
    
    IF lgErrorStatus = "NO"	Then
    		'Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)		'�۾��� �Ϸ�Ǿ����ϴ� 
	END IF
End Sub


'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti()
	 On Error Resume NEXT
	 Err.Clear
	 
    Dim IntRetCD
    Dim strMsg_cd, strMsg_text
    Dim strSp
    
    Dim  strNativeErr   
    
    strSp = "USP_A_DMS"

    lgstrData = ""
    strNativeErr = ""

	'���� ���� �߰� 
	Dim BizAreaCd
	
	BizAreaCd = txtBizArea

	If BizAreaCd = "" Then	 	 	
		If lgAuthBizAreaCd <> "" Then			
			BizAreaCd = lgAuthBizAreaCd
		End If
	Else
		If lgAuthBizAreaCd <> "" Then
			If UCASE(lgAuthBizAreaCd) <> UCASE(BizAreaCd) Then
		        Call DisplayMsgBox("124200", vbInformation, "", "", I_MKSCRIPT)
				Response.end
			End If
		End If
	End If

	'2003/12/01 Oh Soo Min ���� 
	If Trim(BizAreaCd) = "" Then
		BizAreaCd = "%"
	End If 
	 
	With lgObjComm
	   .CommandText = strSp
	   .CommandType = adCmdStoredProc		

	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	 adInteger,	adParamReturnValue)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@start_dt",		 adWChar,	adParamInput,		10, txtStartDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@end_dt",			 adWChar,	adParamInput,		10, txtEndDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@class_type",		 adVarWChar,	adParamInput,		20, txtClassType)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@biz_area_cd",	 adVarWChar,	adParamInput,		10, BizAreaCd)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@zero_fg",		 adWChar,	adParamInput,		1, strZeroFg)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id",		 adVarWChar,	adParamInput,		13, strUserId)	   	   
	   
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pre_day_amt",	 adVarWChar,	adParamOutput,		20)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@now_day_amt",	 adVarWChar,	adParamOutput,		20)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",		     adVarWChar,	adParamOutput,		6)	   
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@spid",		     adVarWChar,	adParamOutput,		13)
	   
	   lgObjComm.Execute ,, adExecuteNoRecords	
	End With

    If Err.number = 0 Then
       IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
		'Response.Write "IntRetCD=" & intRetCd
       If IntRetCD <> 1 then
          strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
          if strMsg_Cd <> "" Then
		       Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
		       Call SubCloseCommandObject(lgObjComm)
			   Call HideStatusWnd                                                               '��: Hide Processing message
  		  END IF
          Response.end          

       Else	
		strSPID		    = lgObjComm.Parameters("@spid").Value
		lgPreDayAmt		= lgObjComm.Parameters("@pre_day_amt").Value
		lgNowDayAmt		= lgObjComm.Parameters("@now_day_amt").Value
		
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtSPID.value = "<%=ConvSPChars(strSPID)%>"
			.txtYAmt.value = "<%=UNINumClientFormat(lgPreDayAmt,  ggAmtOfMoney.DecPoint, 0)%>"							     
			.txtTAmt.value = "<%=UNINumClientFormat(lgNowDayAmt,  ggAmtOfMoney.DecPoint, 0)%>"
			.txtOUT.value = "1"
			End With
		</Script>
<%		

       End If
        
   Else    
   
	  lgErrorStatus     = "YES"
	  Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
	  Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	     
   
      lgErrorStatus     = "YES"
      If lgObjComm.ActiveConnection.Errors.Count > 0 then
			strNativeErr = lgObjComm.ActiveConnection.Errors(0).NativeError
	  End If
		
		Select Case Trim(strNativeErr)
			Case "8115"																'%1!��(��) ������ ���� %2!(��)�� ��ȯ�ϴ� �� ��� �����÷� ������ �߻��߽��ϴ�.
				Call DisplayMsgBox("121515", vbInformation, "", "", I_MKSCRIPT)
			Case Else				
				Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
		End Select	  
	        
      Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
   End if

End Sub


'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear
    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub

%>

<Script Language="VBScript">
	
	If Trim("<%=lgErrorStatus%>") = "NO" Then
'		Parent.ggoSpread.Source  = Parent.frm1.vspdData2
'		Parent.ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data			
		Parent.DbQueryOk		
    End If

</Script>
