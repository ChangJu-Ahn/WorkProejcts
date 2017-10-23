<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 매출채권일괄등록 
'*  3. Program ID           : S5111BB1
'*  4. Program Name         : 
'*  5. Program Desc         : 매출채권관리 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/06/30
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
' =======================================================================================================
%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%
	Call LoadBasisGlobalInf()

	Call loadInfTB19029B("I", "*","NOCOOKIE","BB")

    Call HideStatusWnd                                                               '☜: Hide Processing message
     Call SubCreateCommandObject(lgObjComm)
     Call SubBizBatch()
     Call SubCloseCommandObject(lgObjComm)
  
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

    Dim intRetCD
    Dim iObjRs
    Dim iArrBillNo		' 추가된 매출채권번호 
    Dim iStrArFlag		' 확정여부 
	Dim iStrWorkType	' 작업유형 
	
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    IntRetCD = 0

	iStrArFlag = Request("txtArFlag")
	iStrWorkType = Request("txtWorkType")
	
    Set iObjRs = Server.CreateObject("ADODB.Recordset")

    With lgObjComm
		.CommandTimeout = 0
		' 등록 
		If iStrWorkType = "C" Then
			.CommandText = "dbo.usp_s_CreateBillByBatch"
		' 삭제 
		Else
			.CommandText = "dbo.usp_s_DeleteBillByBatch"
		End If
		
        .CommandType = adCmdStoredProc

	    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger,adParamReturnValue)
		.Parameters.Append .CreateParameter("@from_dt", adDBTimeStamp,adParamInput,,UNIConvDate(Request("txtConFromDt")))
	    .Parameters.Append .CreateParameter("@to_dt", adDBTimeStamp,adParamInput,,UNIConvDate(Request("txtConToDt")))

		If iStrWorkType = "C" Then
			If Trim(Request("txtConShipToParty")) <> "" Then
				.Parameters.Append .CreateParameter("@ship_to_party", adVarXChar,adParamInput,10,Replace(Request("txtConShipToParty"), "'", "''"))
			Else
				.Parameters.Append .CreateParameter("@ship_to_party", adVarXChar,adParamInput,10,"%")
			End If
		
			If Trim(Request("txtConMovType")) <> "" Then
			    .Parameters.Append .CreateParameter("@mov_type", adVarXChar,adParamInput,3,Replace(Request("txtConMovType"), "'", "''"))
			Else
			    .Parameters.Append .CreateParameter("@mov_type", adVarXChar,adParamInput,3,"%")
			End If
		
			If Trim(Request("txtConSalesGrp")) <> "" Then
			    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4,Replace(Request("txtConSalesGrp"), "'", "''"))
			Else
			    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4,"%")
			End If
		
			.Parameters.Append .CreateParameter("@bill_dt", adDBTimeStamp,adParamInput,,UNIConvDate(Request("txtBillDt")))
			.Parameters.Append .CreateParameter("@user_id", adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))

			' 세금계산서 자동발행여부 
			.Parameters.Append .CreateParameter("@vat_auto_flag", adXChar,adParamInput,1,Request("txtVatFlag"))

			' 출고처리를 할 경우 추가된 모든 D/N정보를 결과값으로 Return 한다.
			If iStrArFlag = "Y" Then
				.Parameters.Append .CreateParameter("@return_flag", adXChar,adParamInput,1,"A")
			Else
				.Parameters.Append .CreateParameter("@return_flag", adXChar,adParamInput,1,"R")
			End If
		Else
			If Trim(Request("txtConShipToParty")) <> "" Then
				.Parameters.Append .CreateParameter("@sold_to_party", adVarXChar,adParamInput,10,Replace(Request("txtConShipToParty"), "'", "''"))
			Else
				.Parameters.Append .CreateParameter("@sold_to_party", adVarXChar,adParamInput,10,"%")
			End If
		
			If Trim(Request("txtConSalesGrp")) <> "" Then
			    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4,Replace(Request("txtConSalesGrp"), "'", "''"))
			Else
			    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4,"%")
			End If
		
			.Parameters.Append .CreateParameter("@user_id", adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))
			.Parameters.Append .CreateParameter("@return_flag", adXChar,adParamInput,1,"R")
		End If
		
        Set iObjRs = .Execute
    End With
    
    If CheckSYSTEMError(Err,True) = True Then
       IntRetCD = -1
       Exit Sub
    End If
    
    IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
    
    If CDbl(intRetCD) = 0 Then
		iArrBillNo = iObjRs.GetRows
		
		iObjRs.Close
    	Set iObjRs = Nothing

		If iStrArFlag = "N" Then
			If iStrWorkType = "C" Then
				Call DisplayMsgBox("204262", vbOKOnly, iArrBillNo(0, 0) & "~" & iArrBillNo(1, 0) & " (" & iArrBillNo(2, 0) & ")", "", I_MKSCRIPT)
			Else
				Call DisplayMsgBox("204268", vbOKOnly, iArrBillNo(0, 0) & "~" & iArrBillNo(1, 0) & " (" & iArrBillNo(2, 0) & ")", "", I_MKSCRIPT)
			End If
		Else
			Call DisplayMsgBox("204262", vbOKOnly, iArrBillNo(0, 0) & "~" & iArrBillNo(0, UBound(iArrBillNo, 2)) & " (" & UBound(iArrBillNo, 2) + 1 & ")", "", I_MKSCRIPT)
			
			' 추가된 매출에 대해 확정처리 
			Call PostBill(iArrBillNo)
	    End If
    Else
       Call DisplayMsgBox(IntRetCd, vbInformation, "", "", I_MKSCRIPT)
       If Not(iObjRs Is Nothing) then
			Set iObjRs = Nothing
       End If
    End If
    
	Response.Write  " <Script Language=vbscript> " & vbCr
	Response.Write  "  Call Parent.SetFocusToDocument(""M"")  " & vbCr
	Response.Write  "  Call Parent.frm1.txtConFromDt.focus  " & vbCr
	Response.Write  " </Script>                  " & vbCr
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
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'=========================================================================================
' Post Billing
'=========================================================================================
Sub PostBill(ByRef prArrBillNo)
	On Error Resume Next

	Dim iIntLoop, iIntLastRow
	Dim pvCB
	Dim iObjPS7G115

	pvCB = "F" 	   

	iIntLastRow = UBound(prArrBillNo, 2)

	Set iObjPS7G115 = Server.CreateObject("PS7G115.cSPostOpenArSvr")

	If CheckSYSTEMError(Err,True) = True Then		                                                 '☜: Unload Comproxy DLL
       Exit Sub
    End If  

	For iIntLoop = 0 To iIntLastRow
		
	    Call iObjPS7G115.S_POST_OPEN_AR_SVR(pvCB, gStrGlobalCollection, prArrBillNo(0, iIntLoop))
		    
		If CheckSYSTEMError2(Err, True, "(매출채권번호 : " & prArrBillNo(0, iIntLoop) & ")","","","","") = True Then
			Set iObjPS7G115 = Nothing
			' 일부만 처리 된 경우 처리된 정보를 보여준다.
			If iIntLoop > 0 Then
				Call DisplayMsgBox("204267", vbOKOnly, prArrBillNo(0, 0) & "~" & prArrBillNo(0, iIntLoop - 1) & " (" & iIntLastRow & ")", "", I_MKSCRIPT)
			End If
			
			Exit Sub
		End If
	Next

	Set iObjPS7G115 = Nothing
	
	Call DisplayMsgBox("204267", vbOKOnly, prArrBillNo(0, 0) & "~" & prArrBillNo(0, iIntLastRow) & " (" & iIntLastRow + 1 & ")", "", I_MKSCRIPT)
End Sub

%>

