<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : ����ä���ϰ���� 
'*  3. Program ID           : S5111BB1
'*  4. Program Name         : 
'*  5. Program Desc         : ����ä�ǰ��� 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/06/30
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
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

    Call HideStatusWnd                                                               '��: Hide Processing message
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
    Dim iArrDnNo		' �߰��� ����ȣ�� ������ �迭 
    Dim iStrWorkType	' �۾�����('C' : ����, 'D' : ����)
    Dim iStrGiFlag		' ���ó�� ���� 

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    IntRetCD = 0

	iStrWorkType = Request("txtWorkType")
	iStrGiFlag = Request("txtGiFlag")
	
    Set iObjRs = Server.CreateObject("ADODB.Recordset")

    With lgObjComm
		.CommandTimeout = 0
		
		If iStrWorkType = "C" Then
			.CommandText = "dbo.usp_s_CreateDnByBatch"		' ��� 
        Else
			.CommandText = "dbo.usp_s_DeleteDnByBatch"		' ���� 
        End If
        
		.CommandType = adCmdStoredProc

	    .Parameters.Append .CreateParameter("RETURN_VALUE", adInteger,adParamReturnValue)
		.Parameters.Append .CreateParameter("@plant_cd", adVarXChar,adParamInput,4,FilterVar(Request("txtConPlant"), "''", "S"))
		.Parameters.Append .CreateParameter("@fr_promise_dt", adDBTimeStamp,adParamInput,,UNIConvDate(Request("txtConFromDt")))
	    .Parameters.Append .CreateParameter("@to_promise_dt", adDBTimeStamp,adParamInput,,UNIConvDate(Request("txtConToDt")))

		If Trim(Request("txtConMovType")) <> "" Then
		    .Parameters.Append .CreateParameter("@mov_type", adVarXChar,adParamInput,3,FilterVar(Request("txtConMovType"), "''", "S"))
		Else
		    .Parameters.Append .CreateParameter("@mov_type", adVarXChar,adParamInput,3,"%")
		End If
		
		If Trim(Request("txtConShipToParty")) <> "" Then
			.Parameters.Append .CreateParameter("@ship_to_party", adVarXChar,adParamInput,10,FilterVar(Request("txtConShipToParty"), "''", "S"))
		Else
			.Parameters.Append .CreateParameter("@ship_to_party", adVarXChar,adParamInput,10,"%")
		End If
		
		If Trim(Request("txtConSalesGrp")) <> "" Then
		    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4,FilterVar(Request("txtConSalesGrp"), "''", "S"))
		Else
		    .Parameters.Append .CreateParameter("@sales_grp", adVarXChar,adParamInput,4,"%")
		End If

		' ������ ��쿡�� �ʿ� ����.
		If iStrWorkType = "C" Then
			.Parameters.Append .CreateParameter("@promise_dt", adDBTimeStamp,adParamInput,,UNIConvDate(Request("txtPromiseDt")))
			.Parameters.Append .CreateParameter("@trans_meth", adVarXChar,adParamInput,5,Replace(Request("txtTransMeth"), "'", "''"))
		End If
		
	    .Parameters.Append .CreateParameter("@user_id", adVarXChar,adParamInput,13,Replace(Request("txtUserId"), "'", "''"))
	    
		' ���ó���� �� ��� �߰��� ��� D/N������ ��������� Return �Ѵ�.
	    If iStrGiFlag = "Y" Then
			.Parameters.Append .CreateParameter("@return_flag", adXChar,adParamInput,1,"A")
		Else
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
		iArrDnNo = iObjRs.GetRows
		
		iObjRs.Close
		Set iObjRs = Nothing
		
		If iStrGiFlag = "N" Then
			If iStrWorkType = "C" Then
				Call DisplayMsgBox("204262", vbOKOnly, iArrDnNo(0, 0) & "~" & iArrDnNo(1, 0) & " (" & iArrDnNo(2, 0) & ")", "", I_MKSCRIPT)
			Else
				Call DisplayMsgBox("204268", vbOKOnly, iArrDnNo(0, 0) & "~" & iArrDnNo(1, 0) & " (" & iArrDnNo(2, 0) & ")", "", I_MKSCRIPT)
			End If
		Else
			Call DisplayMsgBox("204262", vbOKOnly, iArrDnNo(0, 0) & "~" & iArrDnNo(0, UBound(iArrDnNo, 2)) & " (" & UBound(iArrDnNo, 2) + 1 & ")", "", I_MKSCRIPT)
			
			' �߰��� ���Ͽ� ���� ���ó���� 
			Call PostGi(iArrDnNo)
	    End If
    Else
       Call DisplayMsgBox(IntRetCd, vbInformation, "", "", I_MKSCRIPT)
       If Not(iObjRs Is Nothing) then
			Set iObjRs = Nothing
       End If
    End If
    
	Response.Write  " <Script Language=vbscript> " & vbCr
	Response.Write  "  Call Parent.frm1.txtConPlant.focus  " & vbCr
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
' Post Goods Issue
'=========================================================================================
Sub PostGi(ByRef prArrDnNo)
	On Error Resume Next

	Dim iIntLoop, iIntLastRow
	Dim pvCB
	Dim iStrCommand			
	Dim iArrPostInfo
	Dim iObjPS5G115

	Redim iArrPostInfo(5)
		
	' ��� Ȯ������ ���� ���� 
	iArrPostInfo(1) = UNIConvDate(Request("txtActualGiDt"))	' ���� ����� 
	iArrPostInfo(2) = Trim(Request("txtArFlag"))			' ����������� 
	iArrPostInfo(3) = Trim(Request("txtVatFlag"))			' ���ݰ�꼭 �������� 
	iArrPostInfo(4) = "N"									' ��������� 
	iArrPostInfo(5) = "ST"									' STO ���� 
	
	pvCB = "F" 	   
	iStrCommand = "POST"					' �׻� �빮�� 

	iIntLastRow = UBound(prArrDnNo, 2)

	Set iObjPS5G115 = CreateObject("PS5G115.cSPOSTGISvr")

	If CheckSYSTEMError(Err,True) = True Then		                                                 '��: Unload Comproxy DLL
       Exit Sub
    End If  

	For iIntLoop = 0 To iIntLastRow
		
	    iArrPostInfo(0) = prArrDnNo(0, iIntLoop)					' ����ȣ 
	    Call iObjPS5G115.S_POST_GOODS_ISSUE_SVR(pvCB, gStrGlobalCollection, iStrCommand, Array(""), iArrPostInfo)
		    
		If CheckSYSTEMError2(Err, True, "(����ȣ : " & prArrDnNo(0, iIntLoop) & ")","","","","") = True Then
			Set iObjPS5G115 = Nothing
			' �Ϻθ� ó�� �� ��� ó���� ������ �����ش�.
			If iIntLoop > 0 Then
				Call DisplayMsgBox("204267", vbOKOnly, prArrDnNo(0, 0) & "~" & prArrDnNo(0, iIntLoop - 1) & " (" & iIntLastRow & ")", "", I_MKSCRIPT)
			End If
			
			Exit Sub
		End If
	Next

	Set iObjPS5G115 = Nothing
	
	Call DisplayMsgBox("204267", vbOKOnly, prArrDnNo(0, 0) & "~" & prArrDnNo(0, iIntLastRow) & " (" & iIntLastRow + 1 & ")", "", I_MKSCRIPT)
End Sub
%>

