<%@ LANGUAGE=VBSCript%>
<!--'********************************************************************************************************
'*  1. Module Name          : 구매																		*
'*  2. Function Name        : L/C관리																	*
'*  3. Program ID           : M3212PA1																	*
'*  4. Program Name         : B/L내역팝업																*
'*  5. Program Desc         : 수입진행현황조회를 위한 B/L내역팝업 *
'*  7. Modified date(First) : 2003/07/01																*
'*  8. Modified date(Last)  :           																*
'*  9. Modifier (First)     : Lee Eun hee																*
'* 10. Modifier (Last)      :           
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 												*
'*				            : 												*
'*				            : 												*
'********************************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->

<%		
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M","NOCOOKIE","PB")   
	Call LoadBNumericFormatB("Q","M","NOCOOKIE","PB")
	Call HideStatusWnd 
												
	On Error Resume Next
	Err.Clear

	         
	Dim strBlNo, lgCurrency

	Dim IntRetCD
	Dim iEndRow
	Dim iPrevEndRow
	
	Dim strBpNm
	Dim strDocAmt
	Dim strIssueDt
	'---------------------------------------Common-----------------------------------------------------------

	lgLngMaxRow       = Cint(Request("txtMaxRows"))
	lgMaxCount        = 100
	lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)   
	lgErrorStatus     = "NO"
	lgErrorPos        = ""
	iPrevEndRow = 0
    iEndRow = 0
						
	Call SubOpenDB(lgObjConn)
	Call SubCreateCommandObject(lgObjComm)


	strBlNo		= FilterVar(Trim(UCase(Request("txtBLNo"))), " " , "SNM")
	lgCurrency	= UCase(Request("txtCurrency"))
	
	If lgStrPrevKeyIndex = 0 Then
		Call SubBizQueryHdr()
	End If
	
	Call SubBizQuery()

	Call SubCloseCommandObject(lgObjComm)    
	Call SubCloseDB(lgObjConn)      

'============================================================================================================
' Name : SubBizQueryHdr
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryHdr()
	
	On Error Resume Next
    Err.Clear
        
	lgStrSQL =	" SELECT B.BP_NM, A.DOC_AMT, convert(char(10), A.BL_ISSUE_DT, 20) " & _
				" FROM M_BL_HDR A, B_BIZ_PARTNER B " & _
				" WHERE A.BENEFICIARY = B.BP_CD "	& _
				"   AND A.BL_NO = '" & strBlNo	& "' " 
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus()
		
		Call SubCloseRs(lgObjRs)

		Response.End 
	Else
		
		strBpNm			= lgObjRs(0)
		strDocAmt		= lgObjRs(1)
		strIssueDt		= lgObjRs(2)
			        
    End If
	
    Call SubCloseRs(lgObjRs)                                             
		
	
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	
	Dim iDx
	Dim PvArr
	
	On Error Resume Next
    Err.Clear
    
    If CInt(lgStrPrevKeyIndex) > 0 Then
       iPrevEndRow = lgMaxCount * CInt(lgStrPrevKeyIndex)
    End If
        
	Call SubMakeSQLStatements()
	
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		IntRetCD = -1
		lgStrPrevKeyIndex = ""		
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus()
		
		Call SubCloseRs(lgObjRs)

		Exit Sub
	Else
			
		IntRetCD = 1
		 Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)
        lgstrData = ""
        iDx       = 1
        ReDim PvArr(0)
                
        Do While Not lgObjRs.EOF
			ReDim Preserve PvArr(iDx - 1)
            lgstrData =		Chr(11) & ConvSPChars(lgObjRs(0)) & _
							Chr(11) & ConvSPChars(lgObjRs(1)) & _
							Chr(11) & ConvSPChars(lgObjRs(2)) & _
							Chr(11) & ConvSPChars(lgObjRs(3)) & _
							Chr(11) & UniNumClientFormat(lgObjRs(4),ggQty.DecPoint,0) & _
							Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs(5),0) & _
							Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs(6),0) & _
							Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs(7),0) & _
							Chr(11) & UniNumClientFormat(lgObjRs(8),ggQty.DecPoint,0) & _
							Chr(11) & UniNumClientFormat(lgObjRs(9),ggQty.DecPoint,0) & _
							Chr(11) & ConvSPChars(lgObjRs(10)) & _
							Chr(11) & ConvSPChars(lgObjRs(11)) & _
							Chr(11) & ConvSPChars(lgObjRs(12)) & _
							Chr(11) & ConvSPChars(lgObjRs(13)) & _
							Chr(11) & ConvSPChars(lgObjRs(14)) & _
							Chr(11) & ConvSPChars(lgObjRs(15)) & _
							Chr(11) & ConvSPChars(lgObjRs(16))
		if UniNumClientFormat(lgObjRs(22),ggQty.DecPoint,0) > 0 then
			lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs(17),0) & _
									Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs(18),0)
		else'발주tolerance로 수정(2003.07.01)
			lgstrData = lgstrData & Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs(19),0) & _
									Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs(20),0)
		end if
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(21)) & _
									Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)
			
			PvArr(iDx - 1) = lgstrData            

		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If   
               
        Loop 
			lgstrData = Join(PvArr, "")
        
    End If
    If iDx <= lgMaxCount Then
		iEndRow = iPrevEndRow + iDx -1
       lgStrPrevKeyIndex = ""
    Else
		iEndRow = iPrevEndRow + iDx -1
    End If   
	
    Call SubCloseRs(lgObjRs)                                             
		
End Sub	

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements()

    On Error Resume Next
    Err.Clear
    
	lgStrSQL =	" SELECT A.ITEM_CD, B.ITEM_NM, B.SPEC, A.UNIT, A.QTY, A.PRICE, A.DOC_AMT, A.LOC_AMT, " & _
				" A.GROSS_WEIGHT, A.VOLUME_SIZE, A.HS_CD, C.HS_NM, A.BL_SEQ, A.PO_NO, A.PO_SEQ_NO, " & _
				" A.LC_NO, A.LC_SEQ, E.OVER_TOLERANCE, E.UNDER_TOLERANCE, D.OVER_TOL, D.UNDER_TOL, A.TRACKING_NO, D.LC_QTY " & _
				" FROM M_BL_DTL A " & _
				" LEFT OUTER JOIN M_PUR_ORD_DTL D ON (A.PO_NO = D.PO_NO AND A.PO_SEQ_NO = D.PO_SEQ_NO) " & _
				" LEFT OUTER JOIN M_LC_DTL E ON (A.LC_NO = E.LC_NO AND A.LC_SEQ = E.LC_SEQ), B_ITEM B, B_HS_CODE C " & _
				" WHERE A.ITEM_CD = B.ITEM_CD AND A.HS_CD = C.HS_CD      "	& _
				"	AND A.BL_NO			= '" & strBlNo	& "' " & _
				"	ORDER BY A.BL_SEQ "
			

End Sub    


'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub


'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear

    Select Case pOpCode
        Case "MC"
            If CheckSYSTEMError(pErr,True) = True Then
               ObjectContext.SetAbort
               Call SetErrorStatus
            Else
               If CheckSQLError(pConn,True) = True Then
                  ObjectContext.SetAbort
                  Call SetErrorStatus
               End If
            End If
        Case "MD"
        Case "MR"
        Case "MU"
            If CheckSYSTEMError(pErr,True) = True Then
               ObjectContext.SetAbort
               Call SetErrorStatus
            Else
               If CheckSQLError(pConn,True) = True Then
                  ObjectContext.SetAbort
                  Call SetErrorStatus
               End If
            End If
        Case "MB"
			ObjectContext.SetAbort
            Call SetErrorStatus        
    End Select
End Sub

%>
<Script Language="VBScript">
	With parent		
		If "<%=lgErrorStatus%>" = "NO" And "<%=IntRetCd%>" <> -1 Then
		    .ggoSpread.Source	= .frm1.vspdData
			.lgStrPrevKeyIndex	= "<%=lgStrPrevKeyIndex%>"
			
			If "<%=iPrevEndRow%>" = 0 Then
				.CurFormatNumericOCX
				.frm1.txtBeneficiaryNm.value = "<%=ConvSPChars(strBpNm)%>"
				.frm1.txtIssueDt.text = "<%=UNIDateClientFormat(strIssueDt)%>"
				.frm1.txtDocAmt.text = "<%=UNIConvNumDBToCompanyByCurrency(strDocAmt,lgCurrency,ggAmtOfMoneyNo,"X","X")%>"
			End if
			
			.frm1.vspdData.ReDraw = False
			.ggoSpread.SSShowData "<%=lgstrData%>", "F"
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>","<%=iEndRow%>","<%=lgCurrency%>",.C_Price,"C","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>","<%=iEndRow%>","<%=lgCurrency%>",.C_DocAmt,"A","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>","<%=iEndRow%>",gCurrency,.C_LocAmt,"A","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>","<%=iEndRow%>","<%=lgCurrency%>",.C_OverTolerance,"D","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>","<%=iEndRow%>","<%=lgCurrency%>",.C_UnderTolerance,"D","Q","X","X")
			
			.DbQueryOk
			.frm1.vspdData.focus
			.frm1.vspdData.ReDraw = True
		Else
			If "<%=iPrevEndRow%>" = 0 Then
				.CurFormatNumericOCX
				.frm1.txtBeneficiaryNm.value = "<%=ConvSPChars(strBpNm)%>"
				.frm1.txtIssueDt.text = "<%=UNIDateClientFormat(strIssueDt)%>"
				.frm1.txtDocAmt.text = "<%=UNIConvNumDBToCompanyByCurrency(strDocAmt,lgCurrency,ggAmtOfMoneyNo,"X","X")%>"
			End if
		End If
	End With	
</Script>	


	
