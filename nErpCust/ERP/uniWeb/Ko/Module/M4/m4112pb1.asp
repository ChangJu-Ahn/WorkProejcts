<%@ LANGUAGE=VBSCript%>
<!--'********************************************************************************************************
'*  1. Module Name          : 구매																		*
'*  2. Function Name        : L/C관리																	*
'*  3. Program ID           : M4212PA1																	*
'*  4. Program Name         : 통관내역팝업																*
'*  5. Program Desc         : 수입진행현황조회를 위한 통관내역팝업 *
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

	         
	Dim strRcptNo, lgCurrency

	Dim IntRetCD
	Dim iEndRow
	Dim iPrevEndRow

	Dim strMvmtDt
	Dim strMvmtType
	Dim strMvmtTypeNm
	Dim strBpNm
	Dim strPurGrp
	Dim strPurGrpNm
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


	strRcptNo		= FilterVar(Trim(UCase(Request("txtMvmtNo"))), " " , "SNM")
	
	Call SubBizQuery()

	Call SubCloseCommandObject(lgObjComm)    
	Call SubCloseDB(lgObjConn)      
	
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

		Response.End 
	Else
		If iPrevEndRow = 0 Then
			strMvmtDt		= lgObjRs(34)
			strMvmtType		= lgObjRs(32)
			strMvmtTypeNm	= lgObjRs(33)
			strBpNm			= lgObjRs(36)
			strPurGrp		= lgObjRs(37)
			strPurGrpNm		= lgObjRs(38)
		End if
		
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
							Chr(11) & ConvSPChars(lgObjRs(4)) & _
							Chr(11) & ConvSPChars(lgObjRs(5))
			If ConvSPChars(lgObjRs(6)) = "Y" Then
			lgstrData =	lgstrData &	Chr(11) & "1"
			Else
			lgstrData =	lgstrData &	Chr(11) & "0"
			End If
			
			lgstrData =	lgstrData &	Chr(11) & UniNumClientFormat(lgObjRs(7),ggQty.DecPoint,0) & _
							Chr(11) & UniNumClientFormat(lgObjRs(8),ggQty.DecPoint,0) & _
							Chr(11) & ConvSPChars(lgObjRs(9)) & _
							Chr(11) & ConvSPChars(lgObjRs(10)) & _
							Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs(11),0) & _
							Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs(12),0) & _
							Chr(11) & UniConvNumDBToCompanyWithOutChange(lgObjRs(13),0) & _
							Chr(11) & ConvSPChars(lgObjRs(14)) & _
							Chr(11) & ConvSPChars(lgObjRs(15)) & _
							Chr(11) & ConvSPChars(lgObjRs(16))
			If ConvSPChars(lgObjRs(6)) = "Y" Then
			lgstrData =	lgstrData &	Chr(11) & ConvSPChars(lgObjRs(17))
			Else
			lgstrData =	lgstrData &	Chr(11) & ""
			End If	
			
			lgstrData =	lgstrData &	Chr(11) & ConvSPChars(lgObjRs(18)) & _
							Chr(11) & ConvSPChars(lgObjRs(19)) & _
							Chr(11) & ConvSPChars(lgObjRs(20)) & _
							Chr(11) & ConvSPChars(lgObjRs(21)) & _
							Chr(11) & ConvSPChars(lgObjRs(22)) & _
							Chr(11) & ConvSPChars(lgObjRs(23)) & _
							Chr(11) & ConvSPChars(lgObjRs(24)) & _
							Chr(11) & ConvSPChars(lgObjRs(25)) & _
							Chr(11) & ConvSPChars(lgObjRs(26)) & _
							Chr(11) & ConvSPChars(lgObjRs(27)) & _
							Chr(11) & ConvSPChars(lgObjRs(28)) & _
							Chr(11) & ConvSPChars(lgObjRs(29)) & _
							Chr(11) & ConvSPChars(lgObjRs(30)) & _
							Chr(11) & ConvSPChars(lgObjRs(31)) & _
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
    
	lgStrSQL =	" SELECT A.PLANT_CD, B.PLANT_NM, A.ITEM_CD, C.ITEM_NM, C.SPEC, A.TRACKING_NO, A.INSPECT_FLG, " & _
				" A.MVMT_RCPT_QTY, A.MVMT_QTY, A.MVMT_RCPT_UNIT, A.MVMT_CUR, A.MVMT_PRC, A.MVMT_DOC_AMT, A.MVMT_LOC_AMT, " & _
				" A.MVMT_SL_CD,  D.SL_NM, I.MINOR_NM, J.MINOR_NM, A.LOT_NO, A.LOT_SUB_NO,  A.MAKER_LOT_NO, A.MAKER_LOT_SUB_NO, " & _
				" A.GM_NO, A.GM_SEQ_NO, A.INSPECT_REQ_NO, A.INSPECT_RESULT_NO, A.PO_NO, A.PO_SEQ_NO, A.CC_NO, A.CC_SEQ, " & _
				" A.LC_NO, A.LC_SEQ, A.IO_TYPE_CD, E.IO_TYPE_NM, A.MVMT_RCPT_DT, A.BP_CD, F.BP_NM, A.PUR_GRP, G.PUR_GRP_NM  " & _
				" FROM M_PUR_GOODS_MVMT A LEFT OUTER JOIN B_MINOR I ON (A.INSPECT_STS =  I.MINOR_CD AND I.MAJOR_CD =  " & FilterVar("M4103", "''", "S") & ") " & _
				"      LEFT OUTER JOIN B_MINOR J ON (A.MVMT_METHOD =  J.MINOR_CD AND I.MAJOR_CD =  " & FilterVar("B9016", "''", "S") & ") " & _
				"      LEFT OUTER JOIN B_STORAGE_LOCATION D ON (A.MVMT_SL_CD =  D.SL_CD) " & _
				"      , B_PLANT B, B_ITEM C, M_MVMT_TYPE E, B_BIZ_PARTNER F, B_PUR_GRP G " & _
				" WHERE A.PLANT_CD = B.PLANT_CD AND A.ITEM_CD = C.ITEM_CD "	& _
				" AND A.BP_CD = F.BP_CD AND A.PUR_GRP = G.PUR_GRP AND A.IO_TYPE_CD = E.IO_TYPE_CD "	& _
				" AND A.MVMT_RCPT_NO =  " & FilterVar(strRcptNo , "''", "S") & " " & _
				" ORDER BY A.MVMT_NO "
			

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
				.frm1.txtSupplierNm.value = "<%=ConvSPChars(strBpNm)%>"
				.frm1.txtGroupCd.value = "<%=ConvSPChars(strPurGrp)%>"
				.frm1.txtGroupNm.value = "<%=ConvSPChars(strPurGrpNm)%>"
				.frm1.txtMvmtType.value = "<%=ConvSPChars(strMvmtType)%>"
				.frm1.txtMvmtTypeNm.value = "<%=ConvSPChars(strMvmtTypeNm)%>"
				.frm1.txtGmDt.text = "<%=UNIDateClientFormat(strMvmtDt)%>"
			End if
			
			.frm1.vspdData.ReDraw = False
			.ggoSpread.SSShowData "<%=lgstrData%>", "F"
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>","<%=iEndRow%>", .C_Cur,.C_MvmtPrc,"C","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>","<%=iEndRow%>", .C_Cur,.C_DocAmt,"A","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>","<%=iEndRow%>",gCurrency,.C_LocAmt,"A","Q","X","X")
			
			.DbQueryOk
			.frm1.vspdData.focus
			.frm1.vspdData.ReDraw = True

		End If   
	End With	
</Script>	


	
