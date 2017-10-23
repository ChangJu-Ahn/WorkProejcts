<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        :
'*  3. Program ID           : XI316QB1_ko119
'*  4. Program Name         : Q-FOCUS 전송현황
'*  5. Program Desc         :
'*  6. Component List       :
'*  7. Modified date(First) : 2007/05/10
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Kim Jin Tae
'* 10. Modifier (Last)      :
'* 11. Comment				:
'* 12. Common Coding Guide  : this mark(☜) Means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd															'☜: 모든 작업 완료후 작업진행중 표시창을 Hide
Err.Clear

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "Q", "NOCOOKIE", "QB")

Dim IntRetCD
Dim PvArr
Dim NextKey1
Dim strNextKey1

Const C_SHEETMAXROWS_D = 500

lgLngMaxRow     = Request("txtMaxRows")
lgErrorStatus   = "NO"

Call HideStatusWnd

On Error Resume Next

Call SubOpenDB(lgObjConn)
Call SubCreateCommandObject(lgObjComm)

Call SubBizQuery()

Call SubCloseCommandObject(lgObjComm)
Call SubCloseDB(lgObjConn)

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	Dim iDx

	On Error Resume Next
    Err.Clear

	Call SubMakeSQLStatements

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
		IntRetCD = -1
		lgStrPrevKeyIndex = ""
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		Call SetErrorStatus()
		Call SubCloseRs(lgObjRs)
		Response.End
	Else
		IntRetCD = 1

        iDx = 0
        ReDim PvArr(C_SHEETMAXROWS_D)

        Do While Not lgObjRs.EOF
 
            If iDx = C_SHEETMAXROWS_D Then
               NextKey1 = ConvSPChars(lgObjRs(8))
               Exit Do
            End If

			If trim(lgObjRs(0)) <> "" Then
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(0))
			Else
            lgstrData = Chr(11) & ConvSPChars(lgObjRs(27))
			End If

			lgstrData = lgstrData & _
						Chr(11) & ConvSPChars(lgObjRs(1)) & _
						Chr(11) & ConvSPChars(lgObjRs(2)) & _
						Chr(11) & ConvSPChars(lgObjRs(3)) & _
						Chr(11) & ConvSPChars(lgObjRs(4)) & _
						Chr(11) & ConvSPChars(lgObjRs(5)) & _
						Chr(11) & ConvSPChars(lgObjRs(6)) & _
						Chr(11) & UNIDateClientFormat(lgObjRs(7)) & _
						Chr(11) & ConvSPChars(lgObjRs(8)) & _
						Chr(11) & ConvSPChars(lgObjRs(9)) & _
						Chr(11) & ConvSPChars(lgObjRs(10)) & _
						Chr(11) & ConvSPChars(lgObjRs(11)) & _
						Chr(11) & ConvSPChars(lgObjRs(12)) & _
						Chr(11) & ConvSPChars(lgObjRs(13)) & _
						Chr(11) & ConvSPChars(lgObjRs(14)) & _
                        Chr(11) & ConvSPChars(lgObjRs(15)) & _
						Chr(11) & ConvSPChars(lgObjRs(16)) & _
						Chr(11) & ConvSPChars(lgObjRs(17)) & _
						Chr(11) & ConvSPChars(lgObjRs(18)) & _
						Chr(11) & ConvSPChars(lgObjRs(19)) & _
						Chr(11) & ConvSPChars(lgObjRs(20)) & _
						Chr(11) & ConvSPChars(lgObjRs(21)) & _
						Chr(11) & ConvSPChars(lgObjRs(22)) & _
						Chr(11) & ConvSPChars(lgObjRs(23)) & _
						Chr(11) & ConvSPChars(lgObjRs(24)) & _
						Chr(11) & UniConvNumberDBToCompany(lgObjRs(25), ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit, 0) & _
						Chr(11) & ConvSPChars(lgObjRs(26)) & _
						Chr(11) & lgLngMaxRow + iDx & Chr(11) & Chr(12)

			PvArr(iDx) = lgstrData
			iDx = iDx + 1
		    lgObjRs.MoveNext
        Loop
    End If

	lgstrData = Join(PvArr, "")

	Call SubHandleError("MR", lgObjConn,lgObjRs, Err)
    Call SubCloseRs(lgObjRs)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements()

    On Error Resume Next
    Err.Clear

	dim str_fr_dt, str_to_dt
	Dim str_item_cd

	str_fr_dt = FilterVar(UNIConvDate(Request("txtDtFr")), "", "S")
	str_to_dt = FilterVar(UNIConvDate(Request("txtDtTo")), "", "S")
	str_item_cd = trim(Request("txtItemCd"))

	lgStrSQL = " SELECT Top " & C_SHEETMAXROWS_D + 1 & _
					" b.SEC_INVOICE_NO, b.item_cd, c.item_nm, c.spec, b.pallet_no, " & _
					" b.tray_no, b.lot_no, b.production_dt, a.lot_no as mat_lot_no, " & _
					" a.PART1_NO, a.PART2_NO, a.PART3_NO, a.PART4_NO, a.PART5_NO, a.PART6_NO, " & _
					" a.PART7_NO, a.PART8_NO, a.PART9_NO, a.PART10_NO, a.PART11_NO, a.PART12_NO, " & _
					" a.erp_receive_dt, " & _
					" a.erp_apply_flag, case when a.erp_apply_flag = 'Y' then a.updt_dt Else '' end as send_dt, " & _
					" a.create_type, b.pallet_item_qty, b.prodt_order_no, d.prev_invoice_no " & _
				" FROM T_IF_RCV_LOT_ASSY_INFO_KO119 a " & _
					" left outer join T_IF_RCV_TRAY_INFO_KO119 b " & _
						" ON a.plant_cd = b.plant_cd and a.pallet_no = b.pallet_no and a.tray_no = b.tray_no " & _
					" left outer join b_item c " & _
						" ON b.item_cd = c.item_cd " & _
					" left outer join S_PREV_INTEGRATE_LBL_KO119 d " & _
					"	ON a.pallet_no = d.pallet_no " & _
				" WHERE a.ERP_RECEIVE_DT >= left(" & str_fr_dt & ", 4) + substring(" & str_fr_dt & ", 6,2) + right(" & str_fr_dt & ", 2) + '000000' " & _
					" AND a.ERP_RECEIVE_DT <=left(" & str_to_dt & ", 4) + substring(" & str_to_dt & ", 6,2) + right(" & str_to_dt & ", 2) + '999999' "

	If Request("lgStrPrevKey") <> "" Then
		lgStrSQL = lgStrSQL & " and a.lot_no >= " & FilterVar(Request("lgStrPrevKey"), "", "S")
	End If

	If Trim(Request("txtrdoFlag")) = "Y" Then
		lgStrSQL = lgStrSQL & " AND a.erp_apply_flag = 'Y' "
	ElseIf Trim(Request("txtrdoFlag")) = "N" Then
		lgStrSQL = lgStrSQL & " AND a.erp_apply_flag = 'N' "
	Else
	End If

	If Trim(Request("txtItemCd")) <> "" Then
		lgStrSQL = lgStrSQL & " AND b.item_cd = " & FilterVar(str_item_cd, "", "S")
	End If 
	
'	lgStrSQL = lgStrSQL & " ORDER BY a.updt_dt desc, b.SEC_INVOICE_NO, a.lot_no "
	lgStrSQL = lgStrSQL & " ORDER BY a.lot_no, b.SEC_INVOICE_NO "
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

Response.Write "<Script language=vbs> " & vbCr
Response.Write " With Parent "      	& vbCr
Response.Write "	If """ & lgErrorStatus & """ = ""NO"" And """ & IntRetCd & """ <> -1 Then "	& vbCr
Response.Write "    .lgStrPrevKey  = """ & NextKey1 & """"		& vbCr
Response.Write "	.ggoSpread.Source	= .frm1.vspdData "		& vbCr
Response.Write "	.ggoSpread.SSShowDataByClip  """ & lgstrData  & """"        & vbCr
Response.Write "		If .frm1.vspdData.MaxRows < .parent.VisibleRowCnt(.frm1.vspdData, 0) And .lgStrPrevKey <> """" Then "	& vbCr
Response.Write "			.DbQuery			"	& vbCr
Response.Write "		Else					"	& vbCr
Response.Write "			.DbQueryOK			"	& vbCr
Response.Write "		End If					"	& vbCr
Response.Write "		.frm1.vspdData.focus	"	& vbCr
Response.Write "    End If						"	& vbCr
Response.Write " End With "             & vbCr
Response.Write "</Script> "             & vbCr
Response.End
%>