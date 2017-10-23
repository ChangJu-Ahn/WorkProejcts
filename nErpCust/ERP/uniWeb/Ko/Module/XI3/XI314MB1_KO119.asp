<%@ LANGUAGE=VBSCript%>
<% Option Explicit%>
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->

<%

    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()
    Call LoadInfTB19029B( "I", "*", "NOCOOKIE", "MB")
    Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "MB")

    Dim strSQL
	Dim istrMode
	Dim strRetMsg
	Dim IntRetCd
		
    Dim txtPlantCd
	Dim txtUserId

	On Error Resume Next
	Err.Clear                                                                        '☜: Clear Error status

	txtPlantCd	= UCase(Trim(Request("txtPlant")))
	txtUserId	= UCase(Trim(Request("txtUserId")))

	istrMode = Request("txtMode")

	Select Case istrMode
			Case CStr("T")								'MES의 검사요청 자료를 ERP로 복사
				Call SubBizMesRcv()

			Case CStr("U")								' 출하중지여부 수정
				Call SubBizSaveMulti()

			Case CStr("R")
				Call SubBizQueryMulti()

	End Select

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	Const C_SHEETMAXROWS_D  = 100
	
	Const C_IG_plant_cd    = 0
	Const C_IG_pallet_no   = 1
	Const C_IG_tray_no     = 2
	Const C_IG_item_cd     = 3
	Const C_IG_sub_lot_no  = 4
	Const C_IG_if_seq      = 5
	Const C_IG_create_type = 6
		
	Dim Indx
	Dim lgStrPrevKey
	Dim lgArrPrevKey

	On Error Resume Next
	Err.Clear 
	
	Call SubOpenDB(lgObjConn)				' 데이터 베이스 커넥션 개체 생성
	
	lgArrPrevKey = Split(Request("lgStrPrevKey"), Chr(11))
	

	strSQL = " SELECT TOP " & C_SHEETMAXROWS_D + 1
	strSQL = strSQL & vbCrLf & "    A.PALLET_NO"
	strSQL = strSQL & vbCrLf & "  , A.TRAY_NO"
	strSQL = strSQL & vbCrLf & "  , A.SEC_ITEM_CD"
	strSQL = strSQL & vbCrLf & "  , A.SEC_ITEM_NM"
	strSQL = strSQL & vbCrLf & "  , A.ITEM_CD"
	strSQL = strSQL & vbCrLf & "  , A.ITEM_NM"
	strSQL = strSQL & vbCrLf & "  , A.SUB_LOT_NO"
	strSQL = strSQL & vbCrLf & "  , A.IF_SEQ"
	strSQL = strSQL & vbCrLf & "  , A.LOT_NO"
	strSQL = strSQL & vbCrLf & "  , A.PALLET_ITEM_QTY"
	strSQL = strSQL & vbCrLf & "  , A.TRAY_ITEM_QTY"
	strSQL = strSQL & vbCrLf & "  , A.PRODUCTION_DT"
	strSQL = strSQL & vbCrLf & "  , A.PRODT_ORDER_NO"
	strSQL = strSQL & vbCrLf & "  , A.CREATE_TYPE"
	strSQL = strSQL & vbCrLf & "  , A.STATUS"
	strSQL = strSQL & vbCrLf & "  , A.SEND_DT"
	strSQL = strSQL & vbCrLf & "  , A.ERP_RECEIVE_DT"
	strSQL = strSQL & vbCrLf & "  , A.DELIVERY_HOLD_FG"
	strSQL = strSQL & vbCrLf & "  , A.sec_invoice_no "
	strSQL = strSQL & vbCrLf & "  , A.ERR_DESC"
	strSQL = strSQL & vbCrLf & "  , A.PLANT_CD"
	strSQL = strSQL & vbCrLf & "  , A.INTEGRATE_LBL_NO "	
	strSQL = strSQL & vbCrLf & "  FROM ( SELECT "



	strSQL = strSQL & vbCrLf & "    A.PALLET_NO"
	strSQL = strSQL & vbCrLf & "  , A.TRAY_NO"
	strSQL = strSQL & vbCrLf & "  , D.SEC_ITEM_CD"
	strSQL = strSQL & vbCrLf & "  , D.SEC_ITEM_NM"
	strSQL = strSQL & vbCrLf & "  , A.ITEM_CD"
	strSQL = strSQL & vbCrLf & "  , B.ITEM_NM"
	strSQL = strSQL & vbCrLf & "  , A.SUB_LOT_NO"
	strSQL = strSQL & vbCrLf & "  , A.IF_SEQ"
	strSQL = strSQL & vbCrLf & "  , A.LOT_NO"
	strSQL = strSQL & vbCrLf & "  , A.PALLET_ITEM_QTY"
	strSQL = strSQL & vbCrLf & "  , A.TRAY_ITEM_QTY"
	strSQL = strSQL & vbCrLf & "  , A.PRODUCTION_DT"
	strSQL = strSQL & vbCrLf & "  , A.PRODT_ORDER_NO"
	strSQL = strSQL & vbCrLf & "  , A.CREATE_TYPE"
	strSQL = strSQL & vbCrLf & "  , A.STATUS"
	strSQL = strSQL & vbCrLf & "  , A.SEND_DT"
	strSQL = strSQL & vbCrLf & "  , A.ERP_RECEIVE_DT"
	strSQL = strSQL & vbCrLf & "  , A.DELIVERY_HOLD_FG"
	strSQL = strSQL & vbCrLf & "  , case when isnull(A.SEC_INVOICE_NO, '') = '' then z.prev_invoice_no else a.sec_invoice_no end  as sec_invoice_no "
	strSQL = strSQL & vbCrLf & "  , A.ERR_DESC"
	strSQL = strSQL & vbCrLf & "  , A.PLANT_CD"
	strSQL = strSQL & vbCrLf & "  , case when isnull(y.INTEGRATE_LBL_NO, '') = '' then z.INTEGRATE_LBL_NO else y.INTEGRATE_LBL_NO end  as INTEGRATE_LBL_NO "	

	strSQL = strSQL & vbCrLf & " From T_IF_RCV_TRAY_INFO_KO119  a "
	strSQL = strSQL & vbCrLf & " 	inner join b_item b on a.plant_cd = '" & Request("txtPlant") & "' and a.item_cd = b.item_cd "
	strSQL = strSQL & vbCrLf & " 	inner join (select distinct item_cd, sec_item_cd, sec_item_nm from b_item_mapping_ko119)  d on a.item_cd = d.item_cd "
	strSQL = strSQL & vbCrLf & " 	left outer join S_PREV_INTEGRATE_LBL_KO119 z on a.pallet_no = z.pallet_no  "
	strSQL = strSQL & vbCrLf & " 	left outer join s_integrate_lbl_hdr_ko119 y on a.pallet_no = y.pallet_no "
	

	
	strSQL = strSQL & vbCrLf & " WHERE"
	strSQL = strSQL & vbCrLf & "     CONVERT(VARCHAR(10), A.SEND_DT, 121) >= '" & UNIConvDate(Request("txtConSoFrDt")) & "' and CONVERT(VARCHAR(10), A.SEND_DT, 121) <= '" & UNIConvDate(Request("txtConSoToDt")) & "' "

	IF TRIM(Request("txtSecItemCd")) <> "" THEN
		strSQL = strSQL & vbCrLf & "  AND D.SEC_ITEM_CD =  '" & TRIM(Request("txtSecItemCd")) & "'"
	END IF
	
	IF TRIM(Request("txtMakOrdNo")) <> "" THEN
		strSQL = strSQL & vbCrLf & "  AND A.PRODT_ORDER_NO = '" & TRIM(Request("txtMakOrdNo")) & "'"
	END IF

	IF TRIM(Request("txtPalletNo")) <> "" THEN
		strSQL = strSQL & vbCrLf & "  AND A.PALLET_NO = '" & TRIM(Request("txtPalletNo")) & "'"
	END IF

	IF TRIM(Request("txtTrayNo")) <> "" THEN
		strSQL = strSQL & vbCrLf & "  AND A.TRAY_NO = '" & TRIM(Request("txtTrayNo")) & "'"
	END IF

	If Ucase(Trim(Request("rdoFlag"))) = "A" Then      '전체
		strSQL = strSQL & vbCrLf & "  AND ISNULL(A.SEC_INVOICE_NO, '') = ISNULL(A.SEC_INVOICE_NO, '')"
	ElseIf Ucase(Trim(Request("rdoFlag"))) = "C" Then  '완료
		strSQL = strSQL & vbCrLf & "  AND ISNULL(A.SEC_INVOICE_NO, '') <> ''"
	Else                                    '미완료
		strSQL = strSQL & vbCrLf & "  AND ISNULL(A.SEC_INVOICE_NO, '') = ''"
	End If

	strSQL = strSQL & vbCrLf & "  ) A"
	strSQL = strSQL & vbCrLf & "  WHERE 1= 1 "

	If Trim(Request("txtInvoiceNo")) <> "" Then
		strSQL = strSQL & vbCrLf & "  and A.SEC_INVOICE_NO = '" & Trim(Request("txtInvoiceNo")) & "'"
	End If


		
	If IsArray(lgArrPrevKey) Then
		If Ubound(lgArrPrevKey, 1) >= C_IG_if_seq Then
			strSQL = strSQL & vbCrLf & "   AND (A.PLANT_CD > '" & Trim(lgArrPrevKey(C_IG_plant_cd)) & "'"
			strSQL = strSQL & vbCrLf & "   OR   (A.PLANT_CD = '" & Trim(lgArrPrevKey(C_IG_plant_cd)) & "'"
			strSQL = strSQL & vbCrLf & "   AND   (A.PALLET_NO > '" & Trim(lgArrPrevKey(C_IG_pallet_no)) & "'"
			strSQL = strSQL & vbCrLf & "   OR     (A.PALLET_NO = '" & Trim(lgArrPrevKey(C_IG_pallet_no)) & "'"
			strSQL = strSQL & vbCrLf & "   AND     (A.TRAY_NO > '" & Trim(lgArrPrevKey(C_IG_tray_no)) & "'"
			strSQL = strSQL & vbCrLf & "   OR       (A.TRAY_NO = '" & Trim(lgArrPrevKey(C_IG_tray_no)) & "'"
			strSQL = strSQL & vbCrLf & "   AND       (A.ITEM_CD > '" & Trim(lgArrPrevKey(C_IG_item_cd)) & "'"
			strSQL = strSQL & vbCrLf & "   OR         (A.ITEM_CD = '" & Trim(lgArrPrevKey(C_IG_item_cd)) & "'"
			strSQL = strSQL & vbCrLf & "   AND         (A.SUB_LOT_NO > '" & Trim(lgArrPrevKey(C_IG_sub_lot_no)) & "'"
			strSQL = strSQL & vbCrLf & "   OR           (A.SUB_LOT_NO = '" & Trim(lgArrPrevKey(C_IG_sub_lot_no)) & "'"
			strSQL = strSQL & vbCrLf & "   AND           (A.IF_SEQ > " & Trim(lgArrPrevKey(C_IG_if_seq))
			strSQL = strSQL & vbCrLf & "   OR             (A.IF_SEQ = " & Trim(lgArrPrevKey(C_IG_if_seq))
			strSQL = strSQL & vbCrLf & "   AND              A.CREATE_TYPE >= '" & Trim(lgArrPrevKey(C_IG_create_type)) & "'))))))))))))"
		End If
	End If

	strSQL = strSQL & vbCrLf & " ORDER BY"
	strSQL = strSQL & vbCrLf & "    A.PLANT_CD"
	strSQL = strSQL & vbCrLf & "  , A.PALLET_NO"
	strSQL = strSQL & vbCrLf & "  , A.TRAY_NO"
	strSQL = strSQL & vbCrLf & "  , A.ITEM_CD"
	strSQL = strSQL & vbCrLf & "  , A.SUB_LOT_NO"
	strSQL = strSQL & vbCrLf & "  , A.IF_SEQ"
	strSQL = strSQL & vbCrLf & "  , A.CREATE_TYPE"


'Response.Write strSQl
'Response.End

	Call FncOpenRs("R", lgObjConn, lgObjRs, strSQL, "X", "X")

	If lgObjRs.EOF Then
		Response.Write "<Script Language=""VBScript"">"           & vbCrLf
		Response.Write " MsgBox ""해당 조회 조건의 Tray Info 수신자료가 존재하지 않습니다."""
		Response.Write "</Script>" & vbCrLf

		Call SubCloseRs(lgObjRs)
		Call SubCloseDB(lgObjConn)
		Response.End
	End If
		
	Response.Write "<Script Language=vbScript>                 " & vbCr
	Response.Write "	With Parent                            " & vbCr
	Response.Write "		.ggoSpread.Source = .frm1.vspdData " & vbCr
	Response.Write "		.ggoSpread.SSShowDataByClip """

	
	Indx = 0
	lgStrPrevKey = ""
	Do While Not lgObjRs.EOF
	
		If Indx < C_SHEETMAXROWS_D Then

			If Ucase(lgObjRs("DELIVERY_HOLD_FG")) = "Y" Then
				Response.Write Chr(11) & "1"
			 Else
				Response.Write Chr(11) & "0"
			End If

			Response.Write Chr(11) & ConvSPChars(lgObjRs("PALLET_NO"))
			Response.Write Chr(11) & ConvSPChars(lgObjRs("TRAY_NO"))
			Response.Write Chr(11) & lgObjRs("SEC_ITEM_CD")
			Response.Write Chr(11) & Replace(lgObjRs("SEC_ITEM_NM"), """", """""")
			Response.Write Chr(11) & lgObjRs("ITEM_CD")
			Response.Write Chr(11) & Replace(lgObjRs("ITEM_NM"), """", """""")
			Response.Write Chr(11) & lgObjRs("SUB_LOT_NO")
			Response.Write Chr(11) & lgObjRs("IF_SEQ")
			Response.Write Chr(11) & lgObjRs("LOT_NO")
			Response.Write Chr(11) & UniNumClientFormat(lgObjRs("PALLET_ITEM_QTY"), ggQty.DecPoint, 0)
			Response.Write Chr(11) & UniNumClientFormat(lgObjRs("TRAY_ITEM_QTY"), ggQty.DecPoint, 0)
			Response.Write Chr(11) & lgObjRs("PRODUCTION_DT")
			Response.Write Chr(11) & lgObjRs("PRODT_ORDER_NO")
			Response.Write Chr(11) & ConvSPChars(lgObjRs("INTEGRATE_LBL_NO"))			
			Response.Write Chr(11) & lgObjRs("CREATE_TYPE")
			Response.Write Chr(11) & lgObjRs("STATUS")
			Response.Write Chr(11) & lgObjRs("SEND_DT")
			Response.Write Chr(11) & lgObjRs("ERP_RECEIVE_DT")
			Response.Write Chr(11) & ConvSPChars(lgObjRs("ERR_DESC"))
			Response.Write Chr(11) & ConvSPChars(lgObjRs("SEC_INVOICE_NO"))
			Response.Write Chr(11) & Chr(12)
		Else
			lgStrPrevKey = lgStrPrevKey & lgObjRs("PLANT_CD")
			lgStrPrevKey = lgStrPrevKey & Chr(11) & lgObjRs("PALLET_NO")
			lgStrPrevKey = lgStrPrevKey & Chr(11) & lgObjRs("TRAY_NO")
			lgStrPrevKey = lgStrPrevKey & Chr(11) & lgObjRs("ITEM_CD")
			lgStrPrevKey = lgStrPrevKey & Chr(11) & lgObjRs("SUB_LOT_NO")
			lgStrPrevKey = lgStrPrevKey & Chr(11) & lgObjRs("IF_SEQ")
			lgStrPrevKey = lgStrPrevKey & Chr(11) & lgObjRs("CREATE_TYPE")
		End If
						
		Indx = Indx + 1
		lgObjRs.MoveNext
	Loop
	
	Response.Write                                                """" & vbCr
	Response.Write "		.lgStrPrevKey = """ & lgStrPrevKey  & """" & vbCr
	Response.Write "		.DBQueryOk "						& vbCr
	Response.Write "	End with "								& vbCr
	Response.Write "</Script>"									& vbCr

	Call SubCloseRs(lgObjRs)
	Call SubCloseDB(lgObjConn)

End Sub
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	On Error Resume Next                                                            '☜: Protect system from crashing
	Err.Clear 

    Const C_IG_row_flag         = 0
    Const C_IG_plant_cd         = 1
	Const C_IG_pallet_no        = 2
	Const C_IG_tray_no          = 3
	Const C_IG_item_cd          = 4
	Const C_IG_sub_lot_no       = 5
	Const C_IG_if_seq           = 6
	Const C_IG_create_type      = 7
	Const C_IG_delivery_hold_fg = 8
	Const C_IG_pallet_qty       = 9
	Const C_IG_row_num          = 10

	Dim ObjPxi2g19Ko119
    Dim iErrorPosition
    Dim iUpdtUserId
    Dim itxtSpread
    '-------------------
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount
    Dim Indx

    Dim iCUCount
    Dim iDCount
    
    Dim arrColVal, arrRowVal
    
    itxtSpread = ""
             
    iCUCount = Request.Form("txtCUSpread").Count
    iDCount  = Request.Form("txtDSpread").Count

    itxtSpreadArrCount = -1
             
    ReDim itxtSpreadArr(iCUCount + iDCount)
    
    For Indx = 1 To iDCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtDSpread")(Indx)
    Next
    
    For Indx = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(Indx)
    Next
    
    itxtSpread = Join(itxtSpreadArr,"")

	If itxtSpread <> "" Then
	
		arrRowVal = Split(itxtSpread, gRowSep)
		
        ReDim spd_data_set(UBound(arrRowVal) - 1, C_IG_row_num)
        
        lgStrSQL = " Begin Tran "

        For Indx = 0 To UBound(arrRowVal) - 1
            arrColVal = Split(arrRowVal(Indx), gColSep)
            
            Select Case Ucase(Trim(arrColVal(C_IG_row_flag)))
				Case "U"
					lgStrSQL = lgStrSQL & " UPDATE  T_IF_RCV_TRAY_INFO_KO119 SET "
					lgStrSQL = lgStrSQL & "	   DELIVERY_HOLD_FG = " & FilterVar(arrColVal(C_IG_delivery_hold_fg), "''", "S")
					lgStrSQL = lgStrSQL & "  , PALLET_ITEM_QTY = " & arrColVal(C_IG_pallet_qty)
					lgStrSQL = lgStrSQL & " WHERE"
					lgStrSQL = lgStrSQL & "      PLANT_CD    = " & FilterVar(arrColVal(C_IG_plant_cd), "''", "S")
					lgStrSQL = lgStrSQL & "  AND PALLET_NO   = " & FilterVar(arrColVal(C_IG_pallet_no), "''", "S") & vbCr
					lgStrSQL = lgStrSQL & "  AND TRAY_NO     = " & FilterVar(arrColVal(C_IG_tray_no), "''", "S")
					lgStrSQL = lgStrSQL & "  AND ITEM_CD     = " & FilterVar(arrColVal(C_IG_item_cd), "''", "S")
					lgStrSQL = lgStrSQL & "  AND SUB_LOT_NO  = " & FilterVar(arrColVal(C_IG_sub_lot_no), "''", "S")
					lgStrSQL = lgStrSQL & "  AND IF_SEQ      = " & FilterVar(arrColVal(C_IG_if_seq), "''", "SNM")
					lgStrSQL = lgStrSQL & "  AND CREATE_TYPE = " & FilterVar(arrColVal(C_IG_create_type), "''", "S") & vbCr
				
				Case "D"
					lgStrSQL = lgStrSQL & " DELETE FROM T_IF_RCV_TRAY_INFO_KO119 "
					lgStrSQL = lgStrSQL & " WHERE"
					lgStrSQL = lgStrSQL & "      PLANT_CD    = " & FilterVar(arrColVal(C_IG_plant_cd), "''", "S")
					lgStrSQL = lgStrSQL & "  AND PALLET_NO   = " & FilterVar(arrColVal(C_IG_pallet_no), "''", "S")
					lgStrSQL = lgStrSQL & "  AND TRAY_NO     = " & FilterVar(arrColVal(C_IG_tray_no), "''", "S")
					lgStrSQL = lgStrSQL & "  AND ITEM_CD     = " & FilterVar(arrColVal(C_IG_item_cd), "''", "S")
					lgStrSQL = lgStrSQL & "  AND SUB_LOT_NO  = " & FilterVar(arrColVal(C_IG_sub_lot_no), "''", "S")
					lgStrSQL = lgStrSQL & "  AND IF_SEQ      = " & FilterVar(arrColVal(C_IG_if_seq), "''", "SNM")
					lgStrSQL = lgStrSQL & "  AND CREATE_TYPE = " & FilterVar(arrColVal(C_IG_create_type), "''", "S") & vbCr
			End Select
        Next
        
    End If
    '---------------------         

    Call SubOpenDB(lgObjConn)				' 데이터 베이스 커넥션 개체 생성
    
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
	
    If CheckSYSTEMError(Err,True) = True Then
		lgStrSQL = " Rollback Tran "
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    Else
		lgStrSQL = " Commit Tran "		
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords    

		Response.Write "<Script language=vbScript>                                    " & vbCr  
		Response.Write " With parent                                                  " & vbCr       
		Response.Write "    Call .DBSaveOk()                                          " & vbCr   
		Response.Write " End With                                                     " & vbCr
		Response.Write "</Script> "
	End If
    
    Call SubCloseDB(lgObjConn)

End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
'    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus = "YES"
       ObjectContext.SetAbort
    End If

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
'    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus = "YES"
       ObjectContext.SetAbort
    End If

End Sub

'============================================================================================================
' Name : SubBizSaveMultiDelete
' Desc : Delete Data from Db
'============================================================================================================
Sub SubBizSaveMultiDelete(arrColVal)

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    
    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus = "YES"
       ObjectContext.SetAbort
    End If
	
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizMesRcv()

	Call SubOpenDB(lgObjConn)						' 데이터 베이스 커넥션 개체 생성
	Call SubCreateCommandObject(lgObjComm)

	With lgObjComm
		.CommandText = "usp_TRAY_INFO_MES_RCV_KO119"
		.CommandType = adCmdStoredProc
		.CommandTimeout = 1800	

		lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	adInteger,adParamReturnValue)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CON_PLANT_CD",	adVarChar, adParamInput,    4, txtPlantCd)
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CON_USER_ID",	adVarChar, adParamInput,   13, txtUserId)	   
		lgObjComm.Parameters.Append lgObjComm.CreateParameter("@CON_ERROR_DESC",adVarChar, adParamOutput, 200)

		lgObjComm.Execute ,, adExecuteNoRecords
	End With

	If Err.number = 0 Then
		intRetCd = lgObjComm.Parameters("RETURN_VALUE").Value
						
		If intRetCd <> 0 Then
			strRetMsg = lgObjComm.Parameters("@CON_ERROR_DESC").Value
			If strRetMsg <> "" Then
				Call DisplayMsgBox(strRetMsg, vbInformation, "", "", I_MKSCRIPT)
			End If	
		End If
	Else
		Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)	
	End If
		
	Response.Write "<Script Language=vbscript>	"	& vbcr
	Response.Write "With parent.frm1			"	& vbcr
	Response.Write "	parent.FncQuery			"	& vbcr
	Response.Write "End With					"	& vbcr
	Response.Write "</Script>					"	& vbcr	 

	Call SubCloseCommandObject(lgObjComm)
	Call SubCloseDB(lgObjConn) 

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
'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub
%>