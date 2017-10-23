<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/AdoVbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServerAdoDB.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComAsp/LoadinfTB19029.asp" -->
<%
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

	Call LoadBasisGlobalInf
	Call LoadinfTB19029B("I", "*", "NOCOOKIE", "MB")

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	Dim IntRetCD
	Dim strPlantCd
	Dim strItemCd
	Dim strBomNo
	Dim strBaseDt
	Dim strExpFlg
	Dim strItemAcct
	Dim strItemAcctGrp
	Dim strPhantomFlg
	Dim strProcurType
	Dim strPrntItemCd
	Dim strPrntBomNo
	Dim strChildItemSeq
	Dim strChildBomNo
	Dim strChildItemCd
	Dim strBOMHeader
	Dim QueryType
	Dim strSpId
	Dim DtlFlg
	Dim iPos
	Dim BaseDt
	
	DtlFlg = 0
	strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")
	strItemCd = FilterVar(Trim(Request("txtItemCd"))	,"''", "S")
	strBomNo = FilterVar(Trim(Request("txtBomNo"))	,"''", "S")
	BaseDt = FilterVar(UNIConvYYYYMMDDToDate(gAPDateFormat,"1900","01","01"),"''","S")
	strCurDate = FilterVar(Request("CurDate"), BaseDt, "S")
	 
	QueryType = Request("QueryType")
	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
    Select Case QueryType
        Case "A"				'☜: 전체Query
			strBaseDt = FilterVar(Trim(Request("txtBaseDt")), "'1900-01-01'", "D")
			strExpFlg = FilterVar(Trim(Request("rdoSrchType")), "''", "S")
			
			Call SubBizQuery("CK")
			
			Call SubBizQuery("B_CK")
														
			If strBomNo <> "''" Then
				Call SubCreateCommandObject(lgObjComm)
				Call SubBizQuery("H")
				Call SubBizBatch()
				Call SubBizQueryMulti()
				Call SubCloseCommandObject(lgObjComm)
			Else
				Call DisplayMsgBox("182600", vbOKOnly, "", "", I_MKSCRIPT)
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.frm1.hBomType.value = """"" & vbCrLf
					Response.Write "Call parent.DbQueryNotOk()" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End 
			End If
			
        Case "H"								        							'☜: Header Query
			Call SubBizQuery("B_CK")
			Call SubBizQuery("H")
			
		Case "D"													'☜: Detail Query
			strItemCd = FilterVar(Trim(Request("txtChildItemCd"))	,"''", "S")
			strBomNo = FilterVar(Trim(Request("txtChildBomNo"))	,"''", "S")                                                     '☜: Detail Query
			strPrntItemCd =  FilterVar(Trim(Request("txtPrntItemCd"))	,"''", "S")
			strPrntBomNo =  FilterVar(Trim(Request("txtPrntBomNo"))	,"''", "S")
			strChildItemSeq = Request("intChildItemSeq")
			strChildBomNo =  FilterVar(Trim(Request("txtChildBomNo"))	,"''", "S")
			strChildItemCd =  FilterVar(Trim(Request("txtChildItemCd"))	,"''", "S")
			strBOMHeader =  Request("txtBOMHeader")
			DtlFlg = 1

			If strBOMHeader = "0" Then				'☜: Header 정보가 없는 Detail Query
				Call SubBizQuery("CK")
				Call SubBizQuery("P_CK")
				Call SubBizQuery("D")			
			Else								'☜: Header 정보가 있는 Detail Query
				Call SubBizQuery("B_CK")
				Call SubBizQuery("P_CK")
				Call SubBizQuery("H_D")
			End If

		Case "I"	
			strPrntItemCd =  FilterVar(Trim(Request("txtPrntItemCd"))	,"''", "S")
			iPos = Request("CurPos")
																	'☜: Lookup Item by plant
			Call SubBizQuery("I")
			Call SubBizQuery("P_CK")

		Case "B"													'☜: Lookup Bom Header 
			Call SubBizQuery("B")
			
    End Select
     
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pOpCode)

	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	Select Case pOpCode
		
		Case "B_CK"
			'--------------
			'공장 체크		
			'--------------	
			lgStrSQL = ""
			Call SubMakeSQLStatements("P_CK",strPlantCd,"","","","")           '☜ : Make sql statements
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "Parent.Frm1.txtPlantNm.Value  = """"" & vbCrLf
					Response.Write "Parent.Frm1.txtPlantCd.focus" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "Parent.Frm1.txtPlantNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			End If
		
			Call SubCloseRs(lgObjRs) 
			
			'------------------
			' bom type 체크 
			'------------------
			lgStrSQL = ""
			
			Call SubMakeSQLStatements("BT_CK", strBomNo, "", "", "", "")           '☜ : Make sql statements
			
		    If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				
				Call DisplayMsgBox("182622", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.frm1.hBomType.value = """"" & vbCrLf
					Response.Write "parent.frm1.txtBomNo.focus" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End							

		    Else
				IntRetCD = 1
			End If
		
			Call SubCloseRs(lgObjRs) 
			
			'------------------
			'품목, bom no 체크 
			'------------------
			lgStrSQL = ""
			Call SubMakeSQLStatements("B_CK", strPlantCd, strItemCd, strBomNo, "", "")           '☜ : Make sql statements
			
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				
				Call DisplayMsgBox("182600", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.frm1.hBomType.value = """"" & vbCrLf
					Response.Write "Call parent.DbQueryNotOk()" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End

		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						If DtlFlg <> 1 Then
							Response.Write ".frm1.txtItemNm.Value = """ & ConvSPChars(Trim(lgObjRs(14))) & """" & vbCrLf
						End If
						
						Response.Write ".Frm1.txtItemCd1.Value = """ & ConvSPChars(Trim(lgObjRs(1))) & """" & vbCrLf
						Response.Write ".Frm1.txtItemNm1.Value = """ & ConvSPChars(Trim(lgObjRs("ITEM_NM"))) & """" & vbCrLf
						Response.Write ".frm1.cboItemAcct.value = """ & ConvSPChars(Trim(lgObjRs(18))) & """" & vbCrLf
						Response.Write ".frm1.txtItemAcctGrp.value = """ & ConvSPChars(Trim(lgObjRs("ITEM_ACCT_GRP"))) & """" & vbCrLf
						Response.Write ".frm1.txtSpec.value = """ & ConvSPChars(lgObjRs(15)) & """" & vbCrLf
						Response.Write ".frm1.txtPlantItemFromDt.text = """ & UNIDateClientFormat(lgObjRs(12)) & """" & vbCrLf
						Response.Write ".frm1.txtPlantItemToDt.text = """ & UNIDateClientFormat(lgObjRs(13)) & """" & vbCrLf
						Response.Write ".frm1.txtProcType.value = """ & ConvSPChars(Trim(lgObjRs(17))) & """" & vbCrLf
						Response.Write ".frm1.txtBasicUnit.value = """ & ConvSPChars(Trim(lgObjRs(19))) & """" & vbCrLf
				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			End If
		
			Call SubCloseRs(lgObjRs) 
		
		Case "CK"
			'--------------
			'공장 체크		
			'--------------	
			lgStrSQL = ""
			Call SubMakeSQLStatements("P_CK", strPlantCd, "", "", "", "")           '☜ : Make sql statements
			
			If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "Parent.Frm1.txtPlantNm.Value  = """"" & vbCrLf
					Response.Write "Parent.Frm1.txtPlantCd.focus" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End
			
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "Parent.Frm1.txtPlantNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			End If
		
			Call SubCloseRs(lgObjRs) 
			
			'------------------
			'품목체크 
			'------------------
			lgStrSQL = ""
			Call SubMakeSQLStatements("I_CK",strPlantCd,strItemCd,"","","")           '☜ : Make sql statements

		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				
				Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.Frm1.txtItemNm.Value  = """"" & vbCrLf
					Response.Write "parent.Frm1.txtItemCd.focus" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".Frm1.txtItemCd1.Value = """ & ConvSPChars(Trim(lgObjRs("ITEM_CD"))) & """" & vbCrLf
						Response.Write ".Frm1.txtItemNm1.Value = """ & ConvSPChars(Trim(lgObjRs("ITEM_NM"))) & """" & vbCrLf
						Response.Write ".frm1.cboItemAcct.value = """ & ConvSPChars(Trim(lgObjRs("ITEM_ACCT"))) & """" & vbCrLf
						Response.Write ".frm1.txtItemAcctGrp.value = """ & ConvSPChars(Trim(lgObjRs("ITEM_ACCT_GRP"))) & """" & vbCrLf
						Response.Write ".frm1.txtSpec.value = """ & ConvSPChars(Trim(lgObjRs("SPEC"))) & """" & vbCrLf
						Response.Write ".frm1.txtPlantItemFromDt.text = """ & UNIDateClientFormat(lgObjRs("VALID_FROM_DT")) & """" & vbCrLf
						Response.Write ".frm1.txtPlantItemToDt.text = """ & UNIDateClientFormat(lgObjRs("VALID_TO_DT")) & """" & vbCrLf
						Response.Write ".frm1.txtProcType.value = """ & ConvSPChars(Trim(lgObjRs("PROCUR_TYPE"))) & """" & vbCrLf

				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			End If
		
			Call SubCloseRs(lgObjRs) 
			
		Case "P_CK"
			'------------------
			'상위 품목체크 
			'------------------
			lgStrSQL = ""
			Call SubMakeSQLStatements("PI_CK", strPlantCd, strPrntItemCd, "", "", "")           '☜ : Make sql statements

		    If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		    
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.frm1.hPrntProcType.value = """"" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End
     
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "parent.frm1.hPrntProcType.value	= """ & ConvSPChars(Trim(lgObjRs("PROCUR_TYPE"))) & """" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			
			End If
		
			Call SubCloseRs(lgObjRs) 
		
		Case "H"																	'☜: header 조회 경우 
			
			lgStrSQL = ""
			
		    Call SubMakeSQLStatements("H", strPlantCd, strItemCd, strBomNo, "", "")           '☜ : Make sql statements

		    If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
				
				IntRetCD = -1
				
				Call DisplayMsgBox("182600", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				Response.End 
		    Else
				IntRetCD = 1

				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".frm1.txtBomNo1.value = """ & ConvSPChars(Trim(lgObjRs(2))) & """" & vbCrLf
						Response.Write ".frm1.txtBOMDesc.value = """ & ConvSPChars(Trim(lgObjRs(3))) & """" & vbCrLf
						Response.Write ".frm1.txtDrawPath.value = """ & ConvSPChars(Trim(lgObjRs(7))) & """" & vbCrLf
						Response.Write ".frm1.hBomType.value = """ & ConvSPChars(Trim(lgObjRs(2))) & """" & vbCrLf
						Response.Write ".frm1.hPlantCd.value = """ & ConvSPChars(Trim(lgObjRs(0))) & """" & vbCrLf
				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
     
		    End If
		    
		    Call SubCloseRs(lgObjRs) 
		    
		Case "H_D"
		    
			'-----------------------
			' Level Setting
			'-----------------------
			strLevel = ""
			iLevelCnt = CInt(Request("intLevel"))
	
			For i = 1 To iLevelCnt
				strLevel = strLevel & "."
			Next 
	
			strLevel = strLevel & Request("intLevel")
			
			'-------------------------------
			' detail 정보 query
			'-------------------------------
			lgStrSQL = ""
			
		    Call SubMakeSQLStatements("H&D", strPlantCd, strPrntItemCd, strChildItemSeq, strPrntBomNo, "")                                       '☜ : Make sql statements
		
		    If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		    	Call DisplayMsgBox("182700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				IntRetCD = -1
				Response.End 
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".frm1.txtBomNo1.value = """ & ConvSPChars(Trim(lgObjRs("PRNT_BOM_NO"))) & """" & vbCrLf
						Response.Write ".frm1.txtBOMDesc.value = """ & ConvSPChars(Trim(lgObjRs("DESCRIPTION"))) & """" & vbCrLf
						Response.Write ".frm1.txtDrawPath.value = """ & ConvSPChars(Trim(lgObjRs("DRAWING_PATH"))) & """" & vbCrLf
						
						Response.Write ".frm1.txtItemSeq.value = """ & Trim(lgObjRs(3)) & """" & vbCrLf
						Response.Write ".frm1.txtChildItemQty.Text = """ & UniConvNumberDBToCompany(lgObjRs(9), 6, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtChildItemUnit.value = """ & ConvSPChars(Trim(lgObjRs(10))) & """" & vbCrLf
						Response.Write ".frm1.txtPrntItemQty.Text = """ & UniConvNumberDBToCompany(lgObjRs(7), 6, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtPrntItemUnit.value = """ & ConvSPChars(Trim(lgObjRs(8))) & """" & vbCrLf
						Response.Write ".frm1.txtSafetyLt.Text = """ & lgObjRs(12) & """" & vbCrLf
						Response.Write ".frm1.txtLossRate.Text = """ & UniConvNumberDBToCompany(lgObjRs(11), 2, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtRemark.value = """ & ConvSPChars(Trim(lgObjRs(15))) & """" & vbCrLf
						Response.Write ".frm1.txtValidFromDt1.text = """ & UNIDateClientFormat(lgObjRs(16)) & """" & vbCrLf
						Response.Write ".frm1.txtValidToDt1.text = """ & UNIDateClientFormat(lgObjRs(17)) & """" & vbCrLf
    
						If Ucase(Trim(lgObjRs(13))) = "F" Then
						    Response.Write ".frm1.rdoSupplyFlg1.Checked = True" & vbCrLf
						    Response.Write ".lgRdoOldVal2 = 1" & vbCrLf
						Else
						    Response.Write ".frm1.rdoSupplyFlg2.Checked = True" & vbCrLf
						    Response.Write ".lgRdoOldVal2 = 2" & vbCrLf
						End If

						Response.Write ".frm1.txtECNNo1.value = """ & ConvSPChars(Trim(lgObjRs("ECN_NO"))) & """" & vbCrLf
						Response.Write ".frm1.txtReasonCd1.value = """ & ConvSPChars(Trim(lgObjRs("REASON_CD"))) & """" & vbCrLf
						Response.Write ".frm1.txtReasonNm1.value = """ & ConvSPChars(Trim(lgObjRs("REASON_NM"))) & """" & vbCrLf
						Response.Write ".frm1.txtECNDesc1.value = """ & ConvSPChars(Trim(lgObjRs("ECN_DESC"))) & """" & vbCrLf
						
						Response.Write ".frm1.hBomType.value = """ & ConvSPChars(Trim(Request("txtChildBomNo"))) & """" & vbCrLf
    
				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				
		    End If
		    
		    Call SubCloseRs(lgObjRs) 
		    
   		Case "D"																		'☜: detail 조회 경우 
   			
			'-----------------------
			' Level Setting
			'-----------------------
			strLevel = ""
			iLevelCnt = CInt(Request("intLevel"))
	
			For i = 1 To iLevelCnt
				strLevel = strLevel & "."
			Next 
	
			strLevel = strLevel & Request("intLevel")
			
			'-------------------------------
			' detail 정보 query
			'-------------------------------
			lgStrSQL = ""
			
		    Call SubMakeSQLStatements("D",strPlantCd,strPrntItemCd,strChildItemCd,strPrntBomNo,strChildItemSeq)                                       '☜ : Make sql statements
		
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    	Call DisplayMsgBox("182700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				IntRetCD = -1
				Response.End 
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".frm1.txtItemSeq.Text = """ & Trim(lgObjRs(3)) & """" & vbCrLf
						Response.Write ".frm1.txtChildItemQty.Text = """ & UniConvNumberDBToCompany(lgObjRs(9), 6, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtChildItemUnit.value = """ & ConvSPChars(Trim(lgObjRs(10))) & """" & vbCrLf
						Response.Write ".frm1.txtPrntItemQty.Text = """ & UniConvNumberDBToCompany(lgObjRs(7), 6, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtPrntItemUnit.value = """ & ConvSPChars(Trim(lgObjRs(8))) & """" & vbCrLf
						Response.Write ".frm1.txtSafetyLt.Text = """ & lgObjRs(12) & """" & vbCrLf
						Response.Write ".frm1.txtLossRate.Text = """ & UniConvNumberDBToCompany(lgObjRs(11), 2, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtRemark.value = """ & ConvSPChars(lgObjRs(15)) & """" & vbCrLf
						Response.Write ".frm1.txtValidFromDt1.text = """ & UNIDateClientFormat(lgObjRs(16)) & """" & vbCrLf
						Response.Write ".frm1.txtValidToDt1.text = """ & UNIDateClientFormat(lgObjRs(17)) & """" & vbCrLf
       
						IF Ucase(lgObjRs(13)) = "F" Then
						    Response.Write ".frm1.rdoSupplyFlg1.Checked = True" & vbCrLf
						    Response.Write ".lgRdoOldVal2 = 1" & vbCrLf
						Else
						    Response.Write ".frm1.rdoSupplyFlg2.Checked = True" & vbCrLf
						    Response.Write ".lgRdoOldVal2 = 2" & vbCrLf
						End If
    
						Response.Write ".frm1.txtECNNo1.value = """ & ConvSPChars(Trim(lgObjRs("ECN_NO"))) & """" & vbCrLf
						Response.Write ".frm1.txtReasonCd1.value = """ & ConvSPChars(Trim(lgObjRs("REASON_CD"))) & """" & vbCrLf
						Response.Write ".frm1.txtReasonNm1.value = """ & ConvSPChars(Trim(lgObjRs("REASON_NM"))) & """" & vbCrLf
						Response.Write ".frm1.txtECNDesc1.value = """ & ConvSPChars(Trim(lgObjRs("ECN_DESC"))) & """" & vbCrLf

				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				
		    End If
		    
		    Call SubCloseRs(lgObjRs) 
		    
		Case "B"																	'☜: header 조회 경우 
			
			lgStrSQL = ""
			
		    Call SubMakeSQLStatements("H", strPlantCd, strItemCd, strBomNo,"","")           '☜ : Make sql statements

		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
				
				IntRetCD = -1
				Call SetErrorStatus()
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "Call parent.LookUpChildBomNoNotOk()" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End 
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".frm1.txtBomNo1.value = """ & ConvSPChars(Trim(lgObjRs(2))) & """" & vbCrLf
						Response.Write ".frm1.txtBOMDesc.value = """ & ConvSPChars(Trim(lgObjRs(3))) & """" & vbCrLf
						Response.Write ".frm1.txtDrawPath.value = """ & ConvSPChars(Trim(lgObjRs(7))) & """" & vbCrLf
				    Response.Write "End With" & vbCrLf
				    
				    Response.Write "Call parent.LookUpChildBomNoOk()" & vbCrLf
				Response.Write "</Script>" & vbCrLf

		    End If
		    
		    Call SubCloseRs(lgObjRs) 

			Call parent.LookUpItemByPlantOk()		
		Case "I"		

			lgStrSQL = ""
			Call SubMakeSQLStatements("I_CK",strPlantCd,strItemCd,"","","")           '☜ : Make sql statements
			
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				
				Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()

				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "Call parent.LookUpItemByPlantNotOk" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End 

		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".Frm1.txtItemCd1.Value = """ & ConvSPChars(Trim(lgObjRs("ITEM_CD"))) & """" & vbCrLf
						Response.Write ".Frm1.txtItemNm1.Value = """ & ConvSPChars(Trim(lgObjRs("ITEM_NM"))) & """" & vbCrLf
						Response.Write ".frm1.cboItemAcct.value = """ & ConvSPChars(Trim(lgObjRs("ITEM_ACCT"))) & """" & vbCrLf
						Response.Write ".frm1.txtItemAcctGrp.value = """ & ConvSPChars(Trim(lgObjRs("ITEM_ACCT_GRP"))) & """" & vbCrLf
						Response.Write ".frm1.txtSpec.value = """ & ConvSPChars(Trim(lgObjRs("SPEC"))) & """" & vbCrLf
						Response.Write ".frm1.txtPlantItemFromDt.text = """ & UNIDateClientFormat(lgObjRs("VALID_FROM_DT")) & """" & vbCrLf
						Response.Write ".frm1.txtPlantItemToDt.text = """ & UNIDateClientFormat(lgObjRs("VALID_TO_DT")) & """" & vbCrLf
						Response.Write ".frm1.txtProcType.value = """ & ConvSPChars(Trim(lgObjRs("PROCUR_TYPE"))) & """" & vbCrLf

						If iPos = "1" Then
							Response.Write ".frm1.txtChildItemUnit.value = """ & ConvSPChars(lgObjRs("BASIC_UNIT")) & """" & vbCrLf
							Response.Write ".frm1.txtChildItemQty.Text = ""1""" & vbCrLf
						End If
				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				
				If iPos = "0" and  ConvSPChars(lgObjRs(3)) = "P" Then
					Call DisplayMsgBox("182618", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
					Call SetErrorStatus()
					IntRetCD = -1
					Response.End 
				End If
							
			End If
			
			Call SubCloseRs(lgObjRs) 
		    
	End Select
    
End Sub    
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    
    Dim PrntKey
	Dim NodX
	Dim Node		    

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Call SubMakeSQLStatements("I_CK", strPlantCd, strItemCd, "", "", "")                                   '☆ : Make sql statements
    	
    If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then
    
        Call SetErrorStatus()
    Else

		'========================================================================
		' BOM 전개를 실시한다.
		'========================================================================
		strPlantCd = Trim(Request("txtPlantCd"))
		strItemCd = Trim(Request("txtItemCd"))
		strBomNo = Trim(Request("txtBomNo"))
		strItemAcct = lgObjRs("ITEM_ACCT")
		strItemAcctGrp = lgObjRs("ITEM_ACCT_GRP")
		strPhantomFlg = lgObjRs("PHANTOM_FLG")
		strProcurType = lgObjRs("PROCUR_TYPE")
		
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "PrntKey = """ & ConvSPChars(UCase(Trim(strItemCd))) & "|^|^|" & ConvSPChars(UCase(Trim(strBomNo))) & """" & vbCrLf

		'----------------------------------------------------
		'- Parent Node를 Setting
		'---------------------------------------------------
		
		Response.Write "With parent.frm1" & vbCrLf
			If Trim(strItemAcctGrp) = "1FINAL" Or Trim(strItemAcctGrp) = "2SEMI" Then
				If UCase(Trim(strPhantomFlg)) = "Y" Then
					Response.Write "Set NodX = .uniTree1.Nodes.Add(,,PrntKey,""" & ConvSPChars(UCase(Trim(strItemCd))) & """" & ",parent.C_PHANTOM, parent.C_PHANTOM)" & vbCrLf
					Response.Write "NodX.Expanded = True" & vbCrLf
				ElseIf UCase(Trim(strProcurType)) = "O" Then
					Response.Write "Set NodX = .uniTree1.Nodes.Add(,,PrntKey,""" & ConvSPChars(UCase(Trim(strItemCd))) & """" & ",parent.C_SUBCON, parent.C_SUBCON)" & vbCrLf
					Response.Write "NodX.Expanded = True" & vbCrLf
				ElseIf Trim(strItemAcctGrp) = "1FINAL" Then
					Response.Write "Set NodX = .uniTree1.Nodes.Add(,,PrntKey,""" & ConvSPChars(UCase(Trim(strItemCd))) & """" & ",parent.C_PROD, parent.C_PROD)" & vbCrLf
					Response.Write "NodX.Expanded = True" & vbCrLf
				Else                      ' "20"
					Response.Write "Set NodX = .uniTree1.Nodes.Add(,,PrntKey,""" & ConvSPChars(UCase(Trim(strItemCd))) & """" & ",parent.C_ASSEMBLY, parent.C_ASSEMBLY)" & vbCrLf
					Response.Write "NodX.Expanded = True" & vbCrLf
				End If
			Else
				Response.Write "Set NodX = .uniTree1.Nodes.Add(,,PrntKey,""" & ConvSPChars(UCase(Trim(strItemCd))) & """" & ",parent.C_MATL, parent.C_MATL)" & vbCrLf
				Response.Write "NodX.Expanded = True" & vbCrLf
			End If
			Response.Write "Set NodX = Nothing" & vbCrLf

		Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	
		Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
		Call SubCloseRs(lgObjRs) 
		
	End If
	'----------------------------
	' 하위품목 Node Setting
	'----------------------------
	lgStrSQL = ""
	strPlantCd = FilterVar(Trim(Request("txtPlantCd")), "''", "S")			
	
	Call SubMakeSQLStatements("M", strPlantCd, strSpId, "", "", "")					'☜ : Make sql statements

	If 	FncOpenRs("R", lgObjConn, lgObjRs, lgStrSQL, "X", "X") = False Then                    'If data not exists
		lgStrPrevKeyIndex = ""    
		IntRetCD = 1
	Else
		IntRetCD = 1
		Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

		Response.Write "<Script Language = VBScript>" & vbCrLf
				
			Response.Write "With parent.frm1.uniTree1" & vbCrLf
				
				Response.Write ".MousePointer = 11" & vbCrLf                   '⊙: 마우스 포인트 변화 
				Response.Write ".Indentation = 50" & vbCrLf                    '⊙: 부모트리와 자식트리 사이의 간격 

				Do While Not lgObjRs.EOF
				
					If lgObjRs(5) = "M" Then		' 제품일 경우 
						Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(lgObjRs(3)) & """" & ", parent.tvwChild, """ & ConvSPChars(lgObjRs(4)) & """, """ & ConvSPChars(Trim(lgObjRs(10))) & """" & ", parent.C_PROD, parent.C_PROD)" & vbCrLf
						Response.Write "Node.Expanded = True" & vbCrLf
					Elseif lgObjRs(5) = "A" Then		' 반제품, 재공품일 경우 
						Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(lgObjRs(3)) & """" & ", parent.tvwChild, """ & ConvSPChars(lgObjRs(4)) & """, """ & ConvSPChars(Trim(lgObjRs(10))) & """" & ", parent.C_ASSEMBLY, parent.C_ASSEMBLY)" & vbCrLf
						Response.Write "Node.Expanded = True" & vbCrLf
					Elseif lgObjRs(5) = "P" Then		' PHANTOM품일 경우 
						Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(lgObjRs(3)) & """" & ", parent.tvwChild, """ & ConvSPChars(lgObjRs(4)) & """, """ & ConvSPChars(Trim(lgObjRs(10))) & """" & ", parent.C_PHANTOM, parent.C_PHANTOM)" & vbCrLf
						Response.Write "Node.Expanded = True" & vbCrLf
					Elseif lgObjRs(5) = "E" Then		' 외주가공품인 경우 
						Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(lgObjRs(3)) & """" & ", parent.tvwChild, """ & ConvSPChars(lgObjRs(4)) & """, """ & ConvSPChars(Trim(lgObjRs(10))) & """" & ", parent.C_SUBCON, parent.C_SUBCON)" & vbCrLf
						Response.Write "Node.Expanded = True" & vbCrLf
					Else								' 원자재인 경우 
						Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(lgObjRs(3)) & """" & ", parent.tvwChild, """ & ConvSPChars(lgObjRs(4)) & """, """ & ConvSPChars(Trim(lgObjRs(10))) & """" & ", parent.C_MATL, parent.C_MATL)" & vbCrLf
						Response.Write "Node.Expanded = True" & vbCrLf
					End If

					lgObjRs.MoveNext
				
				Loop

				Response.Write ".MousePointer = 1" & vbCrLf
				Response.Write "Set Node = Nothing" & vbCrLf

			Response.Write "End With" & vbCrLf
		Response.Write "</Script>" & vbCrLf

    End If

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)                                                          '☜: Release RecordSSet
    
    lgStrSQL = ""
	'-------------------------
	' 생성된 temp table 삭제 
	'-------------------------
    lgStrSQL = "DELETE FROM P_BOM_FOR_EXPLOSION "
	lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & FilterVar(Trim(Request("txtPlantCd"))	,"''", "S")
	lgStrSQL = lgStrSQL & " AND USER_ID = " & strSpId
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

End Sub    

'============================================================================================================
' Name : SubBizSave
' Desc : Save Data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4)
    Dim iSelCount
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    Select Case pDataType
		Case "H"
			lgStrSQL = "SELECT * FROM P_BOM_HEADER "
            lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & pCode 
            lgStrSQL = lgStrSQL & " AND ITEM_CD = " & pCode1
            lgStrSQL = lgStrSQL & " AND BOM_NO = " & pCode2
            
        Case "D"
			lgStrSQL = "SELECT a.*, b.REASON_CD, dbo.ufn_GetCodeName('P1402', b.REASON_CD) REASON_NM, b.ECN_DESC "
			lgStrSQL = lgStrSQL & " FROM P_BOM_DETAIL a LEFT OUTER JOIN P_ECN_MASTER b ON a.ECN_NO = b.ECN_NO"
			lgStrSQL = lgStrSQL & " WHERE a.PRNT_PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND a.PRNT_ITEM_CD LIKE " & pCode1
			lgStrSQL = lgStrSQL & " AND a.CHILD_ITEM_CD = " & pCode2
			lgStrSQL = lgStrSQL & " AND a.PRNT_BOM_NO LIKE " & pCOde3
			lgStrSQL = lgStrSQL & " AND a.CHILD_ITEM_SEQ = " & pCOde4
		
		Case "H&D"
			lgStrSQL = "SELECT a.*, b.REASON_CD, dbo.ufn_GetCodeName('P1402', b.REASON_CD) REASON_NM, b.ECN_DESC, c.BOM_NO, c.DESCRIPTION, c.DRAWING_PATH "
			lgStrSQL = lgStrSQL & " FROM P_BOM_DETAIL a LEFT OUTER JOIN P_ECN_MASTER b ON a.ECN_NO = b.ECN_NO, P_BOM_HEADER c "
			lgStrSQL = lgStrSQL & " WHERE a.PRNT_PLANT_CD = c.PLANT_CD AND a.CHILD_ITEM_CD = c.ITEM_CD AND a.CHILD_BOM_NO = c.BOM_NO "
			lgStrSQL = lgStrSQL & " AND a.PRNT_PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND a.PRNT_ITEM_CD = " & pCode1
			lgStrSQL = lgStrSQL & " AND a.CHILD_ITEM_SEQ = " & pCode2
			lgStrSQL = lgStrSQL & " AND a.PRNT_BOM_NO = " & pCOde3
		Case "M"
			lgStrSQL = "SELECT a.*, b.ITEM_NM, b.PHANTOM_FLG, c.ITEM_ACCT FROM P_BOM_FOR_EXPLOSION a, B_ITEM b, B_ITEM_BY_PLANT c "
			lgStrSQL = lgStrSQL & " WHERE a.CHILD_ITEM_CD = b.ITEM_CD AND b.ITEM_CD = c.ITEM_CD AND a.PLANT_CD = c.PLANT_CD "
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode
			lgStrSQL = lgStrSQL & " AND a.USER_ID = " & pCode1
			lgStrSQL = lgStrSQL & " ORDER BY a.SEQ "
			
		Case "B_CK"
			lgStrSQL = "SELECT a.*, b.VALID_FROM_DT, b.VALID_TO_DT, c.ITEM_NM, c.SPEC, dbo.ufn_GetCodeName('P1001', b.ITEM_ACCT) ITEM_ACCT_NM, b.PROCUR_TYPE, b.ITEM_ACCT, c.BASIC_UNIT, dbo.ufn_GetItemAcctGrp(b.ITEM_ACCT) ITEM_ACCT_GRP "
			lgStrSQL = lgStrSQL & " FROM P_BOM_HEADER a, B_ITEM_BY_PLANT b, B_ITEM c "
			lgStrSQL = lgStrSQL & " WHERE a.PLANT_CD = b.PLANT_CD AND a.ITEM_CD = b.ITEM_CD AND b.ITEM_CD = c.ITEM_CD"
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.ITEM_CD = " & pCode1
			lgStrSQL = lgStrSQL & " AND a.BOM_NO= " & pCode2
		
		Case "BT_CK"
			lgStrSQL = "SELECT * FROM B_MINOR WHERE MAJOR_CD = 'P1401'"
			lgStrSQL = lgStrSQL & " AND MINOR_CD = " & pCode 
				
		Case "I_CK"
			lgStrSQL = "SELECT a.*, b.ITEM_NM, b.SPEC, dbo.ufn_GetCodeName('P1001', a.ITEM_ACCT) ITEM_ACCT_NM, b.PHANTOM_FLG, b.BASIC_UNIT , dbo.ufn_GetItemAcctGrp(a.ITEM_ACCT) ITEM_ACCT_GRP  "
			lgStrSQL = lgStrSQL & " FROM B_ITEM_BY_PLANT a, B_ITEM b "
			lgStrSQL = lgStrSQL & " WHERE a.ITEM_CD = b.ITEM_CD "
			lgStrSQL = lgStrSQL & " AND b.VALID_FLG = 'Y' "
			lgStrSQL = lgStrSQL & " AND a.PLANT_CD = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.ITEM_CD = " & pCode1
			
		Case "P_CK"
			lgStrSQL = "SELECT * FROM B_PLANT A, P_PLANT_CONFIGURATION B"
			lgStrSQL = lgStrSQL & " WHERE A.PLANT_CD = B.PLANT_CD"
			lgStrSQL = lgStrSQL & " AND B.ENG_BOM_FLAG = 'Y'"
			lgStrSQL = lgStrSQL & " AND A.PLANT_CD = " & pCode
			
		Case "PI_CK"
			lgStrSQL = "SELECT ITEM_CD, PROCUR_TYPE FROM  B_ITEM_BY_PLANT "
			lgStrSQL = lgStrSQL & " WHERE PLANT_CD = " & pCode 
			lgStrSQL = lgStrSQL & " AND ITEM_CD = " & pCode1
		
    End Select

	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()
	
	Dim strMsg_cd
    Dim strMsg_text
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	If Request("txtBomNo") = "" Then
		strBomNo = " "
	Else 
		strBomNo = Request("txtBomNo")
	End If
	
    With lgObjComm
        .CommandText = "usp_BOM_explode_main"
        .CommandType = adCmdStoredProc
        
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@srch_type",	advarXchar,adParamInput,2, Request("rdoSrchType"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	advarXchar,adParamInput,4, Request("txtPlantCd"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_item_cd",	advarXchar,adParamInput,18, Request("txtItemCd"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_bom_no",advarXchar,adParamInput,4,strBomNo)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_dt_s",	advarXchar,adParamInput,10,UniConvDate(Request("txtBaseDt")))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_qty",	adInteger,adParamInput,2,1)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",	advarXchar,adParamOutput,6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text",	advarXchar,adParamOutput,60)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id",	advarXchar,adParamOutput,13)

        lgObjComm.Execute ,, adExecuteNoRecords
        
    End With

    If  Err.number = 0 Then
        IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
        If  IntRetCD <> 1 then
            strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
            strMsg_text = lgObjComm.Parameters("@msg_text").Value
            strSpId = FilterVar(lgObjComm.Parameters("@user_id").Value, "''", "S")
            
            If strMsg_cd <> MSG_OK_STR Then
				Call DisplayMsgBox(strMsg_cd, vbInformation, strMsg_text, "", I_MKSCRIPT)
			End If
            IntRetCD = -1
            Exit Sub
        Else
			IntRetCD = 1
        End if
    Else           
        Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
        Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
        IntRetCD = -1
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
    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

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
		
Response.Write "<Script Language = VBScript>" & vbCrLf
	Select Case QueryType
		Case "A"                                                         '☜ : Query
			If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
				Response.Write "Call parent.DbQueryOk()" & vbCrLf
	        End If   
		Case "H"                                                         '☜ : Save
			If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
				Response.Write "Call parent.LookUpHdrOk()" & vbCrLf
	        End If   
		Case "D"                                                         '☜ : Delete
			If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
				Response.Write "Call parent.LookUpDtlOk" & vbCrLf
			End If   
		Case "I"
			If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
				Response.Write "Call parent.LookUpItemByPlantOk()" & vbCrLf
			End If   
		Case "B"
			If Trim(lgErrorStatus) = "NO" And IntRetCd <> -1 Then
				Response.Write "Call parent.LookUpChildBomNoOk()" & vbCrLf
			End If   
	End Select    
       
Response.Write "</Script>" & vbCrLf
%>