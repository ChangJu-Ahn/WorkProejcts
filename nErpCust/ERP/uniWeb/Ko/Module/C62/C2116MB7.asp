<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>


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

'Call ServerMesgBox(strExpFlg, vbOKOnly, I_MKSCRIPT)

Call HideStatusWnd                                                               '☜: Hide Processing message

Call LoadBasisGlobalInf
Call LoadinfTB19029B("Q", "P", "NOCOOKIE", "MB")

lgErrorStatus     = "NO"
lgErrorPos        = ""                                                           '☜: Set to space
lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

Dim IntRetCD
Dim strPlantCd
Dim strItemCd
Dim strItemNm
Dim strBomNo
Dim strBaseDt
Dim strExpFlg
Dim strItemAcct
Dim strPhantomFlg
Dim strProcurType
Dim strPrntItemCd
Dim strPrntBomNo
Dim strChildItemSeq
Dim strChildBomNo
Dim strChildItemCd
DIm strLevel
Dim strSpId
Dim DtlFlg
Dim BaseDt	
Dim strECNNo
Dim StrECNReasonCd	
Dim strECNDescription
DtlFlg = 0
	
strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
strItemCd = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
strBomNo = FilterVar(Trim(Request("txtBomNo"))	, "''", "S")
BaseDt = FilterVar(UNIConvYYYYMMDDToDate(gAPDateFormat,"1900","01","01"), "''", "S")
strBaseDt = FilterVar(Trim(Request("txtBaseDt")), BaseDt , "D")
strExpFlg = FilterVar(Request("rdoSrchType"), "''", "S")

'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
	
Select Case lgOpModeCRUD
    Case CStr(UID_M0001)
		Call SubCreateCommandObject(lgObjComm)
			
		Call SubBizQuery("P_CK")
		Call SubBizQuery("I_CK")
		Call SubBizQuery("BT_CK")
		Call SubBizQuery("H_CK")

		Call SubBizBatch()
		Call SubBizQueryMulti()
			
		Call SubCloseCommandObject(lgObjComm)
			
    Case CStr(UID_M0002)								        							'☜: Header Query
		If strExpFlg = "" & FilterVar("1", "''", "S") & " " Or strExpFlg = "" & FilterVar("2", "''", "S") & " " Then		'bom no를 갖지 않는 역전개 
			Call SubBizQuery("P_CK")
			Call SubBizQuery("I_CK")
			Call SubBizQuery("BT_CK")
			Call SubBizQuery("H_CK")
		ElseIf strExpFlg = "" & FilterVar("3", "''", "S") & " " Or strExpFlg = "" & FilterVar("4", "''", "S") & " " Then		'bom no를 갖지 않는 역전개 
			strChildItemCd = strItemCd
			'strPrntItemCd = "'%%'"
			strPrntBomNo = strBomNo
			Call SubBizQuery("P_CK")
			Call SubBizQuery("I_CK")
			Call SubBizQuery("BT_CK")
			Call SubBizQuery("H_CK")
		End If
			
    Case CStr(UID_M0003)
		If strExpFlg = "" & FilterVar("2", "''", "S") & " " Then													'☜: Detail Query
			strItemCd = FilterVar(Trim(Request("txtChildItemCd"))	, "''", "S")
			strBomNo = FilterVar(Trim(Request("txtChildBomNo"))	, "''", "S")                                                     '☜: Detail Query
		Else 
			strItemCd = FilterVar(Trim(Request("txtPrntItemCd"))	, "''", "S")
			strBomNo = FilterVar(Trim(Request("txtPrntBomNo"))	, "''", "S")                                                     '☜: Detail Query
		End If
			
		strPrntItemCd =  FilterVar(Trim(Request("txtPrntItemCd"))	, "''", "S")
		strPrntBomNo =  FilterVar(Trim(Request("txtPrntBomNo"))	, "''", "S")
		strChildItemSeq = Request("intChildItemSeq")
		strChildBomNo =  FilterVar(Trim(Request("txtChildBomNo"))	, "''", "S")
		strChildItemCd =  FilterVar(Trim(Request("txtChildItemCd"))	, "''", "S")
		DtlFlg = 1
	
		Call SubBizQuery("I_CK")
	    Call SubBizQuery("HD_CK")
			
End Select
    
Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pOpCode)

	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
	
	Dim iIntCnt, iLevelCnt
	Select Case pOpCode
		Case "P_CK"
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
				Response.Write "Parent.Frm1.txtPlantNm.Value = """"" & vbCrLf
				Response.Write "Parent.Frm1.txtPlantCd.focus" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End
			
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "parent.Frm1.txtPlantNm.Value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			
			End If
		
			Call SubCloseRs(lgObjRs) 
			
		Case "I_CK"
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
						If DtlFlg <> 1 Then
							Response.Write ".Frm1.txtItemNm.Value = """ & ConvSPChars(lgObjRs("ITEM_NM")) & """" & vbCrLf
						End If
						Response.Write ".Frm1.txtItemCd1.Value = """ & ConvSPChars(lgObjRs("ITEM_CD")) & """" & vbCrLf
						Response.Write ".Frm1.txtItemNm1.Value = """ & ConvSPChars(lgObjRs("ITEM_NM")) & """" & vbCrLf
						Response.Write ".frm1.txtItemAcctNm.value = """ & ConvSPChars(lgObjRs("MINOR_NM")) & """" & vbCrLf
						Response.Write ".frm1.txtSpec.value = """ & ConvSPChars(lgObjRs("SPEC")) & """" & vbCrLf
						Response.Write ".frm1.txtPlantItemFromDt.text = """ & UNIDateClientFormat(lgObjRs("VALID_FROM_DT")) & """" & vbCrLf
						Response.Write ".frm1.txtPlantItemToDt.text = """ & UNIDateClientFormat(lgObjRs("VALID_TO_DT")) & """" & vbCrLf
						
				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf

				strItemNm = Trim(lgObjRs(77))

			End If
		
			Call SubCloseRs(lgObjRs) 

		Case "BT_CK"
			'------------------
			' bom type 체크 
			'------------------
			lgStrSQL = ""
			
			Call SubMakeSQLStatements("BT_CK",strBomNo,"","","","")           '☜ : Make sql statements
			
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    
				IntRetCD = -1
				
				Call DisplayMsgBox("182622", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()

				Response.Write "<Script Language = VBScript>" & vbCrLf
				Response.Write "parent.frm1.txtBomNo.focus" & vbCrLf
				Response.Write "</Script>" & vbCrLf
				Response.End							
		    Else
				IntRetCD = 1
			End If
		
			Call SubCloseRs(lgObjRs) 

		Case "H_CK"
	    
			lgStrSQL = ""
			
		    Call SubMakeSQLStatements("H_CK",strPlantCd,strItemCd,strBomNo,"","")           '☜ : Make sql statements
 
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
				
				Call SubBizQuery("I_CK")
				
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".frm1.txtBomNo1.value = """ & ConvSPChars(lgObjRs(2)) & """" & vbCrLf
						Response.Write ".frm1.txtBOMDesc.value = """ & ConvSPChars(lgObjRs(3)) & """" & vbCrLf
						Response.Write ".frm1.txtDrawNo.value = """ & ConvSPChars(lgObjRs(7)) & """" & vbCrLf
				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
     
		    End If
		    
		    Call SubCloseRs(lgObjRs) 

		Case "HD_CK"
		    
			'-----------------------
			' Level Setting
			'-----------------------
			strLevel = ""
			iLevelCnt = CInt(Request("intLevel"))
	
			For iIntCnt = 1 To iLevelCnt
				strLevel = strLevel & "."
			Next 
	
			strLevel = strLevel & Request("intLevel")
			
			'-------------------------------
			' detail 정보 query
			'-------------------------------
			lgStrSQL = ""
			
			Call SubBizQuery("H_CK")
			Call SubMakeSQLStatements("HD_CK",strPlantCd,strPrntItemCd,strChildItemSeq,strPrntBomNo,"")  
			
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
				Call SubBizQuery("D_CK")
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf

						'Response.Write ".frm1.txtBomNo1.value = """ & ConvSPChars(lgObjRs(23)) & """" & vbCrLf
 						'Response.Write ".frm1.txtBOMDesc.value = """ & ConvSPChars(lgObjRs(24)) & """" & vbCrLf
						'Response.Write ".frm1.txtDrawNo.value = """ & ConvSPChars(lgObjRs(25)) & """" & vbCrLf
						
						Response.Write ".frm1.txtLevel.value = """ & ConvSPChars(strLevel) & """" & vbCrLf
						Response.Write ".frm1.txtItemSeq.Text = """ & lgObjRs(3) & """" & vbCrLf
						Response.Write ".frm1.txtChildItemQty.Text = """ & UniConvNumberDBToCompany(lgObjRs(9), 4, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtChildItemUnit.value = """ & ConvSPChars(lgObjRs(10)) & """" & vbCrLf
						Response.Write ".frm1.txtPrntItemQty.Text = """ & UniConvNumberDBToCompany(lgObjRs(7), 4, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtPrntItemUnit.value = """ & ConvSPChars(lgObjRs(8)) & """" & vbCrLf
						Response.Write ".frm1.txtSafetyLt.Text = """ & lgObjRs(12) & """" & vbCrLf
						Response.Write ".frm1.txtLossRate.Text = """ & UniConvNumberDBToCompany(lgObjRs(11), 2, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtRemark.value = """ & ConvSPChars(lgObjRs(15)) & """" & vbCrLf
						Response.Write ".frm1.txtValidFromDt1.text = """ & UNIDateClientFormat(lgObjRs(16)) & """" & vbCrLf
						Response.Write ".frm1.txtValidToDt1.text = """ & UNIDateClientFormat(lgObjRs(17)) & """" & vbCrLf
    
						Response.Write ".frm1.txtECNNo.value = """ & ConvSPChars(lgObjRs("ECN_NO")) & """" & vbCrLf
						Response.Write ".frm1.txtECNReasonCd.value = """ & ConvSPChars(lgObjRs("REASON_NM")) & """" & vbCrLf
						Response.Write ".frm1.txtECNDescription.value = """ & ConvSPChars(lgObjRs("ECN_DESC")) & """" & vbCrLf

						IF UCase(lgObjRs(13)) = "F" Then
						    Response.Write ".frm1.rdoSupplyFlg1.Checked = True" & vbCrLf
						Else
						    Response.Write ".frm1.rdoSupplyFlg2.Checked = True" & vbCrLf
						End If
				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
		    End If
		    
		    Call SubCloseRs(lgObjRs) 
		    
   		Case "D_CK"																		'☜: detail 조회 경우 
   			
			'-----------------------
			' Level Setting
			'-----------------------
			strLevel = ""
			iLevelCnt = CInt(Request("intLevel"))
				
			For iIntCnt = 1 To iLevelCnt
				strLevel = strLevel & "."
			Next 
	
			strLevel = strLevel & Request("intLevel")
			
			'-------------------------------
			' detail 정보 query
			'-------------------------------
			lgStrSQL = ""
			
		    Call SubMakeSQLStatements("D_CK",strPlantCd,strPrntItemCd,strChildItemCd,strPrntBomNo,strChildItemSeq)                                       '☜ : Make sql statements
		
		    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		    	Call DisplayMsgBox("182700", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
				Call SetErrorStatus()
				IntRetCD = -1
				Response.End 
		    Else
				IntRetCD = 1
				Response.Write "<Script Language = VBScript>" & vbCrLf
					Response.Write "With Parent" & vbCrLf
						Response.Write ".frm1.txtLevel.value = """ & ConvSPChars(strLevel) & """" & vbCrLf
						Response.Write ".frm1.txtItemSeq.Text = """ & lgObjRs(3) & """" & vbCrLf
						Response.Write ".frm1.txtChildItemQty.Text = """ & UniConvNumberDBToCompany(lgObjRs(9), 4, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtChildItemUnit.value = """ & ConvSPChars(lgObjRs(10)) & """" & vbCrLf
						Response.Write ".frm1.txtPrntItemQty.Text = """ & UniConvNumberDBToCompany(lgObjRs(7), 4, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtPrntItemUnit.value = """ & ConvSPChars(lgObjRs(8)) & """" & vbCrLf
						Response.Write ".frm1.txtSafetyLt.Text = """ & lgObjRs(12) & """" & vbCrLf
						Response.Write ".frm1.txtLossRate.Text = """ & UniConvNumberDBToCompany(lgObjRs(11), 2, 3, "", 0) & """" & vbCrLf
						Response.Write ".frm1.txtRemark.value = """ & ConvSPChars(lgObjRs(15)) & """" & vbCrLf
						Response.Write ".frm1.txtValidFromDt1.text = """ & UNIDateClientFormat(lgObjRs(16)) & """" & vbCrLf
						Response.Write ".frm1.txtValidToDt1.text = """ & UNIDateClientFormat(lgObjRs(17)) & """" & vbCrLf
       
						Response.Write ".frm1.txtECNNo.value = """ & ConvSPChars(lgObjRs("ECN_NO")) & """" & vbCrLf
						Response.Write ".frm1.txtECNReasonCd.value = """ & ConvSPChars(lgObjRs("REASON_NM")) & """" & vbCrLf
						Response.Write ".frm1.txtECNDescription.value = """ & ConvSPChars(lgObjRs("ECN_DESC")) & """" & vbCrLf

						IF UCase(lgObjRs(13)) = "F" Then
						    Response.Write ".frm1.rdoSupplyFlg1.Checked = True" & vbCrLf
						Else
						    Response.Write ".frm1.rdoSupplyFlg2.Checked = True" & vbCrLf
						End If
				    Response.Write "End With" & vbCrLf
				Response.Write "</Script>" & vbCrLf
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
    Call SubMakeSQLStatements("I_CK",strPlantCd,strItemCd,"","","")                                   '☆ : Make sql statements
    	
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then   
        Call SetErrorStatus()
    Else
		'========================================================================
		' BOM 전개를 실시한다.
		'========================================================================
		strPlantCd = Trim(UCase(Request("txtPlantCd")))
		strItemCd = Trim(UCase(Request("txtItemCd")))
		strBomNo = Trim(UCase(Request("txtBomNo")))
		
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "PrntKey = """ & ConvSPChars(strItemCd) & "|^|^|" & ConvSPChars(strBomNo) & """" & vbCrLf

		'----------------------------------------------------
		'- Parent Node를 Setting
		'---------------------------------------------------
		strItemNm = Trim(lgObjRs("ITEM_NM"))
		strItemAcct = Trim(lgObjRs("ITEM_ACCT"))
		strPhantomFlg = Trim(lgObjRs("PHANTOM_FLG"))
		strProcurType = Trim(lgObjRs("PROCUR_TYPE"))

		Response.Write "With parent.frm1" & vbCrLf
			If strItemAcct < "30" Then
			
				If UCase(strPhantomFlg) = "Y" Then
					Response.Write "Set NodX = .uniTree1.Nodes.Add(,,PrntKey, """ & ConvSPChars(strItemCd) & "    (" & ConvSPChars(strItemNm) & ")"", parent.C_PHANTOM, parent.C_PHANTOM)" & vbCrLf
					Response.Write "NodX.Expanded = True" & vbCrLf
				ElseIf UCase(strProcurType) = "O" Then
					Response.Write "Set NodX = .uniTree1.Nodes.Add(,,PrntKey, """ & ConvSPChars(strItemCd) & "    (" & ConvSPChars(strItemNm) & ")"", parent.C_SUBCON, parent.C_SUBCON)" & vbCrLf
					Response.Write "NodX.Expanded = True" & vbCrLf
				ElseIf strItemAcct = "10" Then
					Response.Write "Set NodX = .uniTree1.Nodes.Add(,,PrntKey, """ & ConvSPChars(strItemCd) & "    (" & ConvSPChars(strItemNm) & ")"", parent.C_PROD, parent.C_PROD)" & vbCrLf
					Response.Write "NodX.Expanded = True" & vbCrLf
				Else
					Response.Write "Set NodX = .uniTree1.Nodes.Add(,,PrntKey, """ & ConvSPChars(strItemCd) & "    (" & ConvSPChars(strItemNm) & ")"", parent.C_ASSEMBLY, parent.C_ASSEMBLY)" & vbCrLf
					Response.Write "NodX.Expanded = True" & vbCrLf
				End If
			Else
				Response.Write "Set NodX = .uniTree1.Nodes.Add(,,PrntKey, """ & ConvSPChars(strItemCd) & "    (" & ConvSPChars(strItemNm) & ")"", parent.C_MATL, parent.C_MATL)" & vbCrLf
				Response.Write "NodX.Expanded = True" & vbCrLf
			End If
			Response.Write "NodX.Expanded = True" & vbCrLf
		Response.Write "End With" & vbCrLf
	Response.Write "</Script>" & vbCrLf
	
	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
	Call SubCloseRs(lgObjRs) 
	
	End If

	'----------------------------
	' 하위품목 Node Setting
	'----------------------------
	lgStrSQL = ""
	strPlantCd = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")			
	
	Call SubMakeSQLStatements("M",strPlantCd,strSpId,"","","")					'☜ : Make sql statements

	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
		lgStrPrevKeyIndex = ""    
		IntRetCD = -1
'		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
		Call SetErrorStatus()
		Response.End
			
	Else
		IntRetCD = 1
		Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)
		Response.Write "<Script Language = VBScript>" & vbCrLf
			Response.Write "With parent.frm1.uniTree1" & vbCrLf
				
				Response.Write ".MousePointer = 11" & vbCrLf		'⊙: 마우스 포인트 변화 
				Response.Write ".Indentation = 50" & vbCrLf			'⊙: 부모트리와 자식트리 사이의 간격 

				Do While Not lgObjRs.EOF
				
					If lgObjRs(5) = "M" Then		' 제품일 경우 
						Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(lgObjRs(3)) & """, parent.tvwChild, """ & ConvSPChars(lgObjRs(4)) & """, """ & ConvSPChars(Trim(lgObjRs(10))) & "    (" & ConvSPChars(Trim(lgObjRs(24))) & ")"", parent.C_PROD, parent.C_PROD)" & vbCrLf
						Response.Write "Node.Expanded = True" & vbCrLf
					Elseif lgObjRs(5) = "A" Then		' 반제품일 경우 
						Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(lgObjRs(3)) & """, parent.tvwChild, """ & ConvSPChars(lgObjRs(4)) & """, """ & ConvSPChars(Trim(lgObjRs(10))) & "    (" & ConvSPChars(Trim(lgObjRs(24))) & ")"", parent.C_ASSEMBLY, parent.C_ASSEMBLY)" & vbCrLf
						Response.Write "Node.Expanded = True" & vbCrLf
					Elseif lgObjRs(5) = "P" Then		' PHANTOM품일 경우 
						Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(lgObjRs(3)) & """, parent.tvwChild, """ & ConvSPChars(lgObjRs(4)) & """, """ & ConvSPChars(Trim(lgObjRs(10))) & "    (" & ConvSPChars(Trim(lgObjRs(24))) & ")"", parent.C_PHANTOM, parent.C_PHANTOM)" & vbCrLf
						Response.Write "Node.Expanded = True" & vbCrLf
					Elseif lgObjRs(5) = "E" Then		' 외주가공품인 경우 
						Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(lgObjRs(3)) & """, parent.tvwChild, """ & ConvSPChars(lgObjRs(4)) & """, """ & ConvSPChars(Trim(lgObjRs(10))) & "    (" & ConvSPChars(Trim(lgObjRs(24))) & ")"", parent.C_SUBCON, parent.C_SUBCON)" & vbCrLf
						Response.Write "Node.Expanded = True" & vbCrLf
					Else																		'원자재인 경우 
						Response.Write "Set Node = .Nodes.Add(""" & ConvSPChars(lgObjRs(3)) & """, parent.tvwChild, """ & ConvSPChars(lgObjRs(4)) & """, """ & ConvSPChars(Trim(lgObjRs(10))) & "    (" & ConvSPChars(Trim(lgObjRs(24))) & ")"", parent.C_MATL, parent.C_MATL)" & vbCrLf
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
    lgStrSQL = "DELETE FROM p_bom_for_explosion "
	lgStrSQL = lgStrSQL & " WHERE plant_cd = " & FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	lgStrSQL = lgStrSQL & " AND user_id = " & strSpId
    
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
        Case "D_CK"
			lgStrSQL = "SELECT a.*, A.ECN_NO, d.ECN_DESC, d.REASON_CD, dbo.ufn_GetCodeName(" & FilterVar("P1402", "''", "S") & ", d.REASON_CD) REASON_NM "
			lgStrSQL = lgStrSQL & " FROM P_BOM_DETAIL a LEFT OUTER JOIN P_ECN_MASTER d ON a.ECN_NO = d.ECN_NO"
			lgStrSQL = lgStrSQL & " WHERE a.prnt_plant_cd = " & pCode
			lgStrSQL = lgStrSQL & " AND a.prnt_item_cd LIKE " & pCode1
			lgStrSQL = lgStrSQL & " AND a.child_item_cd = " & pCode2
			lgStrSQL = lgStrSQL & " AND a.prnt_bom_no LIKE " & pCOde3
			lgStrSQL = lgStrSQL & " AND a.child_item_seq = " & pCOde4
		
		Case "HD_CK"
			lgStrSQL = "SELECT a.* , b.bom_no, b.description ,b.drawing_path,  "
			lgStrSQL = lgStrSQL & " A.ECN_NO, d.ECN_DESC, d.REASON_CD, dbo.ufn_GetCodeName(" & FilterVar("P1402", "''", "S") & ", d.REASON_CD) REASON_NM "
			lgStrSQL = lgStrSQL & " FROM P_BOM_DETAIL a LEFT OUTER JOIN P_ECN_MASTER d ON a.ECN_NO = d.ECN_NO, p_bom_header b "
			lgStrSQL = lgStrSQL & " WHERE a.prnt_plant_cd = b.plant_cd and a.child_item_cd = b.item_cd and a.child_bom_no = b.bom_no "
			lgStrSQL = lgStrSQL & " AND a.prnt_plant_cd = " & pCode
			lgStrSQL = lgStrSQL & " AND a.prnt_item_cd = " & pCode1
			lgStrSQL = lgStrSQL & " AND a.child_item_seq = " & pCode2
			lgStrSQL = lgStrSQL & " AND a.prnt_bom_no = " & pCOde3
			
		Case "M"
			lgStrSQL = "SELECT a.*, b.item_nm, b.phantom_flg, c.item_acct "
			lgStrSQL = lgStrSQL & " FROM p_bom_for_explosion a, b_item b, b_item_by_plant c "
			lgStrSQL = lgStrSQL & " WHERE a.child_item_Cd = b.item_cd and b.item_cd = c.item_cd and a.plant_cd = c.plant_cd "
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode
			lgStrSQL = lgStrSQL & " AND a.user_id = " & pCode1
			lgStrSQL = lgStrSQL & " ORDER BY a.SEQ "

		Case "H_CK"
			lgStrSQL = "SELECT a.*, b.valid_from_dt, b.valid_to_dt, c.item_nm, c.spec, d.minor_nm  "
			lgStrSQL = lgStrSQL & " FROM p_bom_header a, b_item_by_plant b, b_item c, b_minor d "
			lgStrSQL = lgStrSQL & " WHERE a.plant_cd = b.plant_cd and a.item_cd = b.item_cd and b.item_cd = c.item_cd and b.item_acct = d.minor_cd and d.major_cd=" & FilterVar("p1001", "''", "S") & " "
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.item_cd = " & pCode1
			lgStrSQL = lgStrSQL & " AND a.bom_no= " & pCode2
			
		Case "BT_CK"
			lgStrSQL = "SELECT * FROM b_minor WHERE major_cd = " & FilterVar("P1401", "''", "S") & " "
			lgStrSQL = lgStrSQL & " AND minor_cd = " & pCode 
			
		Case "I_CK"
			lgStrSQL = "SELECT a.*, b.item_nm, b.spec, c.minor_nm, b.phantom_flg"
			lgStrSQL = lgStrSQL & " FROM b_item_by_plant a, b_item b, b_minor c "
			lgStrSQL = lgStrSQL & " WHERE a.item_cd =b.item_cd and c.minor_cd = a.item_acct and c.major_cd = " & FilterVar("p1001", "''", "S") & ""
			lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode 
			lgStrSQL = lgStrSQL & " AND a.item_cd = " & pCode1
			
		Case "P_CK"
			lgStrSQL = "SELECT * FROM b_plant where plant_cd = " & pCode 
		
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
	
	strBomNo = Request("txtBomNo")
	
    With lgObjComm
        .CommandText = "usp_BOM_explode_main"
        .CommandType = adCmdStoredProc
        
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@srch_type",	adVarChar,adParamInput,2, Request("rdoSrchType"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@plant_cd",	adVarChar,adParamInput,4, Request("txtPlantCd"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_item_cd",	adVarChar,adParamInput,18, Request("txtItemCd"))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@par_bom_no",adVarChar,adParamInput,4,strBomNo)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_dt_s",	adVarChar,adParamInput,10,UniConvDate(Request("txtBaseDt")))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@base_qty",	adInteger,adParamInput,2,1)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",	adVarChar,adParamOutput,6)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_text",	adVarChar,adParamOutput,60)
        lgObjComm.Parameters.Append lgObjComm.CreateParameter("@user_id",	adVarChar,adParamOutput,13)

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
%>	
