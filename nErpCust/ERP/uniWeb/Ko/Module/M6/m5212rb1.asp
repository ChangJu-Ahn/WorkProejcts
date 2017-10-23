<%

'********************************************************************************************************
'*  1. Module Name          : Procuremant																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m5212rb1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : B/L 내역참조 PopUp Transaction 처리용 ASP									*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  :																			*
'*  9. Modifier (First)     : Sun-joung Lee																*
'* 10. Modifier (Last)      : Jin-hyun Shin																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/22 : Coding Start												*
'********************************************************************************************************

%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
                                                                         
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3			   '☜ : DBAgent Parameter 선언 
	
	Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim iTotstrData
	
	Dim iItemCode
	Dim iPurGrp
	Dim iBeneficiary
	Dim iCurrency
	Dim iPayTerms
	Dim iIncoterms
	Dim iBlDocNo

	Dim strItemName
	Dim iFrPoint
	iFrPoint=0

    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist    = "No"
	
	iItemCode		= Request("txtItemCode")
	iPurGrp			= Request("txtPurGrp")
	iBeneficiary	= Request("txtBeneficiary")
	iCurrency		= Request("txtCurrency")
	iPayTerms		= Request("txtPayTerms")
	iIncoterms		= Request("txtIncoterms")
	iBlDocNo		= Request("txtBlDocNo")
	
    Call MakeHeaderData()
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Make Header data
' 2002/07/18 : Kim Jin Ha
'---------------------------------------------------------------------------------------------------------- 
Sub MakeHeaderData()
	
	Dim OBJ_PM6G119
	Dim E1_m_cc_hdr
	Const M418_E1_pur_grp = 101
	Const M418_E2_pur_grp_nm = 102
	Const M418_E1_beneficiary = 105
	Const M418_E1_beneficiary_nm = 106
	Const M418_E1_currency = 38
	Const M418_E1_pay_method = 17
	Const M418_E13_pay_method_nm = 115
	Const M418_E1_incoterms = 37
	Const M418_E1_bl_doc_no = 4
	
	Set OBJ_PM6G119 = Server.CreateObject("PM6G119.cMLkImportCcHdrS")

	If CheckSYSTEMError(Err,True) = True Then
		Set OBJ_PM6G119 = Nothing
		Exit Sub
	End If
	
	Call OBJ_PM6G119.M_LOOKUP_IMPORT_CC_HDR_SVR(gStrGlobalCollection, Request("txtCCNo"), E1_m_cc_hdr)
		
	If CheckSYSTEMError2(Err,True,"","","","","") = true then 
	   Set OBJ_PM6G119 = Nothing
	   Exit Sub
	End If

	Set OBJ_PM6G119 = Nothing									'☜: ComProxy UnLoad
	
	lgCurrency = ConvSPChars(E1_m_cc_hdr(M418_E1_currency))
	
	Call DisplayIncotermsNm(E1_m_cc_hdr(M418_E1_incoterms))
	
	iPurGrp			= ConvSPChars(E1_m_cc_hdr(M418_E1_pur_grp))
	iBeneficiary	= ConvSPChars(E1_m_cc_hdr(M418_E1_beneficiary))
	iCurrency		= ConvSPChars(E1_m_cc_hdr(M418_E1_currency))
	iPayTerms		= ConvSPChars(E1_m_cc_hdr(M418_E1_pay_method))
	iIncoterms		= ConvSPChars(E1_m_cc_hdr(M418_E1_incoterms))
	iBlDocNo		= ConvSPChars(E1_m_cc_hdr(M418_E1_bl_doc_no))
			
	Response.Write "<Script Language=VBScript>" & vbCr
	Response.Write "With parent.frm1"		 & vbCr
	Response.Write "	.txtPurGrp.Value			= """ & ConvSPChars(E1_m_cc_hdr(M418_E1_pur_grp))		& """" & vbCr		
	Response.Write "	.txtPurGrpNm.Value			= """ & ConvSPChars(E1_m_cc_hdr(M418_E2_pur_grp_nm))		& """" & vbCr		
	Response.Write "	.txtBeneficiary.Value		= """ & ConvSPChars(E1_m_cc_hdr(M418_E1_beneficiary))			& """" & vbCr		
	Response.Write "	.txtBeneficiaryNm.Value		= """ & ConvSPChars(E1_m_cc_hdr(M418_E1_beneficiary_nm))			& """" & vbCr		
	Response.Write "	.txtCurrency.Value			= """ & ConvSPChars(E1_m_cc_hdr(M418_E1_currency))		& """" & vbCr		
	Response.Write "	.txtPayTerms.Value			= """ & ConvSPChars(E1_m_cc_hdr(M418_E1_pay_method))		& """" & vbCr		
	Response.Write "	.txtPayTermsNm.Value		= """ & ConvSPChars(E1_m_cc_hdr(M418_E13_pay_method_nm))			& """" & vbCr 
	Response.Write "	.txtIncoterms.Value			= """ & ConvSPChars(E1_m_cc_hdr(M418_E1_incoterms))		& """" & vbCr 
	'Response.Write "	.txtIncotermsNm.Value		= """ & ConvSPChars(L_E7_b_minor5)			& """" & vbCr 
	Response.Write "	.txtBlDocNo.value			= """ & ConvSPChars(E1_m_cc_hdr(M418_E1_bl_doc_no))		& """" & vbCr
	Response.Write "End With" & vbCr
	Response.Write "</Script>" & vbCr
	
End Sub    
'=============================================================================================
Sub DisplayIncotermsNm(incoterms)
	
	Const iStrMajorCd = "B9006"
	
	Call SubOpenDB(lgObjConn)
	
	lgStrSQL = " SELECT minor_nm FROM B_MINOR " 
	lgStrSQL = lgStrSQL & " WHERE major_cd =  " & FilterVar(iStrMajorCd , "''", "S") & " AND minor_cd =  " & FilterVar(incoterms, "''", "S") & " "		
	
	If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X")  then
	
		Response.Write "<Script Language=VBScript>" & vbCr
		Response.Write "With parent.frm1" & vbCr
		Response.Write "	.txtIncotermsNm.value	=	""" & lgObjRs("minor_nm") & """ " & vbCr
		Response.Write "End With" & vbCr
		Response.Write "</Script>" & vbCr
			
		Call SubCloseRs(lgObjRs)  
	End if
	
	Call SubCloseDB(lgObjConn)
	
End Sub    
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    Const C_SHEETMAXROWS_D = 100 
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint     = C_SHEETMAXROWS_D * CLng(lgPageNo)
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
			lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
			PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")
	
    If iLoopCount < C_SHEETMAXROWS_D Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    
    SetConditionData = true

    If Not(rs1.EOF Or rs1.BOF) Then			' 거래처코드/명 
		strItemName = rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtItemCode")) Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			SetConditionData = FALSE
			exit function
		End If
	End If      
    
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

 
	Dim strVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
	Dim arrVal(2)														  '☜: 화면에서 팝업하여 query
		
	Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(1,2)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 

	UNISqlId(0) = "M5212RA101" 												' main query(spread sheet에 뿌려지는 query statement)
	UNISqlId(1) = "S0000QA001"	

	'--- 2004-08-20 by Byun Jee Hyun for UNICODE	
	
	UNIValue(0,0) = Trim(lgSelectList)                                          '☜: Select list
	
   	strVal = " "
    'strVal = strVal &  " AND D.CC_NO =  '" & FilterVar(Trim(UCase(Request("txtCcNo"))), " " , "SNM") & "'"
    
    If Len(iBlDocNo) Then
		strVal = strVal & " AND B.bl_doc_no = " & FilterVar(Trim(UCase(iBlDocNo)), " " , "S") & " "
    End if

    If Len(iItemCode) Then
		strVal =  strVal & " AND A.ITEM_CD = " & FilterVar(Trim(UCase(iItemCode)), " " , "S") & " "
		arrVal(1) = FilterVar(Trim(UCase(iItemCode)), " " , "S")
	End if
	
	If Len(iPurGrp) Then
		strVal = strVal & " AND B.PUR_GRP = " & FilterVar(Trim(UCase(iPurGrp)), " " , "S") & " "
	End If

	If Len(iBeneficiary) Then
		strVal = strVal & " AND B.Beneficiary = " & FilterVar(Trim(UCase(iBeneficiary)), " " , "S") & " "
	End If
	
	If Len(iIncoterms) Then
		strVal = strVal & " AND B.INCOTERMS = " & FilterVar(Trim(UCase(iIncoterms)), " " , "S") & " "
	End If

    If Len(Trim(iCurrency)) Then
		strVal = strVal & " AND B.CURRENCY = " & FilterVar(Trim(UCase(iCurrency)), " " , "S") & " "		
	End If		
	
	If Len(Trim(iPayTerms)) Then
		strVal = strVal & " AND B.PAY_METHOD = " & FilterVar(Trim(UCase(iPayTerms)), " " , "S") & " "		
	End If
	
	'---2003.07 TrackingNo 추가 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO = " & FilterVar(Trim(UCase(Request("txtTrackingNo"))), " " , "S") & "  "		
	End If
	
	UNIValue(0,1) = strVal    '	UNISqlId(0)의 두번째 ?에 입력됨	
 
    UNIValue(1,0) = FilterVar(Trim(UCase(iItemCode)), " " , "S")

    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                       '☜: set ADO read mode
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    If SetConditionData = false then Exit sub
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub

%>
  <Script Language=vbscript>
  With parent
  
	parent.frm1.txtItemName.value = "<%=ConvSPChars(strItemName)%>"
    If "<%=lgDataExist%>" = "Yes" Then
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.txtHItemCode.Value	 			= "<%=ConvSPChars(Request("txtItemCode"))%>"
			.frm1.txtHPurGrp.Value 				= "<%=ConvSPChars(Request("txtPurGrp"))%>"
			.frm1.txtHBeneficiary.Value 			= "<%=ConvSPChars(Request("txtBeneficiary"))%>"
			.frm1.txtHCurrency.Value 			    = "<%=ConvSPChars(Request("txtCurrency"))%>"
			.frm1.txtHPayTerms.Value 			    = "<%=ConvSPChars(Request("txtPayTerms"))%>"			
			.frm1.txtHIncoterms.Value 			= "<%=ConvSPChars(Request("txtIncoterms"))%>"			
			.frm1.txtBlDocNo.Value 				= "<%=ConvSPChars(Request("txtBlDocNo"))%>"
       End If
       
       .ggoSpread.Source  = .frm1.vspdData
       .frm1.vspdData.Redraw = False
       .ggoSpread.SSShowData "<%=iTotstrData%>","F"          '☜ : Display data
       
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,"<%=ConvSPChars(Request("txtCurrency"))%>",.GetKeyPos("A",22),"C","I","X","X")
       
       .lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       .DbQueryOk
       .frm1.vspdData.Redraw = True
    End If 
       
   End WIth
</Script>	
 	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
