<%@ LANGUAGE=VBSCript%>
<%Option Explicit    %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE","MB")

'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        :
'*  3. Program ID           : m9111ma1
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : PM9G111(Maint)
'							  PM9G112(확정)
'*  7. Modified date(First) : 2002/12/06
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*
'*
'*
'*
'* 14. Business Logic of m9111ma1(재고이동요청)
'**********************************************************************************************
Dim lgOpModeCRUD

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData
Dim iStrFacilityCd
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim iLngMaxRow		' 현재 그리드의 최대Row
Dim iLngRow
Dim GroupCount
Dim lgCurrency
Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 
Dim lgDataExist
Dim lgPageNo
Dim lgMaxCount
Dim strFlag

	Const C_SHEETMAXROWS_D  = 100

	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

	Call HideStatusWnd                                                               '☜: Hide Processing message
	'------ Developer Coding part (Start ) ------------------------------------------------------------------

	'------ Developer Coding part (End   ) ------------------------------------------------------------------

	lgOpModeCRUD  = Request("txtMode")

										                                              '☜: Read Operation Mode (CRUD)
	Select Case lgOpModeCRUD
		Case CStr(UID_M0001)                                                         '☜: Query
			Call  SubBizQueryMulti()
		Case CStr(UID_M0002)                                                         '☜: Save,Update
			Call SubBizSaveMulti()
		Case CStr(UID_M0003)
			Call SubBizDelete()
	End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status
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

	On Error Resume Next
	Err.Clear
	Dim pPY5G110																	'☆ : 입력/수정용 Component Dll 사용 변수 
	Dim Y5_Y_Facility, iCommandSent
	Dim iIntFlgMode

	iStrFacilityCd = Trim(Request("txtFacility_Cd"))


	Const PY51_Y5_FACILITY_CD		= 0


	iIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	'-----------------------
	'Data manipulate area
	'-----------------------
	Redim Y5_Y_Facility(PY51_Y5_FACILITY_CD)


	Y5_Y_Facility(PY51_Y5_FACILITY_CD	)		= (Trim(Request("txtFacility_Cd")))

	iCommandSent = "DELETE"

	Set pPY5G110 = Server.CreateObject("PY5G110.cBMngFacility")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

	Call pPY5G110.B_MANAGE_FACILITY(gStrGlobalCollection, iCommandSent, Y5_Y_Facility)

	Select Case Trim(Cstr(Err.Description))
		Case "B_MESSAGE" & Chr(11) & "970023"

		Case Else
			If CheckSYSTEMError(Err, True) = True Then
				Set pPY5G110 = Nothing															'☜: Unload Component
				Response.End
			End If
	End Select

	Set pPY5G110 = Nothing															'☜: Unload Component

	Response.Write "<Script Language=vbscript>" & vbCr
' 	Response.Write "parent.frm1.hFacility_CD.value = """ & iStrFacilityCd & """" & vbCr
	Response.Write "       Parent.DbDeleteOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "</Script>"		& vbCr

End Sub

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()

	On Error Resume Next

	iStrFacilityCd = Trim(Request("txtFacility_Cd"))
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = C_SHEETMAXROWS_D                           '☜ : 한번에 가져올수 있는 데이타 건수 
	lgDataExist     = "No"
	iLngMaxRow	   = CDbl(lgMaxCount) * CDbl(lgPageNo) + 1

	lgStrPrevKey = Request("lgStrPrevKey")


	Call FixUNISQLData()
	Call QueryData()

	'====================
	'Call PO_DTL List
	'====================
	'-----------------------
	'Result data display area
	'-----------------------
	if GroupCount > 0 then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "	With parent " & vbCr
		Response.Write "	.ggoSpread.Source       = .frm1.vspdData "			& vbCr
'		Response.Write "	.frm1.vspdData.focus "			& vbCr
		
		Response.Write "	.ggoSpread.SSShowData     """ & istrData	 & """" & vbCr
		Response.Write "	.lgPageNo  = """ & lgPageNo   & """" & vbCr
		Response.Write "	.frm1.hFacility_CD.value = """ & iStrFacilityCd & """" & vbCr
	
		Response.Write " 	.DbQueryOk "	& vbCr
		Response.Write "	End With "		& vbCr
		Response.Write "</Script>"		& vbCr
	End if

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Dim strFacility_Cd
	Dim strFacility_Accnt
	Dim strUse_Yn


	Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(2, 3)

	UNISqlId(0) = "P5110P5AA"
	UNISqlId(1) = "P5110P500"


	IF Request("txtFacility_Cd") = "" Then
		strFacility_Cd = "|"
	ELSE
		strFacility_Cd = FilterVar((Trim(Request("txtFacility_Cd"))),"''","S")
	END IF

	IF Request("txtFacility_Accnt") = "" Then
		strFacility_Accnt = "|"
	ELSE
		strFacility_Accnt = FilterVar((Trim(Request("txtFacility_Accnt"))),"''","S")
	END IF

	IF Request("txtUse_Yn") = "" Then
		strUse_Yn = "|"
	ELSE
		strUse_Yn = FilterVar((Trim(Request("txtUse_Yn"))),"''","S")
	END IF


	UNIValue(0, 0) = FilterVar((Trim(Request("txtFacility_Cd"))),"''","S")

	UNIValue(1, 0) = "^"
	UNIValue(1, 1) = strFacility_Accnt
	UNIValue(1, 2) = strFacility_Cd
	UNIValue(1, 3) = strUse_Yn


	UNILock = DISCONNREAD :	UNIFlag = "1"


End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
	Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
	Dim iStr

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	Set lgADF   = Nothing

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
	    Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If


	if Request("txtType") = "B" Then							'☜ : 디테일 검색 
			Response.Write "<Script Language=vbscript>" & vbCr


			Response.Write "parent.frm1.CboFacility_Accnt.value 	= """ & ConvSPChars(rs1("Facility_Accnt")) & """" & vbCr
			Response.Write "parent.frm1.txtFacility_Cd.value 		= """ & ConvSPChars(rs1("Facility_Cd")) & """" & vbCr
			Response.Write "parent.frm1.txtFacility_Nm.value 		= """ & ConvSPChars(rs1("Facility_Nm")) & """" & vbCr
			Response.Write "parent.frm1.txtItemGroupCd1.Value		= """ & ConvSPChars(rs1("Facility_lvl1")) & """" & vbCr
			Response.Write "parent.frm1.txtItemGroupCd2.Value		= """ & ConvSPChars(rs1("Facility_lvl2")) & """" & vbCr
			Response.Write "parent.frm1.txtModel_Sts.Value			= """ & ConvSPChars(rs1("MODEL_STS")) & """" & vbCr
			Response.Write "parent.frm1.txtPress_Power.value		= """ & UNINumClientFormat(rs1("PRESS_POWER"), ggQty.DecPoint, 0) & """" & vbCr
			Response.Write "parent.frm1.txtPlant_Sts.value			= """ & ConvSPChars(rs1("PLANT_STS")) & """" & vbCr
			Response.Write "parent.frm1.txtProd_Amt.value			= """ & UNINumClientFormat(rs1("PROD_AMT"), ggAmtOfMoney.DecPoint, 0)	& """" & vbCr
			Response.Write "parent.frm1.txtCondAsstNo1.Value		= """ & ConvSPChars(rs1("ASST_CD1")) & """" & vbCr
			Response.Write "parent.frm1.txtCondAsstNo2.Value		= """ & ConvSPChars(rs1("ASST_CD2")) & """" & vbCr
			Response.Write "parent.frm1.txtCondAsstNm1.Value		= """ & ConvSPChars(rs1("ASST_NM1")) & """" & vbCr
			Response.Write "parent.frm1.txtCondAsstNm2.Value		= """ & ConvSPChars(rs1("ASST_NM2")) & """" & vbCr
			Response.Write "parent.frm1.txtPlantCd.Value			= """ & ConvSPChars(rs1("SET_PLANT")) & """" & vbCr
			Response.Write "parent.frm1.txtPlantNm.Value			= """ & ConvSPChars(rs1("Plant_Nm")) & """" & vbCr
			Response.Write "parent.frm1.txtSet_Place.Value			= """ & ConvSPChars(rs1("SET_PLACE")) & """" & vbCr
			Response.Write "parent.frm1.txtConWcNm.Value			= """ & ConvSPChars(rs1("WC_NM")) & """" & vbCr
			Response.Write "parent.frm1.CboUse_Yn.Value				= """ & ConvSPChars(rs1("USE_YN")) & """" & vbCr
			Response.Write "parent.frm1.txtSetCoCd.Value			= """ & ConvSPChars(rs1("SET_CO")) & """" & vbCr
			Response.Write "parent.frm1.txtSetDt.text				= """ & UNIDateClientFormat(rs1("SET_DT")) & """" & vbCr
			Response.Write "parent.frm1.txtPurCoCd.Value			= """ & ConvSPChars(rs1("PUR_CO")) & """" & vbCr
			Response.Write "parent.frm1.txtPurDt.Text				= """ & UNIDateClientFormat(rs1("PUR_DT")) & """" & vbCr
			Response.Write "parent.frm1.txtProdCoCd.Value			= """ & ConvSPChars(rs1("PROD_CO")) & """" & vbCr
			Response.Write "parent.frm1.txtSetCoNm.Value			= """ & ConvSPChars(rs1("SetCoNm")) & """" & vbCr
			Response.Write "parent.frm1.txtPurCoNm.Value			= """ & ConvSPChars(rs1("PurCoNm")) & """" & vbCr
			Response.Write "parent.frm1.txtProdCoNm.Value			= """ & ConvSPChars(rs1("ProdCoNm")) & """" & vbCr
			Response.Write "parent.frm1.txtEquip_Area.Value  		= """ & ConvSPChars(rs1("EQUIP_AREA")) & """" & vbCr
			Response.Write "parent.frm1.txtProdNo.Value				= """ & ConvSPChars(rs1("PROD_NO")) & """" & vbCr
			Response.Write "parent.frm1.txtUseVolt.Value			= """ & ConvSPChars(rs1("USE_VOLT")) & """" & vbCr
			Response.Write "parent.frm1.txtProd_Flag.Value			= """ & ConvSPChars(rs1("PROD_FLAG")) & """" & vbCr
			Response.Write "parent.frm1.txtUse_Amount.Value			= """ & ConvSPChars(rs1("USE_AMOUNT")) & """" & vbCr
			Response.Write "parent.frm1.txtLife_Cycle.value			= """ & UNINumClientFormat(rs1("LIFE_CYCLE"), 0, 0) & """" & vbCr
			Response.Write "parent.frm1.txtMoter_Type.Value			= """ & ConvSPChars(rs1("MOTER_TYPE")) & """" & vbCr
			Response.Write "parent.frm1.txtChk_End_dt.Text			= """ & UNIDateClientFormat(rs1("CHK_END_DT")) & """" & vbCr
			Response.Write "parent.frm1.txtOil_Spec1.Value			= """ & ConvSPChars(rs1("OIL_SPEC1")) & """" & vbCr
			Response.Write "parent.frm1.txtChk_Prd1.Value			= """ & UNINumClientFormat(rs1("CHK_PRD1"), 0, 0) & """" & vbCr
			Response.Write "parent.frm1.txtMoter_qty.value			= """ & UNINumClientFormat(rs1("MOTER_QTY"), ggQty.DecPoint, 0) & """" & vbCr
			Response.Write "parent.frm1.txtRep_End_dt.Text			= """ & UNIDateClientFormat(rs1("REP_END_DT")) & """" & vbCr
			Response.Write "parent.frm1.txtOil_Spec2.Value			= """ & ConvSPChars(rs1("OIL_SPEC2")) & """" & vbCr
			Response.Write "parent.frm1.txtChk_Prd2.value			= """ & UNINumClientFormat(rs1("CHK_PRD2"), 0, 0) & """" & vbCr
			Response.Write "parent.frm1.txtMoter_Power.value		= """ & UNINumClientFormat(rs1("MOTER_POWER"), ggQty.DecPoint, 0) & """" & vbCr
			Response.Write "parent.frm1.txtJng_End_dt.Text			= """ & UNIDateClientFormat(rs1("JNG_END_DT")) & """" & vbCr
			Response.Write "parent.frm1.txtOil_Spec3.Value			= """ & ConvSPChars(rs1("OIL_SPEC3")) & """" & vbCr
			Response.Write "parent.frm1.txtEmp_no.Value				= """ & ConvSPChars(rs1("EMP_NO")) & """" & vbCr
' 			Response.Write "parent.frm1.txtName.Value				= """ & ConvSPChars(rs1("EMP_Nm")) & """" & vbCr
			Response.Write "parent.frm1.txtMoter_Cir_Qty.value		= """ & UNINumClientFormat(rs1("MOTER_CIR_QTY"), ggQty.DecPoint, 0) & """" & vbCr
			Response.Write "parent.frm1.txtPm_dt.Text				= """ & UNIDateClientFormat(rs1("PM_DT")) & """" & vbCr
			Response.Write "parent.frm1.txtOil_Spec4.Value			= """ & ConvSPChars(rs1("OIL_SPEC4")) & """" & vbCr
			Response.Write "parent.frm1.txtMoter_Bearing.Value		= """ & ConvSPChars(rs1("MOTER_BEARING")) & """" & vbCr
			Response.Write "parent.frm1.txtPm_Reason.Value			= """ & ConvSPChars(rs1("PM_REASON")) & """" & vbCr
			Response.Write "parent.frm1.txtOil_Spec5.Value			= """ & ConvSPChars(rs1("OIL_SPEC5")) & """" & vbCr
			Response.Write "parent.frm1.txtDocCur.Value				= """ & ConvSPChars(rs1("CURRENCY")) & """" & vbCr

    		Response.Write "parent.DbDtlQueryOk "	& vbCr

	        Response.Write "</Script>"		& vbCr
	        Response.end
	End If


	If  rs0.EOF And rs0.BOF  Then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "parent.frm1.txthFacility_CD.value = """ & "" & """" & vbCr
		Response.Write "parent.frm1.txthFacility_Nm.value = """ & "" & """" & vbCr
		Response.Write "</Script>"		& vbCr
	Else

		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "parent.frm1.txthFacility_Cd.value = """ & ConvSPChars(rs0("Facility_Cd")) & """" & vbCr
		Response.Write "parent.frm1.txthFacility_Nm.value = """ & ConvSPChars(rs0("Facility_Nm")) & """" & vbCr
		Response.Write "</Script>"		& vbCr
	End If

	rs0.Close
	Set rs0 = Nothing


	If  rs1.EOF And rs1.BOF  Then
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write "On error Resume Next" & vbCr
		Response.Write "parent.Frm1.CbohFacility_Accnt.Focus" & vbCr
		Response.Write "On error Goto 0" & vbCr
		Response.Write "</Script>"		& vbCr
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		rs1.Close
		Set rs0 = Nothing
		Set rs1 = Nothing
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.Write " parent.FncNew "	& vbCr
		Response.Write "</Script>"		& vbCr
		Response.end
	Else

'         Call  MakeHeaderData()
		Call  MakeSpreadSheetData()
	End If

'     Call DisplayMsgBox("x", vbInformation, "이상하넹", "FASDFADS1111", I_MKSCRIPT)
End Sub


'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Dim iLoopCount
	Dim iRowStr
	Dim ColCnt

	lgDataExist    = "Yes"
	If CLng(lgPageNo) > 0 Then
	   rs1.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	End If

	iLoopCount = 0
	Do while Not (rs1.EOF Or rs1.BOF)
		iLoopCount =  iLoopCount + 1
		iRowStr = ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1("Facility_Accnt"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1("Facility_Accnt_Nm"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1("Facility_Cd"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1("Facility_Nm"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1("Set_Plant"))
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1("Plant_Nm"))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1("Prod_Co"))
		iRowStr = iRowStr & Chr(11) & UNIDateClientFormat(rs1("Pur_Dt"))
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs1("Prod_Amt"),ggExchRate.DecPoint,0)	'16
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs1("Life_Cycle"),ggExchRate.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & UNINumClientFormat(rs1("Chk_Prd1"),ggExchRate.DecPoint,0)
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs1("Pic_Flag"))
		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount

		If iLoopCount - 1 < lgMaxCount Then
			istrData = istrData & iRowStr & Chr(11) & Chr(12)
		Else
			lgPageNo = lgPageNo + 1
			Exit Do
		End If
		rs1.MoveNext
	Loop



	If iLoopCount <= lgMaxCount Then                                      '☜: Check if next data exists
		lgPageNo = ""
	End If
	GroupCount = iLoopCount
	rs1.Close                                                       '☜: Close recordset object
	Set rs1 = Nothing	                                            '☜: Release ADF
End Sub


'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next
	Err.Clear
	Dim pPY5G110																	'☆ : 입력/수정용 Component Dll 사용 변수 
	Dim Y5_Y_Facility, iCommandSent
	Dim iIntFlgMode

	iStrFacilityCd = Trim(Request("txtFacility_Cd"))


	Const PY51_Y5_FACILITY_CD		= 0
	Const PY51_Y5_FACILITY_NM		= 1
	Const PY51_Y5_FACILITY_ACCNT	= 2
	Const PY51_Y5_FACILITY_LVL1		= 3
	Const PY51_Y5_FACILITY_LVL2		= 4
	Const PY51_Y5_ASST_CD1			= 5
	Const PY51_Y5_ASST_CD2			= 6
	Const PY51_Y5_SET_PLANT			= 7
	Const PY51_Y5_SET_PLACE			= 8
	Const PY51_Y5_SET_DT			= 9
	Const PY51_Y5_SET_CO			= 10
	Const PY51_Y5_PUR_CO			= 11
	Const PY51_Y5_PUR_DT			= 12
	Const PY51_Y5_PROD_CO			= 13
	Const PY51_Y5_PROD_NO			= 14
	Const PY51_Y5_PROD_FLAG			= 15
	Const PY51_Y5_PROD_AMT			= 16
	Const PY51_Y5_MODEL_STS			= 17
	Const PY51_Y5_USE_VOLT			= 18
	Const PY51_Y5_USE_AMOUNT		= 19
	Const PY51_Y5_EQUIP_AREA		= 20
	Const PY51_Y5_CHK_PRD1			= 21
	Const PY51_Y5_CHK_PRD2			= 22
	Const PY51_Y5_CHK_END_DT		= 23
	Const PY51_Y5_REP_END_DT		= 24
	Const PY51_Y5_JNG_END_DT		= 25
	Const PY51_Y5_PM_DT				= 26
	Const PY51_Y5_PM_REASON			= 27
	Const PY51_Y5_LIFE_CYCLE		= 28
	Const PY51_Y5_OIL_SPEC1			= 29
	Const PY51_Y5_OIL_SPEC2			= 30
	Const PY51_Y5_OIL_SPEC3			= 31
	Const PY51_Y5_OIL_SPEC4			= 32
	Const PY51_Y5_OIL_SPEC5			= 33
	Const PY51_Y5_PLANT_STS			= 34
	Const PY51_Y5_MOTER_TYPE		= 35
	Const PY51_Y5_MOTER_QTY			= 36
	Const PY51_Y5_MOTER_POWER		= 37
	Const PY51_Y5_MOTER_CIR_QTY		= 38
	Const PY51_Y5_MOTER_BEARING		= 39
	Const PY51_Y5_PRESS_POWER		= 40
	Const PY51_Y5_EMP_NO			= 41
	Const PY51_Y5_PIC_FLAG			= 42
	Const PY51_Y5_USE_YN			= 43
	Const PY51_Y5_CURRENCY			= 44

	iIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	'-----------------------
	'Data manipulate area
	'-----------------------
	Redim Y5_Y_Facility(PY51_Y5_CURRENCY)


	Y5_Y_Facility(PY51_Y5_FACILITY_CD	)		= (Trim(Request("txtFacility_Cd")))
	Y5_Y_Facility(PY51_Y5_FACILITY_NM	)		= (Trim(Request("txtFacility_Nm")))
	Y5_Y_Facility(PY51_Y5_FACILITY_ACCNT)		= (Trim(Request("CboFacility_Accnt")))
	Y5_Y_Facility(PY51_Y5_FACILITY_LVL1	)		= (Trim(Request("txtItemGroupCd1")))
	Y5_Y_Facility(PY51_Y5_FACILITY_LVL2	)		= (Trim(Request("txtItemGroupCd2")))
	Y5_Y_Facility(PY51_Y5_ASST_CD1		)		= (Trim(Request("txtCondAsstNo1")))
	Y5_Y_Facility(PY51_Y5_ASST_CD2		)		= (Trim(Request("txtCondAsstNo2")))
	Y5_Y_Facility(PY51_Y5_SET_PLANT		)		= (Trim(Request("txtPlantCd")))
	Y5_Y_Facility(PY51_Y5_SET_PLACE		)		= (Trim(Request("txtSet_Place")))
	Y5_Y_Facility(PY51_Y5_SET_CO		)		= (Trim(Request("txtSetCoCd")))
	Y5_Y_Facility(PY51_Y5_PUR_CO		)		= (Trim(Request("txtPurCoCd")))
	Y5_Y_Facility(PY51_Y5_PROD_CO		)		= (Trim(Request("txtProdCoCd")))
	Y5_Y_Facility(PY51_Y5_PROD_NO		)		= (Trim(Request("txtProdNo")))
	Y5_Y_Facility(PY51_Y5_PROD_FLAG		)		= (Trim(Request("txtProd_Flag")))
	Y5_Y_Facility(PY51_Y5_PROD_AMT		)		= UniConvNum(Request("txtProd_Amt"), 0)
	Y5_Y_Facility(PY51_Y5_MODEL_STS		)		= (Trim(Request("txtModel_Sts")))
	Y5_Y_Facility(PY51_Y5_USE_VOLT		)		= (Trim(Request("txtUseVolt")))
	Y5_Y_Facility(PY51_Y5_USE_AMOUNT	)		= (Trim(Request("txtUse_Amount")))
	Y5_Y_Facility(PY51_Y5_EQUIP_AREA	)		= (Trim(Request("txtEquip_Area")))
	Y5_Y_Facility(PY51_Y5_CHK_PRD1	    )		= UniConvNum(Request("txtChk_Prd1"), 0)
	Y5_Y_Facility(PY51_Y5_CHK_PRD2	    )		= UniConvNum(Request("txtChk_Prd2"), 0)
	Y5_Y_Facility(PY51_Y5_PM_REASON	    )		= (Trim(Request("txtPm_Reason")))
	Y5_Y_Facility(PY51_Y5_LIFE_CYCLE	)		= UniConvNum(Request("txtLife_Cycle"), 0)
	Y5_Y_Facility(PY51_Y5_OIL_SPEC1	    )		= (Trim(Request("txtOil_Spec1")))
	Y5_Y_Facility(PY51_Y5_OIL_SPEC2	    )		= (Trim(Request("txtOil_Spec2")))
	Y5_Y_Facility(PY51_Y5_OIL_SPEC3	    )		= (Trim(Request("txtOil_Spec3")))
	Y5_Y_Facility(PY51_Y5_OIL_SPEC4	    )		= (Trim(Request("txtOil_Spec4")))
	Y5_Y_Facility(PY51_Y5_OIL_SPEC5	    )		= (Trim(Request("txtOil_Spec5")))
	Y5_Y_Facility(PY51_Y5_PLANT_STS	    )		= (Trim(Request("txtPlant_Sts")))
	Y5_Y_Facility(PY51_Y5_MOTER_TYPE	)		= (Trim(Request("txtMoter_Type")))
	Y5_Y_Facility(PY51_Y5_MOTER_QTY	    )		= UniConvNum(Request("txtMoter_qty"), 0)
	Y5_Y_Facility(PY51_Y5_MOTER_POWER	)		= UniConvNum(Request("txtMoter_Power"), 0)
	Y5_Y_Facility(PY51_Y5_MOTER_CIR_QTY )		= UniConvNum(Request("txtMoter_Cir_Qty"), 0)
	Y5_Y_Facility(PY51_Y5_MOTER_BEARING )		= (Trim(Request("txtMoter_Bearing")))
	Y5_Y_Facility(PY51_Y5_PRESS_POWER	)		= UniConvNum(Request("txtPress_Power"), 0)
	Y5_Y_Facility(PY51_Y5_EMP_NO		)		= (Trim(Request("txtEmp_no")))
	Y5_Y_Facility(PY51_Y5_PIC_FLAG	    )		= "N"
	Y5_Y_Facility(PY51_Y5_USE_YN		)		= (Trim(Request("CboUse_Yn")))
	Y5_Y_Facility(PY51_Y5_CURRENCY		)		= (Trim(Request("txtDocCur")))


	If Len(Trim(Request("txtSetDt"))) Then
		If UniConvDate(Request("txtSetDt")) = "" Then
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtSetDt", 0, I_MKSCRIPT)
			Response.End
		Else
			Y5_Y_Facility(PY51_Y5_SET_DT		)		= UniConvDate(Request("txtSetDt"))
		End If
	End If

	If Len(Trim(Request("txtPurDt"))) Then
		If UniConvDate(Request("txtPurDt")) = "" Then
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtPurDt", 0, I_MKSCRIPT)
			Response.End
		Else
			Y5_Y_Facility(PY51_Y5_PUR_DT		)		= UniConvDate(Request("txtPurDt"))
		End If
	End If

	If Len(Trim(Request("txtChk_End_dt"))) Then
		If UniConvDate(Request("txtChk_End_dt")) = "" Then
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtChk_End_dt", 0, I_MKSCRIPT)
			Response.End
		Else
			Y5_Y_Facility(PY51_Y5_CHK_END_DT	)		= UniConvDate(Request("txtChk_End_dt"))
		End If
	End If

	If Len(Trim(Request("txtRep_End_dt"))) Then
		If UniConvDate(Request("txtRep_End_dt")) = "" Then
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtRep_End_dt", 0, I_MKSCRIPT)
			Response.End
		Else
			Y5_Y_Facility(PY51_Y5_REP_END_DT	)		= UniConvDate(Request("txtRep_End_dt"))
		End If
	End If

	If Len(Trim(Request("txtJng_End_dt"))) Then
		If UniConvDate(Request("txtJng_End_dt")) = "" Then
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtJng_End_dt", 0, I_MKSCRIPT)
			Response.End
		Else
			Y5_Y_Facility(PY51_Y5_JNG_END_DT	)		= UniConvDate(Request("txtJng_End_dt"))
		End If
	End If

	If Len(Trim(Request("txtPm_dt"))) Then
		If UniConvDate(Request("txtPm_dt")) = "" Then
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtPm_dt", 0, I_MKSCRIPT)
			Response.End
		Else
			Y5_Y_Facility(PY51_Y5_PM_DT		    )		= UniConvDate(Request("txtPm_dt"))
		End If
	End If




	If iIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf iIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	Set pPY5G110 = Server.CreateObject("PY5G110.cBMngFacility")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If






	Call pPY5G110.B_MANAGE_FACILITY(gStrGlobalCollection, iCommandSent, Y5_Y_Facility)

	Select Case Trim(Cstr(Err.Description))
		Case "B_MESSAGE" & Chr(11) & "970023"

		Case Else
			If CheckSYSTEMError(Err, True) = True Then
				Set pPY5G110 = Nothing															'☜: Unload Component
				Response.End
			End If
	End Select

	Set pPY5G110 = Nothing															'☜: Unload Component

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write " parent.frm1.hFacility_CD.value = """ & iStrFacilityCd & """" & vbCr
	Response.Write "       Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "</Script>"		& vbCr



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
End Sub
'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,arrColVal)
	Dim iSelCount

	On Error Resume Next

	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
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


End Sub
'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(Byval lRow, Byval lCol, Byval iLoc)

	If Trim(lRow) = "" Then Exit Function
	If iLoc = I_INSCRIPT Then
		strHTML = "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		Response.Write strHTML
	ElseIf iLoc = I_MKSCRIPT Then
		strHTML = "<" & "Script LANGUAGE=VBScript" & ">" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.focus" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Row = " & lRow & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Col = " & lCol & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.Action = 0" & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelStart = 0 " & vbCrLf
		strHTML = strHTML & "parent.frm1.vspdData.SelLength = len(parent.frm1.vspdData.Text) " & vbCrLf
		strHTML = strHTML & "</" & "Script" & ">" & vbCrLf
		Response.Write strHTML
	End If
End Function

%>
