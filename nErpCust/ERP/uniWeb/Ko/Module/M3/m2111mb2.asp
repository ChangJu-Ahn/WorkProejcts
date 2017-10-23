<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        :
'*  3. Program ID           : m2111mb2
'*  4. Program Name         : 업체지정 
'*  5. Program Desc         : 업체지정 
'*  6. Component List       : PM2G128.cMListAssignSpplS / PM2G121.cMMaintAssignSpplS / PM3SAAS.cMAutoAssignSppl / PM2G139.cMLookupSpplLtS
'*  7. Modified date(First) : 2003/01/14
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Oh Chang Won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<%

call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB")

    Dim lgOpModeCRUD

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message

	lgOpModeCRUD  = Request("txtMode")
							                                              '☜: Read Operation Mode (CRUD)
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call  SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
		Case "AutoAssign"			'☜: 업체자동지정 요청을 받음 
			 Call SubAutoAssign
		Case "LookSppl"				'☜: 공급처 Change Event
			 Call SubLookSppl
    End Select

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	On Error Resume Next                                                             '☜: Protect system from crashing
	Err.Clear                                                                        '☜: Clear Error status

	Dim iPM2G128																	'☆ : 조회용 ComProxy Dll 사용 변수 
	Dim iMax
	Dim PvArr

	Dim I1_ief_supplied_select_char
	Dim I2_b_item_item_cd
	Dim I3_b_plant_plant_cd
	Dim I4_m_pur_req
	Dim I5_m_pur_req_dlvy_dt
	Dim I8_m_pur_req_pur_plan_dt
	Dim I6_b_pur_org_pur_org
	Dim I7_m_pur_req_pr_no
	Dim E1_b_pur_org
	Dim E2_m_pur_req 'pr_no
	Dim EG1_export_group
	Dim E3_b_plant
	Dim E4_b_item

	Const M054_I4_dlvy_dt = 0
	Const M054_I4_mrp_run_no = 1
	Const M054_I4_so_no = 2
	Const M054_I4_tracking_no = 3
	Const M054_I4_pur_plan_dt = 4
	Redim I4_m_pur_req(M054_I4_pur_plan_dt)

    Const M054_E1_pur_org = 0
    Const M054_E1_pur_org_nm = 1

    Const M054_EG1_E1_b_item_by_plant_pur_org = 0
    Const M054_EG1_E2_b_minor_minor_nm = 1
    Const M054_EG1_E3_b_pur_grp_pur_grp_nm = 2
    Const M054_EG1_E4_b_plant_plant_cd = 3
    Const M054_EG1_E4_b_plant_plant_nm = 4
    Const M054_EG1_E5_b_item_item_cd = 5
    Const M054_EG1_E5_b_item_item_nm = 6
    Const M054_EG1_E6_b_biz_partner_bp_nm = 7
    Const M054_EG1_E7_m_pur_req_pr_no = 8
    Const M054_EG1_E7_m_pur_req_req_qty = 9
    Const M054_EG1_E7_m_pur_req_req_unit = 10
    Const M054_EG1_E7_m_pur_req_dlvy_dt = 11
    Const M054_EG1_E7_m_pur_req_pur_plan_dt = 12
    Const M054_EG1_E7_m_pur_req_pr_type = 13
    Const M054_EG1_E7_m_pur_req_mrp_run_no = 14
    Const M054_EG1_E7_m_pur_req_procure_type = 15
    Const M054_EG1_E7_m_pur_req_pr_sts = 16
    Const M054_EG1_E7_m_pur_req_req_cfm_qty = 17
    Const M054_EG1_E7_m_pur_req_ord_qty = 18
    Const M054_EG1_E7_m_pur_req_rcpt_qty = 19
    Const M054_EG1_E7_m_pur_req_iv_qty = 20
    Const M054_EG1_E7_m_pur_req_req_dt = 21
    Const M054_EG1_E7_m_pur_req_req_dept = 22
    Const M054_EG1_E7_m_pur_req_req_prsn = 23
    Const M054_EG1_E7_m_pur_req_sppl_cd = 24
    Const M054_EG1_E7_m_pur_req_sl_cd = 25
    Const M054_EG1_E7_m_pur_req_tracking_no = 26
    Const M054_EG1_E7_m_pur_req_pur_org = 27
    Const M054_EG1_E7_m_pur_req_pur_grp = 28
    Const M054_EG1_E7_m_pur_req_mrp_ord_no = 29
    Const M054_EG1_E7_m_pur_req_base_req_qty = 30
    Const M054_EG1_E7_m_pur_req_base_req_unit = 31
    Const M054_EG1_E7_m_pur_req_so_no = 32
    Const M054_EG1_E7_m_pur_req_so_seq_no = 33
    Const M054_EG1_E7_m_pur_req_ext1_cd = 34
    Const M054_EG1_E7_m_pur_req_ext1_qty = 35
    Const M054_EG1_E7_m_pur_req_ext1_amt = 36
    Const M054_EG1_E7_m_pur_req_ext1_rt = 37
    Const M054_EG1_E7_m_pur_req_ext2_cd = 38
    Const M054_EG1_E7_m_pur_req_ext2_qty = 39
    Const M054_EG1_E7_m_pur_req_ext2_amt = 40
    Const M054_EG1_E7_m_pur_req_ext2_rt = 41
    Const M054_EG1_E7_m_pur_req_ext3_cd = 42
    Const M054_EG1_E7_m_pur_req_ext3_qty = 43
    Const M054_EG1_E7_m_pur_req_ext3_amt = 44
    Const M054_EG1_E7_m_pur_req_ext3_rt = 45
    Const M054_EG1_E8_b_minor_minor_nm = 46
    Const M054_EG1_E5_b_item_item_spec = 47
    Const M054_EG1_E7_b_pur_ogr_pur_org_nm = 48

    Const M054_E3_plant_cd = 0
    Const M054_E3_plant_nm = 1

	Const M054_E4_item_cd = 0
    Const M054_E4_item_nm = 1


	Dim istrData
	Dim StrNextKey		' 다음 값 
	Dim lgStrPrevKey	' 이전 값 
	Dim iLngMaxRow		' 현재 그리드의 최대Row
	Dim iLngRow
	Dim GroupCount
	Dim index,Count     ' 저장 후 Return 해줄 값을 넣을때 쓴는 변수 

	Dim intARows
	Dim intTRows
	intARows=0
	intTRows=0

	Const C_SHEETMAXROWS_D  = 100

    If Len(Trim(Request("txtFrDt"))) Then
		If UNIConvDate(Request("txtFrDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtFrDt", 0, I_MKSCRIPT)
		    Exit Sub
		End If
	End If

	If Len(Trim(Request("txtToDt"))) Then
		If UNIConvDate(Request("txtToDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtToDt", 0, I_MKSCRIPT)
		    Exit Sub
		End If
	End If

	lgStrPrevKey = Trim(Request("lgStrPrevKey"))

    Set iPM2G128 = Server.CreateObject("PM2G128.cMListAssignSpplS")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If CheckSYSTEMError(Err,True) = true Then
		Set iPM2G128 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 

	End if

    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    I6_b_pur_org_pur_org 					= Trim(Request("txtOrgCd"))
    I3_b_plant_plant_cd 					= Trim(Request("txtPlantCd"))
    I2_b_item_item_cd   					= Trim(Request("txtItemCd"))
    If Request("txtFrDt") = "" Then
    	I4_m_pur_req(M054_I4_dlvy_dt) 		= "1900-01-01"
    Else
    	I4_m_pur_req(M054_I4_dlvy_dt) 		= CStr(UNIConvDate(Request("txtFrDt")))
    End if
    If Request("txtToDt") = "" Then
    	I5_m_pur_req_dlvy_dt 				= "2999-12-31"
    Else
    	I5_m_pur_req_dlvy_dt 				= UNIConvDate(Request("txtToDt"))
    End if


    If Request("txtPoFrDt") = "" Then
    	I4_m_pur_req(M054_I4_pur_plan_dt) 		= "1900-01-01"
    Else
    	I4_m_pur_req(M054_I4_pur_plan_dt) 		= CStr(UNIConvDate(Request("txtPoFrDt")))
    End if
    If Request("txtPoToDt") = "" Then
    	I8_m_pur_req_pur_plan_dt 				= "2999-12-31"
    Else
    	I8_m_pur_req_pur_plan_dt 				= UNIConvDate(Request("txtPoToDt"))
    End if

    I4_m_pur_req(M054_I4_mrp_run_no) 		= Trim(Request("txtMRP"))

    I4_m_pur_req(M054_I4_so_no) 			= Trim(Request("txtSoNo"))
    I4_m_pur_req(M054_I4_tracking_no)		= Trim(Request("txtTrkNo"))

    I1_ief_supplied_select_char 			= Request("rdoAppflg")
    I7_m_pur_req_pr_no			 			= Request("lgStrPrevKey")

    '-----------------------
    'Com action area
    '-----------------------***********************************

      Call iPM2G128.M_LIST_ASSIGN_SPPL_SVR(gStrGlobalCollection, C_SHEETMAXROWS_D, _
												I1_ief_supplied_select_char, I2_b_item_item_cd, _
												I3_b_plant_plant_cd, I4_m_pur_req, _
												CStr(I5_m_pur_req_dlvy_dt), I6_b_pur_org_pur_org, _
												I7_m_pur_req_pr_no,I8_m_pur_req_pur_plan_dt ,E1_b_pur_org, E2_m_pur_req, _
												EG1_export_group, E3_b_plant, E4_b_item)


	If CheckSYSTEMError(Err,True) = True Then
		Set iPM2G128 = Nothing
		Response.Write "<Script Language=vbscript>" & vbCr
		Response.write "parent.frm1.vspdData.MaxRows = 0" & vbCr
		Response.Write "parent.DbQueryOk " & intARows & ",iMaxRow"   & vbCr
		Response.Write "</Script>"
		Exit Sub
	End If

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "with parent" & vbCr
	Response.Write "	.frm1.txtPlantNm.value = """ & ConvSPChars(E3_b_plant(M054_E3_plant_nm))      & """" & vbCr
	Response.Write "	.frm1.txtItemNm.value  = """ & ConvSPChars(E4_b_item(M054_E4_item_nm))      & """" & vbCr
	Response.Write "End With "   & vbCr
    Response.Write "</Script>"                  & vbCr


	If lgStrPrevKey = StrNextKey And UBound(EG1_export_group,1) < 0 Then
		Set iPM2G128 = Nothing												'☜: ComProxy Unload
		Exit Sub														'☜: 비지니스 로직 처리를 종료함 
	End If

	iLngMaxRow = CLng(Request("txtMaxRows"))											'Save previous Maxrow
    GroupCount = UBound(EG1_export_group,1)

	If EG1_export_group(GroupCount, M054_EG1_E7_m_pur_req_pr_no) = E2_m_pur_req(0) Then
		StrNextKey = ""
	Else
		StrNextKey = E2_m_pur_req(0)
	End If
	'-----------------------
	'Result data display area
	'-----------------------
	Const strDefDate = "1900-01-01"
	iMax = UBound(EG1_export_group,1)
	ReDim PvArr(iMax)

	For iLngRow = 0 To UBound(EG1_export_group,1)
		If  iLngRow < C_SHEETMAXROWS_D  Then
		Else
		   StrNextKey = E2_m_pur_req(0)
           Exit For
        End If
		istrData = istrData & Chr(11) & "0"                                                                     '1
		istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E7_m_pur_req_pr_no))     '2
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E4_b_plant_plant_cd))    '3
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E4_b_plant_plant_nm))    '4
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E5_b_item_item_cd))		'5
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E5_b_item_item_nm))		'6
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E5_b_item_item_spec))	'7
        istrData = istrData & Chr(11) & UNINumClientFormat(EG1_export_group(iLngRow, M054_EG1_E7_m_pur_req_req_qty),ggQty.DecPoint,0) '8
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E7_m_pur_req_req_unit))	'9
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E7_m_pur_req_tracking_no))	'26
        istrData = istrData & Chr(11) & UNIDateClientFormat(EG1_export_group(iLngRow, M054_EG1_E7_m_pur_req_dlvy_dt)) '10
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E7_m_pur_req_pr_sts))    '11
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E8_b_minor_minor_nm))    '12
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E7_m_pur_req_pr_type))   '13
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E2_b_minor_minor_nm))    '14

        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E1_b_item_by_plant_pur_org)) '15
        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E7_b_pur_ogr_pur_org_nm))  '16

        istrData = istrData & Chr(11) & ConvSPChars(EG1_export_group(iLngRow, M054_EG1_E7_m_pur_req_mrp_run_no)) '17
        istrData = istrData & Chr(11) & iLngMaxRow + iLngRow + 1
        istrData = istrData & Chr(11) & Chr(12)

		PvArr(iLngRow) = istrData
		istrData=""
    Next
    istrData = Join(PvArr, "")


	intARows=iLngRow+1

	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "With parent" & vbCr												'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "Dim iMaxRow " & vbCr
	Response.Write " iMaxRow = .frm1.vspdData.maxrows" & vbCr
    Response.Write "	.ggoSpread.Source    = .frm1.vspdData " & vbCr
    Response.Write "	.ggoSpread.SSShowData  """ & istrData	& """" & vbCr
    Response.Write "	.lgStrPrevKey        = """ & StrNextKey & """" & vbCr
    Response.Write " .frm1.hdnPlant.value    = """ & ConvSPChars(Request("txtPlantCd"))   & """" & vbCr
	Response.Write " .frm1.hdnItem.value     = """ & ConvSPChars(Request("txtItemCd"))    & """" & vbCr
	Response.Write " .frm1.hdnState.value    = """ & ConvSPChars(Request("cboReqStatus")) & """" & vbCr
	Response.Write " .frm1.hdnFrDt.value     = """ & ConvSPChars(Request("txtFrDt"))      & """" & vbCr
	Response.Write " .frm1.hdnToDt.value     = """ & ConvSPChars(Request("txtToDt"))      & """" & vbCr
	Response.Write " .frm1.hdnPoFrDt.value   = """ & ConvSPChars(Request("txtPoFrDt"))      & """" & vbCr
	Response.Write " .frm1.hdnPoToDt.value   = """ & ConvSPChars(Request("txtPoToDt"))      & """" & vbCr
	Response.Write " .frm1.hdnMrp.value      = """ & ConvSPChars(Request("txtMRP"))       & """" & vbCr
	Response.Write " .frm1.hdnflg.value      = """ & ConvSPChars(Request("rdoAppflg"))    & """" & vbCr
	Response.Write " .frm1.hdnOrg.value      = """ & ConvSPChars(Request("txtOrgCd"))     & """" & vbCr
	Response.Write " .frm1.hdnSoNo.value     = """ & ConvSPChars(Request("txtSoNo"))      & """" & vbCr
	Response.Write " .frm1.hdnTrkNo.value    = """ & ConvSPChars(Request("txtTrkNo"))     & """" & vbCr
    Response.Write " .DbQueryOk " & intARows & ",iMaxRow"   & vbCr
    Response.Write " .frm1.vspdData.focus "		   & vbCr
    Response.Write "End With"   & vbCr
    Response.Write "</Script>" & vbCr


    Set iPM2G128 = Nothing											'☜: Unload Comproxy

End Sub
'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()
	On Error Resume Next
    Err.Clear																		'☜: Protect system from crashing

	Dim LngMaxRow
	Dim LngRow
	Dim LngRow1
	Dim iPM2G121
	Dim iErrorPosition
	Dim iErrorMsg
	Dim arrVal, arrTemp,arrTemp1														'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	Dim	lGrpCnt																		'☜: Group Count
    Dim I1_m_pur_req_no
    Dim I2_m_pur_req

    Const M056_pr_no = 0
    Const M056_pur_plan_dt = 1

    Dim Zsep
	Dim itxtSpread
    Dim itxtSpreadArr
    Dim itxtSpreadArrCount

    Dim iCUCount
    Dim ii

    itxtSpread = ""

    iCUCount = Request.Form("txtCUSpread").Count

    itxtSpreadArrCount = -1

    ReDim itxtSpreadArr(iCUCount)

    For ii = 1 To iCUCount
        itxtSpreadArrCount = itxtSpreadArrCount + 1
        itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
    Next

    itxtSpread = Join(itxtSpreadArr,"")

    Zsep = "@"

	arrTemp1 = Split(itxtSpread, Zsep)

    Response.Write "<Script language=vbs> " & vbCr
    Response.Write "Parent.RemovedivTextArea "      & vbCr
    Response.Write "</Script> "      & vbCr

    '1건씩 처리한다 
    For LngRow = 1 To UBound(arrTemp1,1)

		I2_m_pur_req = arrTemp1(LngRow-1)

		arrTemp = Split(arrTemp1(LngRow-1), gRowSep)

		arrVal = Split(arrTemp(0), gColSep)

	    I1_m_pur_req_no = arrVal(7)

	    Set iPM2G121 = Server.CreateObject("PM2G121.cMMaintAssignSpplS")
		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If CheckSYSTEMError(Err,True) = true then
			Set iPM2G121 = Nothing
			Exit Sub
		End If

	    Call iPM2G121.M_MAINT_ASSIGN_SPPL_SVR(gStrGlobalCollection, _
											  I1_m_pur_req_no, _
											  I2_m_pur_req, _
											  iErrorPosition)

		If CheckSYSTEMError2(Err, True, LngRow & "-" & iErrorPosition & "행","","","","") = True Then
			Set iPM2G121 = Nothing
			exit sub
		Else
			'처리가 헤더 완료된것은 Check Box 가 풀림.
			'처리된 Detail은 수정/삭제/입력 플래그가 풀림 200308
			Response.Write "<Script language=vbscript> "		& vbCr
			Response.Write "'On error resume Next"				& vbCr
			Response.Write "	with Parent.frm1.vspdData"      & vbCr
			Response.Write "		Dim iIndex, iRowNo, iStartRowAt2ndGrid	"		& vbCr
			Response.Write "		for iIndex = 1 to .MaxRows	"      & vbCr
			Response.Write "			.Col = Parent.C_ReqNo	"      & vbCr
			Response.Write "			.Row = iIndex	"		& vbCr
			Response.Write "			If Trim(.text) = """	&  I1_m_pur_req_no & """ then "     & vbCr
			Response.Write "				iRowNo = iIndex	"   & vbCr
			Response.Write "			End if	"				& vbCr
			Response.Write "		Next	"					& vbCr
			Response.Write "		.Col = parent.C_CfmFlg	"   & vbCr
			Response.Write "		.Row = iRowNo "				& vbCr
			Response.Write "		.Text = 0 "					& vbCr
			Response.Write "	end with "						& vbCr
			Response.Write "	With Parent.frm1.vspdData2	"	& vbCr
			Response.Write "		Dim lngRow "	& vbCr
			Response.Write "		For lngRow = 1 To .MaxRows"	& vbCr
			Response.Write "			.Row = lngRow"	& vbCr
			Response.Write "			.Col = Parent.C_ParentRowNo"	& vbCr
			Response.Write "			If iRowNo = CInt(.Text) Then		"	& vbCr
			Response.Write "				iStartRowAt2ndGrid = lngRow"	& vbCr
			Response.Write "				Exit For"	& vbCr
			Response.Write "			End If    "	& vbCr
			Response.Write "		Next	"	& vbCr
			Response.Write "		for iIndex = iStartRowAt2ndGrid	to (iStartRowAt2ndGrid - 1) + Parent.lglngHiddenRows(iRowNo - 1)	"	& vbCr
			Response.Write "			.Col = 0					"	& vbCr
			Response.Write "			.Row = iIndex				"	& vbCr
			Response.Write "			.Text = """"				"	& vbCr
			Response.Write "		Next							"	& vbCr
			Response.Write "	End With							"	& vbCr
			Response.Write "								"	& vbCr
			Response.Write "								"	& vbCr
			Response.Write "</Script> "							& vbCr
		End If


    Next

	Set iPM2G121 = Nothing                                                   '☜: Unload Comproxy
    Response.Write "<Script language=vbscript> " & vbCr
    Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "             & vbCr

End Sub
'============================================================================================================
' Name : SubAutoAssign
' Desc : 업체자동지정 
'============================================================================================================
Sub SubAutoAssign()
	On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear																		 '☜: Clear Error status

	Dim iErrorPosition
	Dim LngMaxRow
	Dim arrTemp
	Dim arrVal
	Dim lGrpCnt
	Dim LngRow
	Dim I5_row_cnt
    Dim I3_b_plant_plant_cd
    Dim I4_b_item_item_cd
    Dim I2_m_pur_req_pr_no
    Dim I1_assign_type

	Dim Obj_PM3SAAS

	LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

	arrTemp = Split(Request("txtSpread"), gRowSep)									'☆: Spread Sheet 내용을 담고 있는 Element명 

    lGrpCnt = 0

	Set Obj_PM3SAAS = Server.CreateObject("PM3SAAS.cMAutoAssignSppl")

	If CheckSYSTEMError(Err,True) = true then
		Set Obj_PM3SAAS = Nothing
		Exit Sub
	End If

    For LngRow = 1 To LngMaxRow
		Err.Clear
		lGrpCnt = lGrpCnt +1														'☜: Group Count

		arrVal = Split(arrTemp(LngRow-1), gColSep)

		I1_assign_type		= arrVal(0)
		I2_m_pur_req_pr_no	= arrVal(1)
		I3_b_plant_plant_cd = arrVal(2)
		I4_b_item_item_cd	= arrVal(3)
		I5_row_cnt			= arrVal(4)

		Call Obj_PM3SAAS.M_AUTO_ASSIGN_SPPL(gStrGlobalCollection, _
											I3_b_plant_plant_cd,I4_b_item_item_cd, _
											I2_m_pur_req_pr_no,"","", _
											I1_assign_type)

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If CheckSYSTEMError2(Err, True, I5_row_cnt & "행:", "", "", "", "") = True Then
		    err.Clear
			exit sub
		Else
			'처리가 완료된것은 Check Box 가 풀림.
			Response.Write "<Script language=vbscript> "		& vbCr
			Response.Write "On error resume Next"				& vbCr
			Response.Write "	with Parent.frm1.vspdData"      & vbCr
			Response.Write "		Dim iIndex, iRowNo	"		& vbCr
			Response.Write "		for iIndex = 1 to .MaxRows	"      & vbCr
			Response.Write "			.Col = Parent.C_ReqNo	"      & vbCr
			Response.Write "			.Row = iIndex	"		& vbCr
			Response.Write "			If Trim(.text) = """	&  I2_m_pur_req_pr_no & """ then "     & vbCr
			Response.Write "				iRowNo = iIndex	"   & vbCr
			Response.Write "			End if	"				& vbCr
			Response.Write "		Next	"					& vbCr
			Response.Write "		.Col = parent.C_CfmFlg	"   & vbCr
			Response.Write "		.Row = iRowNo "				& vbCr
			Response.Write "		.Text = 0 "					& vbCr
			Response.Write "	end with "						& vbCr
			Response.Write "</Script> "

		End If

	Next

	If NOT(Obj_PM3SAAS is Nothing) Then
		Set Obj_PM3SAAS = Nothing
	End If

    Response.Write "<Script language=vbs> " & vbCr
    Response.Write " Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
    Response.Write "</Script> "

End Sub
'============================================================================================================
' Name : SubLookSppl
' Desc : 공급처 Change Event
'============================================================================================================
Sub SubLookSppl

   ' On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Dim iPM2G139
    Dim I1_m_pur_req_pr_no
    Dim I2_b_biz_partner_bp_cd
    Dim E1_m_supplier_item_by_plant
    Dim E3_b_pur_grp
    Dim E4_m_sppl_cal
    Dim E5_b_pur_org

    Const M029_E1_sppl_dlvy_lt = 0
    Redim E1_m_supplier_item_by_plant(M029_E1_sppl_dlvy_lt)

    Const M029_E3_pur_grp = 0
    Const M029_E3_pur_grp_nm = 1
    Redim E3_b_pur_grp(M029_E3_pur_grp_nm)

    Const M029_E4_cal_dt = 0
    Redim E4_m_sppl_cal(M029_E4_cal_dt)

    Const M029_E5_pur_org = 0
    Const M029_E5_pur_org_nm = 1
    Redim E5_b_pur_org(M029_E5_pur_org_nm)

    Set iPM2G139 = Server.CreateObject("PM2G139.cMLookupSpplLtS")

    '-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
		If CheckSYSTEMError(Err, True) = True Then
			Set iPM2G139 = Nothing
			Exit Sub
		End If

	I2_b_biz_partner_bp_cd = Trim(Request("txtBpCd"))
	I1_m_pur_req_pr_no = Trim(Request("txtPrNo"))

	Call iPM2G139.M_LOOKUP_SPPL_LT_SVR(gStrGlobalCollection, I1_m_pur_req_pr_no, _
										I2_b_biz_partner_bp_cd, E1_m_supplier_item_by_plant, _
										E3_b_pur_grp, E4_m_sppl_cal)

		If CheckSYSTEMError(Err, True) = True Then
			Set iPM2G139 = Nothing
			Exit Sub
		End If

	'Response.Write "구매그룹==>"&E3_b_pur_grp(M029_E3_pur_grp)
	'Response.Write "구매그룹명==>"&ConvSPChars(E3_b_pur_grp(M029_E5_pur_org_nm))

	Response.Write "<Script language=vbs> " & vbCr
    Response.Write " With Parent.frm1.vspdData2 "   & vbCr
    Response.Write " .Col 	= Parent.C_GrpCd  "     & vbCr
    Response.Write " .Row 	= .ActiveRow		"	    & vbCr
    '공급처변경시 구매그룹 및 발주예정일 재세팅(공급처에 대한 구매그룹 없을때엔 빈칸으로 재세팅) 200308
    'Response.Write "   If .text = """" Then "		& vbCr
    Response.Write "      .text   = """ & ConvSPChars(E3_b_pur_grp(M029_E3_pur_grp)) & """" & vbCr
    Response.Write "      .Col 	= Parent.C_GrpNm "    & vbCr
    Response.Write "      .text   = """ & ConvSPChars(E3_b_pur_grp(M029_E3_pur_grp_nm)) & """" & vbCr
    '발주예정일 Display 200308
    Response.Write "      .Col 	= Parent.C_PlanDt "    & vbCr
    Response.Write "      .text   = """ & UNIDateClientFormat(E4_m_sppl_cal) & """" & vbCr
   'Response.Write "   End If "             & vbCr
    Response.Write " End With "             & vbCr
    Response.Write "</Script> "


Set iPM2G139 = Nothing


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
