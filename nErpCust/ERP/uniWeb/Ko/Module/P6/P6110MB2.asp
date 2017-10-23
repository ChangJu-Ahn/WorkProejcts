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
Call LoadInfTB19029B("Q", "P", "NOCOOKIE","MB")

Dim lgOpModeCRUD

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData
Dim iStrCastCd
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
	Dim pPY6G110																	'☆ : 입력/수정용 Component Dll 사용 변수 
	Dim Y6_Y_Cast, iCommandSent
	Dim iIntFlgMode

	iStrCastCd = Trim(Request("txtCastCd1"))
    
	Const PY61_Y6_CAST_CD		= 0


	iIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	'-----------------------
	'Data manipulate area
	'-----------------------
	Redim Y6_Y_Cast(PY61_Y6_Cast_CD)


	Y6_Y_Cast(PY61_Y6_Cast_CD	)		= UCase(Trim(Request("txtCastCd1")))

	iCommandSent = "DELETE"

	Set pPY6G110 = Server.CreateObject("PY6G110.cBMngCast")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

	Call pPY6G110.B_MANAGE_Cast(gStrGlobalCollection, iCommandSent, Y6_Y_Cast)

	Select Case Trim(Cstr(Err.Description))
		Case "B_MESSAGE" & Chr(11) & "970023"

		Case Else
			If CheckSYSTEMError(Err, True) = True Then
				Set pPY6G110 = Nothing															'☜: Unload Component
				Response.End
			End If
	End Select

	Set pPY6G110 = Nothing															'☜: Unload Component

	Response.Write "<Script Language=vbscript>" & vbCr
' 	Response.Write "parent.frm1.hCast_CD.value = """ & iStrCastCd & """" & vbCr
	Response.Write "       Parent.DbDeleteOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "</Script>"		& vbCr

End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()

	On Error Resume Next
	Err.Clear
	Dim pPY6G110																	'☆ : 입력/수정용 Component Dll 사용 변수 
	Dim Y6_Y_Cast, iCommandSent
	Dim iIntFlgMode
	
	iStrCastCd = Trim(Request("txtCastCd"))


    Const PY61_Y6_CAST_CD = 0
    Const PY61_Y6_CAST_NM = 1
    Const PY61_Y6_CAR_KIND = 2
    Const PY61_Y6_MFG_CD = 3
    Const PY61_Y6_ASST_CD1 = 4
    Const PY61_Y6_ASST_CD2 = 5
    Const PY61_Y6_CUSTOM_YN = 6
    Const PY61_Y6_MAKER = 7
    Const PY61_Y6_MAKE_DT = 8
    Const PY61_Y6_STR_TYPE = 9
    Const PY61_Y6_MAT_Q = 10
    Const PY61_Y6_PROCESS_TYPE = 11
    Const PY61_Y6_SPEC = 12
    Const PY61_Y6_WEIGHT_T = 13
    Const PY61_Y6_S_HEIGHT = 14
    Const PY61_Y6_D_HEIGHT = 15
    Const PY61_Y6_FORMING_P = 16
    Const PY61_Y6_CUSHION_PR = 17
    Const PY61_Y6_C_STROKE = 18
    Const PY61_Y6_PUR_AMT = 19
    Const PY61_Y6_LIFE_CYCLE = 20
    Const PY61_Y6_CLOSE_DT = 21
    Const PY61_Y6_USE_MACHINE = 22
    Const PY61_Y6_AUTO_MATH = 23
    Const PY61_Y6_PERSON_COUNT = 24
    Const PY61_Y6_MODIFY_DIRE = 25
    Const PY61_Y6_GUIDE_MATH = 26
    Const PY61_Y6_LOCATE = 27
    Const PY61_Y6_LOADING = 28
    Const PY61_Y6_UNLOADING = 29
    Const PY61_Y6_SCRAP_PROCESS = 30
    Const PY61_Y6_CUSTODY_AREA = 31
    Const PY61_Y6_CHECK_END_DT = 32
    Const PY61_Y6_REP_END_DT = 33
    Const PY61_Y6_PIC_FLAG = 34
    Const PY61_Y6_INSP_PRID = 35
    Const PY61_Y6_CUR_ACCNT = 36
    Const PY61_Y6_FIN_CUR_ACCNT = 37
    Const PY61_Y6_FIN_AJ_DT = 38
    Const PY61_Y6_ITEM_CD_1 = 39
    Const PY61_Y6_ITEM_CD_2 = 40
    Const PY61_Y6_ITEM_CD_3 = 41
    Const PY61_Y6_ITEM_CD_4 = 42
    Const PY61_Y6_ITEM_CD_5 = 43
    Const PY61_Y6_ITEM_CD_6 = 44
    Const PY61_Y6_ITEM_CD_7 = 45
    Const PY61_Y6_ITEM_CD_8 = 46
    Const PY61_Y6_ITEM_CD_9 = 47
    Const PY61_Y6_ITEM_CD_10 = 48
    Const PY61_Y6_LR_FLAG = 49
    Const PY61_Y6_PRS_UNIT = 50
    Const PY61_Y6_SET_PLANT = 51
    Const PY61_Y6_PRS_STS = 52
    Const PY61_Y6_EMP_CD = 53
    Const PY61_Y6_USE_YN = 54
    Const PY61_Y6_SET_PLACE = 55
    Const PY61_Y6_OPR_NO = 56
    Const PY61_Y6_PUR_CUR_CD = 57
    Const PY61_Y6_LIMIT_ACCNT = 58

	iIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 
	'-----------------------
	'Data manipulate area
	'-----------------------
	Redim Y6_Y_Cast(PY61_Y6_LIMIT_ACCNT)

	Y6_Y_Cast(PY61_Y6_CAST_CD)        = UCase(Trim(Request("txtCastCd1")))     
	Y6_Y_Cast(PY61_Y6_CAST_NM)        = (Trim(Request("txtCastNm1")))     
	Y6_Y_Cast(PY61_Y6_CAR_KIND)       = UCase(Trim(Request("txtCarKind1")))  
	'Y6_Y_Cast(PY61_Y6_MFG_CD)         = UCase(Trim(Request("txtMfgCd")))
	Y6_Y_Cast(PY61_Y6_ASST_CD1)       = UCase(Trim(Request("txtAsstCd1"))) 
	Y6_Y_Cast(PY61_Y6_ASST_CD2)       = UCase(Trim(Request("txtAsstCd2"))) 
	Y6_Y_Cast(PY61_Y6_CUSTOM_YN)      = UCase(Trim(Request("CustomYn")))     
	Y6_Y_Cast(PY61_Y6_MAKER)          = (Trim(Request("txtMaker")))   
	Y6_Y_Cast(PY61_Y6_MAKE_DT)        = UCase(Trim(Request("txtMakeDt")))     
	Y6_Y_Cast(PY61_Y6_STR_TYPE)       = Trim(Request("txtStrType"))
	Y6_Y_Cast(PY61_Y6_MAT_Q)          = Trim(Request("txtMatQ"))   
	Y6_Y_Cast(PY61_Y6_PROCESS_TYPE)   = Trim(Request("txtProcessType"))
	Y6_Y_Cast(PY61_Y6_SPEC)           = Trim(Request("txtSpec"))
	Y6_Y_Cast(PY61_Y6_WEIGHT_T)       = UniConvNum(Request("txtWeightT"), 0)  
	Y6_Y_Cast(PY61_Y6_S_HEIGHT)       = UniConvNum(Request("txtSHeight"), 0)   
	Y6_Y_Cast(PY61_Y6_D_HEIGHT)       = UniConvNum(Request("txtDHeight"), 0)     
	Y6_Y_Cast(PY61_Y6_FORMING_P)      = Trim(Request("txtFormingP")) 
	Y6_Y_Cast(PY61_Y6_CUSHION_PR)     = UniConvNum(Request("txtCushionPr"), 0) 
	Y6_Y_Cast(PY61_Y6_C_STROKE)       = UniConvNum(Request("txtCStroke"), 0) 
	Y6_Y_Cast(PY61_Y6_PUR_AMT)        = UniConvNum(Request("txtPurAmt"), 0) 
	Y6_Y_Cast(PY61_Y6_LIFE_CYCLE)     = UniConvNum(Request("txtLifeCycle"), 0)  
	Y6_Y_Cast(PY61_Y6_CLOSE_DT)       = UniConvNum(Request("txtCloseDt"), 0)  
	Y6_Y_Cast(PY61_Y6_USE_MACHINE)    = Trim(Request("txtUseMachine"))     
	Y6_Y_Cast(PY61_Y6_AUTO_MATH)      = Trim(Request("txtAutoMath"))     
	Y6_Y_Cast(PY61_Y6_PERSON_COUNT)   = UniConvNum(Request("txtPersonCount"), 0)     
	Y6_Y_Cast(PY61_Y6_MODIFY_DIRE)    = Trim(Request("txtModifyDire"))    
	Y6_Y_Cast(PY61_Y6_GUIDE_MATH)     = Trim(Request("txtGuideMath"))    
	Y6_Y_Cast(PY61_Y6_LOCATE)         = Trim(Request("txtLocate"))   
	Y6_Y_Cast(PY61_Y6_LOADING)        = Trim(Request("txtLoading"))    
	Y6_Y_Cast(PY61_Y6_UNLOADING)      = Trim(Request("txtUnLoading"))   
	Y6_Y_Cast(PY61_Y6_SCRAP_PROCESS)  = Trim(Request("txtScrapProcess")) 
	Y6_Y_Cast(PY61_Y6_CUSTODY_AREA)   = Trim(Request("txtCustodyArea"))
	Y6_Y_Cast(PY61_Y6_CHECK_END_DT)   = Trim(Request("txtCheckEndDt")) 
	Y6_Y_Cast(PY61_Y6_REP_END_DT)     = Trim(Request("txtRepEndDt")) 
	Y6_Y_Cast(PY61_Y6_PIC_FLAG)       = "N"                  
	Y6_Y_Cast(PY61_Y6_INSP_PRID)      = UniConvNum(Request("txtInspPrid"), 0) 
	Y6_Y_Cast(PY61_Y6_CUR_ACCNT)      = UniConvNum(Request("txtCurAccnt"), 0)  
	Y6_Y_Cast(PY61_Y6_FIN_CUR_ACCNT)  = UniConvNum(Request("txtFinCurAccnt"), 0)     
	Y6_Y_Cast(PY61_Y6_FIN_AJ_DT)      = UCase(Trim(Request("txtFinAjDt")))  
	Y6_Y_Cast(PY61_Y6_ITEM_CD_1)      = UCase(Trim(Request("txtItemCd1")))
	Y6_Y_Cast(PY61_Y6_ITEM_CD_2)      = UCase(Trim(Request("txtItemCd2")))
	Y6_Y_Cast(PY61_Y6_ITEM_CD_3)      = UCase(Trim(Request("txtItemCd3"))) 
	Y6_Y_Cast(PY61_Y6_ITEM_CD_4)      = UCase(Trim(Request("txtItemCd4"))) 
	Y6_Y_Cast(PY61_Y6_ITEM_CD_5)      = UCase(Trim(Request("txtItemCd5"))) 
	Y6_Y_Cast(PY61_Y6_ITEM_CD_6)      = UCase(Trim(Request("txtItemCd6"))) 
	Y6_Y_Cast(PY61_Y6_ITEM_CD_7)      = UCase(Trim(Request("txtItemCd7"))) 
	Y6_Y_Cast(PY61_Y6_ITEM_CD_8)      = UCase(Trim(Request("txtItemCd8"))) 
	Y6_Y_Cast(PY61_Y6_ITEM_CD_9)      = UCase(Trim(Request("txtItemCd9"))) 
	Y6_Y_Cast(PY61_Y6_ITEM_CD_10)     = UCase(Trim(Request("txtItemCd10")))
	Y6_Y_Cast(PY61_Y6_LR_FLAG)        = UCase(Trim(Request("LrFlag")))
	Y6_Y_Cast(PY61_Y6_PRS_UNIT)       = UniConvNum(Request("txtPrsUnit"), 0)  
	Y6_Y_Cast(PY61_Y6_SET_PLANT)	  = UCase(Trim(Request("txtSetPlantCd1")))  
	Y6_Y_Cast(PY61_Y6_PRS_STS)		  = UCase(Trim(Request("cboPrsSts")))    
	Y6_Y_Cast(PY61_Y6_EMP_CD)		  = UCase(Trim(Request("cboEmpCd")))  
	Y6_Y_Cast(PY61_Y6_USE_YN)         = UCase(Trim(Request("cboUseYn")))
	Y6_Y_Cast(PY61_Y6_SET_PLACE)      = UCase(Trim(Request("txtSetPlace")))
	'Y6_Y_Cast(PY61_Y6_OPR_NO)         = UCase(Trim(Request("cboOprNo")))	
	Y6_Y_Cast(PY61_Y6_PUR_CUR_CD)      = UCase(Trim(Request("txtPurCurCd")))
	Y6_Y_Cast(PY61_Y6_LIMIT_ACCNT)      = UniConvNum(Request("txtLimitAccnt"), 0) 
  
	
	If Len(Trim(Request("txtMakeDt"))) Then
		If UniConvDate(Request("txtMakeDt")) = "" Then	 
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtMakeDt", 0, I_MKSCRIPT)
			Response.End	
		Else
			Y6_Y_Cast(PY61_Y6_MAKE_DT		)		= UniConvDate(Request("txtMakeDt"))
		End If
	End If
	
	If Len(Trim(Request("txtCheckEndDt"))) Then
		If UniConvDate(Request("txtCheckEndDt")) = "" Then	 
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtCheckEndDt", 0, I_MKSCRIPT)
			Response.End	
		Else
			Y6_Y_Cast(PY61_Y6_CHECK_END_DT		)		= UniConvDate(Request("txtCheckEndDt"))
		End If
	End If
	
	If Len(Trim(Request("txtRepEndDt"))) Then
		If UniConvDate(Request("txtRepEndDt")) = "" Then	 
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtRepEndDt", 0, I_MKSCRIPT)
			Response.End	
		Else
			Y6_Y_Cast(PY61_Y6_REP_END_DT	)		= UniConvDate(Request("txtRepEndDt"))
		End If
	End If
	
	If Len(Trim(Request("txtCloseDt"))) Then
		If UniConvDate(Request("txtCloseDt")) = "" Then	 
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtCloseDt", 0, I_MKSCRIPT)
			Response.End	
		Else
			Y6_Y_Cast(PY61_Y6_CLOSE_DT	)		= UniConvDate(Request("txtCloseDt"))
		End If
	End If
	
	If Len(Trim(Request("txtFinAjDt"))) Then
		If UniConvDate(Request("txtFinAjDt")) = "" Then	 
			Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
			Call LoadTab("parent.frm1.txtFinAjDt", 0, I_MKSCRIPT)
			Response.End	
		Else
			Y6_Y_Cast(PY61_Y6_FIN_AJ_DT	)		= UniConvDate(Request("txtFinAjDt"))
		End If
	End If
	
	If iIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
	ElseIf iIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
	End If

	Set pPY6G110 = Server.CreateObject("PY6G110.cBMngCast")

	If CheckSYSTEMError(Err,True) = True Then
		Response.End
	End If

	Call pPY6G110.B_MANAGE_Cast(gStrGlobalCollection, iCommandSent, Y6_Y_Cast)

	Select Case Trim(Cstr(Err.Description))
		Case "B_MESSAGE" & Chr(11) & "125000"
			Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(1)
			Exit Sub
		Case "B_MESSAGE" & Chr(11) & "Y60030"
			Call DisplayMsgBox("Y60030", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(2)
			Exit Sub
		Case "B_MESSAGE" & Chr(11) & "182100"
			Call DisplayMsgBox("182100", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(3)
			Exit Sub
		Case "1171001"
			Call DisplayMsgBox("Y60100", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(4)
			Exit Sub
		Case "1171002"
			Call DisplayMsgBox("Y60100", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(5)
			Exit Sub
		Case "1226001"
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(6)
			Exit Sub
		Case "1226002"
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(7)
			Exit Sub
		Case "1226003"
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(8)
			Exit Sub
		Case "1226004"
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(9)
			Exit Sub
		Case "1226005"
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(10)
			Exit Sub
		Case  "1226006"
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(11)
			Exit Sub
		Case  "1226007"
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(12)
			Exit Sub
		Case  "1226008"
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(13)
			Exit Sub
		Case "1226009"
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(14)
			Exit Sub
		Case "12260010"
			Call DisplayMsgBox("122600", vbInformation, "", "", I_MKSCRIPT)
			Call CodeCheck(15)
			Exit Sub
		Case Else
			If CheckSYSTEMError(Err, True) = True Then
				Set pPY6G110 = Nothing															'☜: Unload Component
				Response.End
			End If
	End Select
    
	Set pPY6G110 = Nothing															'☜: Unload Component
	
	Response.Write "<Script Language=vbscript>" & vbCr
	Response.Write "	parent.frm1.txtCastCD.value = """ & Y6_Y_Cast(PY61_Y6_CAST_CD) & """" & vbCr
	Response.Write "	parent.frm1.txtSetPlantCd.value = """ & Y6_Y_Cast(PY61_Y6_SET_PLANT) & """" & vbCr
	Response.Write "	Parent.DbSaveOk "      & vbCr							'☜: 화면 처리 ASP 를 지칭함 
	Response.Write "</Script>"		& vbCr

End Sub

Sub CodeCheck(pCode)

	Response.Write "<Script Language=VBScript>" & vbCrLf
	Response.Write "Call parent.ClickTab2() " & vbCrLf
	Select Case pCode
		Case 1
			Response.Write "parent.frm1.txtSetPlantCd1.focus " & vbCrLf
			Response.Write "parent.frm1.txtSetPlantNm1.value = """ & "" & """" & vbCr
		Case 2
			Response.Write "parent.frm1.txtCarKind1.focus " & vbCrLf
		Case 3
			Response.Write "parent.frm1.txtSetPlace.focus " & vbCrLf
		Case 4
			Response.Write "parent.frm1.txtAsstCd1.focus " & vbCrLf
		Case 5
			Response.Write "parent.frm1.txtAsstCd2.focus " & vbCrLf
		Case 6
			Response.Write "parent.frm1.txtItemCd1.focus " & vbCrLf
		Case 7
			Response.Write "parent.frm1.txtItemCd2.focus " & vbCrLf
		Case 8
			Response.Write "parent.frm1.txtItemCd3.focus " & vbCrLf
		Case 9
			Response.Write "parent.frm1.txtItemCd4.focus " & vbCrLf
		Case 10
			Response.Write "parent.frm1.txtItemCd5.focus " & vbCrLf
		Case 11
			Response.Write "parent.frm1.txtItemCd6.focus " & vbCrLf
		Case 12
			Response.Write "parent.frm1.txtItemCd7.focus " & vbCrLf
		Case 13
			Response.Write "parent.frm1.txtItemCd8.focus " & vbCrLf
		Case 14
			Response.Write "parent.frm1.txtItemCd9.focus " & vbCrLf
		Case 15
			Response.Write "parent.frm1.txtItemCd10.focus " & vbCrLf																	
	End Select

	Response.Write "</Script>" & vbCrLf
	Response.End

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

%>
