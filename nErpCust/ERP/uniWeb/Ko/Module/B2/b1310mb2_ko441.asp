<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->

<%	
							'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf() 

Call HideStatusWnd

On Error Resume Next

Dim pPB2SA05																	'☆ : 조회용 ComProxy Dll 사용 변수 

Dim IntRows
Dim IntCols
Dim vbIntRet
Dim lEndRow
Dim boolCheck
Dim lgIntFlgMode
Dim LngMaxRow

' Com+ Conv. 변수 선언 
    
Dim import_next_b_bank_acct(0) 
Dim import_b_bank(0)

Dim importArray
Dim importArray2
Dim strGroup
    
Dim pvCommandSent
Dim arrCount

Dim strZipcodechk
' 첨자 선언 
Const C_import_b_bank_bank_cd = 0
Const C_import_b_bank_bank_nm = 1
Const C_import_b_bank_bank_full_nm = 2
Const C_import_b_bank_bank_eng_nm = 3
Const C_import_b_bank_zip_cd = 4
Const C_import_b_bank_addr1 = 5
Const C_import_b_bank_addr2 = 6
Const C_import_b_bank_addr3 = 7
Const C_import_b_bank_eng_addr1 = 8
Const C_import_b_bank_eng_addr2 = 9
Const C_import_b_bank_eng_addr3 = 10
Const C_import_b_bank_bank_type = 11
Const C_import_b_bank_country_cd = 12
Const C_import_b_bank_par_bank_cd = 13
Const C_import_b_bank_addr4 = 14
Const C_import_b_bank_bank_fg = 15


LngMaxRow = CInt(Request("txtMaxRows"))											'☜: 최대 업데이트된 갯수 

lgIntFlgMode = CInt(Request("txtFlgMode"))										'☜: 저장시 Create/Update 판별 

Set pPB2SA05 = Server.CreateObject("PB2SA05_KO441.cBMngBankSvr")	    	    

'-----------------------
'Com action result check area(OS,internal)
'-----------------------
If Err.Number <> 0 Then
	Set pPB2SA05 = Nothing		
												'☜: ComProxy Unload
	Call ServerMesgBox(Err.description, vbInformation, I_MKSCRIPT)						'⊙:	
	Response.End															'☜: 비지니스 로직 처리를 종료함 
End If

'-----------------------
'Data manipulate area
'-----------------------	
	ReDim importArray(C_import_b_bank_bank_fg)		'⊙: Single 데이타 저장 
    
    importArray(C_import_b_bank_bank_cd)			= UCase(Trim(Request("txtBankCd1")))
    importArray(C_import_b_bank_bank_nm)			= Trim(Request("txtBankShNm"))
    importArray(C_import_b_bank_bank_full_nm)		= Trim(Request("txtBankFullNm"))
	importArray(C_import_b_bank_bank_eng_nm)		= Trim(Request("txtBankEngNm"))
	importArray(C_import_b_bank_bank_type)			= UCase(Request("cboBankType"))
	importArray(C_import_b_bank_country_cd)			= Trim(Request("txtCountryCd"))
	importArray(C_import_b_bank_zip_cd)				= UCase(Trim(Request("txtZipCd")))
	importArray(C_import_b_bank_addr1)				= Trim(Request("txtAddr1"))
	importArray(C_import_b_bank_addr2)				= Trim(Request("txtAddr2"))
	importArray(C_import_b_bank_addr3)				= Trim(Request("txtAddr3"))
	importArray(C_import_b_bank_eng_addr1)			= Trim(Request("txtEngAddr1"))
	importArray(C_import_b_bank_eng_addr2)			= Trim(Request("txtEngAddr2"))
	importArray(C_import_b_bank_eng_addr3)			= Trim(Request("txtEngAddr3"))
    importArray(C_import_b_bank_par_bank_cd)		= ""
    importArray(C_import_b_bank_addr4)				= ""
    importArray(C_import_b_bank_bank_fg)			= ""

    strZipcodechk			= Trim(Request("txtzipcodechk"))
    
	If lgIntFlgMode = OPMD_CMODE Then
		pvCommandSent = "CREATE"
	ElseIf lgIntFlgMode = OPMD_UMODE Then
		pvCommandSent = "UPDATE"
	End If

	Dim arrRowVal																	'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim arrColVal																	'☜: Spread Sheet 의 값을 받을 Array 변수 
	Dim strStatus
	Dim arrVal																	'☜: Sheet 의 개별 Row의 상태 (Create/Update/Delete)
	
	strGroup = Request("txtSpread")
	IF  Request("txtSpread") <> "" Then

	    Call pPB2SA05.B_MANAGE_BANK_SVR(gStrGlobalCollection, pvCommandSent, importArray, CStr(strGroup), strZipcodechk)


		If CheckSYSTEMError(Err,True) = True Then		
			Set pPB2SA05 = Nothing																	'☜: ComProxy Unload
			Response.End																			'☜: 비지니스 로직 처리를 종료함 
		End If	    
		
		Set pPB2SA05= Nothing	
		

	Else

	    Call pPB2SA05.B_MANAGE_BANK_SVR(gStrGlobalCollection, pvCommandSent, importArray, CStr(strGroup), strZipcodechk)

		'-----------------------
		'Com action result check area(OS,internal)
		'-----------------------
		If CheckSYSTEMError(Err,True) = True Then
			Set pPB2SA05 = Nothing																	'☜: ComProxy Unload
			Response.End																			'☜: 비지니스 로직 처리를 종료함	
		End If


	End If
	
	Set pPB2SA05 = Nothing																'☜: Unload Comproxy

%>
<Script Language=vbscript>

	With parent																		'☜: 화면 처리 ASP 를 지칭함 
		.DbSaveOk
	End With
</Script>
