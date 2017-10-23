'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID  = "a5103mb1_KO441.asp"
Const BIZ_PGM_ID2 = "a5103mb2_KO441.asp"
Const BIZ_PGM_ID3 = "a5103mb3_KO441.asp"
'==========================================  1.2.1 Global 상수 선언  ======================================
'=                       4.2 Constant variables 
'========================================================================================================
Const GRID_POPUP_MENU_NEW	=	"0000111111"
Const GRID_POPUP_MENU_CRT	=	"0000111111"
Const GRID_POPUP_MENU_UPD	=	"0001111111"
Const GRID_POPUP_MENU_PRT	=	"0000111111"		

'==========================================================================================================

' Grid constant
Dim  C_Confirm     
Dim  C_Conf_Nm
Dim  C_Conf_fg     
Dim  C_TempGlDt    
Dim  C_GlDt        
Dim  C_TempGlNo    
Dim  C_DeptNm      
Dim  C_Currency    
Dim  C_TempGlAmt   
Dim  C_TempGlLocAmt
Dim  C_GlNo         
Dim	 C_TempGlDesc	'적요 
Dim	 C_RefNo
Dim  C_USER	

Dim lgStrPrevKeyTempGlNo
Dim lgStrPrevKeyTempGlDt
Dim lgQueryFlag					' 신규조회 및 추가조회 구분 Flag
Dim lgGridPoupMenu              ' Grid Popup Menu Setting
Dim lgAllSelect

Dim lgIsOpenPop
Dim IsOpenPop       
Dim lgPageNo_B
Dim lgSortKey_B
 
Const C_MaxKey = 3
'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'                        5.1 Common Method-1
'========================================================================================================= 
Sub InitSpreadPosVariables()
    C_Confirm      = 1	
    C_Conf_Nm      = 2
    C_Conf_Fg      = 3
    C_USER         = 4 
    C_TempGlDt     = 5 
    C_GlDt         = 6 	
    C_TempGlNo     = 7 
    C_DeptNm       = 8 
    C_Currency     = 9 
    C_TempGlAmt    = 10
    C_TempGlLocAmt = 11 
    C_GlNo         = 12
    C_TempGlDesc   = 13
    C_RefNo		   = 14
    
End Sub

'========================================================================================================= 
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE					'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False							'Indicates that no value changed
    lgIntGrpCount = 0									'initializes Group View Size
   
    lgStrPrevKeyTempGlNo = ""							'initializes Previous Key
    lgLngCurRows = 0									'initializes Deleted Rows Count
    
    lgPageNo_B		 = ""                               'initializes Previous Key for spreadsheet #2    
    lgSortKey_B      = "1"
    
    lgStrPrevKeyTempGlDt = ""
	lgStrPrevKeyTempGlNo = ""    
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	'승인의 일자는 당일의 일자만 조회한다.
    Dim EndDate
	EndDate = UniConvDateAToB(iDBSYSDate ,parent.gServerDateFormat,parent.gDateFormat)
	'승인의 일자는 당일 ~ 당일 이다. 
	frm1.txtFromReqDt.text  =  EndDate
	frm1.txtToReqDt.text    =  EndDate
	frm1.GIDate.text        =  EndDate
	frm1.cboConfFg.value    =	"U"	
	frm1.txtDeptCd.focus
	lgGridPoupMenu          = GRID_POPUP_MENU_PRT
	frm1.hOrgChangeId.value = parent.gChangeOrgId
End Sub

'========================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	With frm1.vspdData	
        .MaxCols = C_RefNo + 1										'☜: 최대 Columns의 항상 1개 증가시킴 
        .Col = .MaxCols												'☆: 사용자 별 Hidden Column
        .ColHidden = True
        
        .MaxRows = 0
        ggoSpread.Source = frm1.vspdData
        .ReDraw = False
        
        ggoSpread.Spreadinit "V20021127",, parent.gAllowDragDropSpread
        .ReDraw = False

        Call GetSpreadColumnPos("A")

        ggoSpread.SSSetCheck C_Confirm,      "",     8,  -10, "", True, -1         
        ggoSpread.SSSetEdit  C_Conf_Nm,      "", 8, 2,,3                  
        ggoSpread.SSSetEdit  C_Conf_Fg,      "결재여부",   10, ,,10
        ggoSpread.SSSetEdit	 C_USER,		 "결재자",   10,,,10
        ggoSpread.SSSetDate  C_TempGlDt,     "결의전표일", 13, 2, parent.gDateFormat
        ggoSpread.SSSetDate  C_GlDt,         "전표일",     13, 2, parent.gDateFormat
        ggoSpread.SSSetEdit  C_TempGlNo,     "결의번호",   15, ,,18
        ggoSpread.SSSetEdit  C_DeptNm,       "부서명",     20,  ,,30
        ggoSpread.SSSetEdit  C_Currency,     "통화",        8, 2,,3                                  
        ggoSpread.SSSetFloat C_TempGlAmt,    "금액",       18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
        ggoSpread.SSSetFloat C_TempGlLocAmt, "금액(자국)", 18, parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec    
        ggoSpread.SSSetEdit  C_GlNo,         "전표번호",   13, ,,18
        ggoSpread.SSSetEdit	 C_TempGlDesc,	 "적요",	   20,,,128
        ggoSpread.SSSetEdit	 C_RefNo,		 "참조번호",   13,,,30
        



        Call ggoSpread.SSSetColHidden(C_Conf_Nm,C_Conf_Nm,True)
        Call ggoSpread.SSSetColHidden(C_Currency, C_Currency, True)

        .ReDraw = True
    End With

	Call SetZAdoSpreadSheet("A5103MA101","S","A","V20051211",Parent.C_SORT_DBAGENT,frm1.vspdData2, C_MaxKey, "X","X")
    Call SetSpreadLock()
    Call SetSpreadLock_B()
End Sub

'========================================================================================
Sub SetSpreadLock()
    With frm1
		ggoSpread.Source = .vspdData
        .vspdData.ReDraw = False
        
		frm1.vspddata.col = C_Confirm
		frm1.vspddata.row = 0
		frm1.vspddata.ColHeadersShow = True
		
        ggoSpread.SpreadLock        C_Conf_Fg      , -1    ,C_Conf_Fg
        ggoSpread.SpreadLock        C_TempGlDt      , -1    ,C_TempGlDt
        ggoSpread.SpreadLock        C_TempGlNo      , -1    ,C_TempGlNo
        ggoSpread.SpreadLock        C_DeptNm        , -1    ,C_DeptNm
        ggoSpread.SpreadLock        C_Currency      , -1    ,C_Currency
        ggoSpread.SpreadLock        C_TempGlAmt     , -1    ,C_TempGlAmt
        ggoSpread.SpreadLock        C_TempGlLocAmt  , -1    ,C_TempGlLocAmt
        ggoSpread.SSSetRequired     C_GlDt          , -1    ,C_GlDt    
        ggoSpread.SpreadLock		C_TempGlDesc	, -1	,C_TempGlDesc
        ggoSpread.SpreadLock		C_RefNo			, -1	,C_RefNo 
        ggoSpread.SpreadLock		C_USER		, -1	,C_USER
        
        
        ggoSpread.SSSetProtected	.vspdData.MaxCols,-1	,-1
        .vspdData.ReDraw = True
    End With
End Sub

Sub SetSpreadLock_B()
	With frm1
        .vspdData2.ReDraw = False
		ggoSpread.Source = .vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()	
        .vspdData2.ReDraw = True
	End With
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal lRow, ByVal lRow2)
	If  lRow2 = "" Then	lRow2 = lRow
    With frm1
		ggoSpread.Source = .vspdData    
	    .vspdData.ReDraw = False
		ggoSpread.SSSetProtected	C_TempGlDt,		lRow, lRow2
		ggoSpread.SSSetProtected	C_TempGlNo,		lRow, lRow2
		ggoSpread.SSSetProtected	C_DeptNm,		lRow, lRow2
		ggoSpread.SSSetProtected	C_Currency,		lRow, lRow2
		ggoSpread.SSSetProtected	C_TempGlAmt,	lRow, lRow2
		ggoSpread.SSSetProtected	C_TempGlLocAmt, lRow, lRow2
		
		If frm1.cboConfFg.value = "C" Then
			ggoSpread.SSSetProtected		C_GlDt,         lRow, lRow2
			ggoSpread.SpreadLock        C_GlNo          , -1    ,C_GlNo 
		Else
			ggoSpread.SSSetRequired		C_GlDt,         lRow, lRow2
			ggoSpread.SpreadUnLock        C_GlNo          , -1    ,C_GlNo 
		End If	
		'ggoSpread.SSSetProtected	C_GlNo,			lRow, lRow2 
		ggoSpread.SSSetProtected	C_TempGlDesc,	lRow, lRow2
		ggoSpread.SSSetProtected	C_RefNo,	lRow, lRow2
		
		.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"

            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
                C_Confirm      = iCurColumnPos(1)	
				C_Conf_Nm      = iCurColumnPos(2   )
				C_Conf_Fg      = iCurColumnPos(3   )
				C_USER         = iCurColumnPos(4   )
				C_TempGlDt     = iCurColumnPos(5   )
				C_GlDt         = iCurColumnPos(6 )
				C_TempGlNo     = iCurColumnPos(7   )
				C_DeptNm       = iCurColumnPos(8   )
				C_Currency     = iCurColumnPos(9   )
				C_TempGlAmt    = iCurColumnPos(10  )
				C_TempGlLocAmt = iCurColumnPos(11  )
				C_GlNo         = iCurColumnPos(12  )
				C_TempGlDesc   = iCurColumnPos(13  )
				C_RefNo		   = iCurColumnPos(14  )

    End Select    
End Sub

'========================================================================================================= 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(8), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.hOrgChangeId.value = parent.gChangeOrgId

	Select Case iWhere
		Case 1			
			arrParam(0) = "전표입력경로"			' 팝업 명칭 
			arrParam(1) = "B_MINOR" 					' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = " MAJOR_CD = " & FilterVar("A1001", "''", "S") & "  "		' Where Condition
			arrParam(5) = "전표입력경로"			' 조건필드의 라벨 명칭 

			arrField(0) = "MINOR_CD"					' Field명(0)
			arrField(1) = "MINOR_NM"					' Field명(1)
    
			arrHeader(0) = "전표입력경로"			' Header명(0)
			arrHeader(1) = "전표입력경로명"			' Header명(1)
			
		Case 2,3
			arrParam(0) = frm1.txtFromReqDt.Text
			arrParam(1) = frm1.txtToReqDt.Text					
			
			' 권한관리 추가 
			arrParam(5) = lgAuthBizAreaCd
			arrParam(6) = lgInternalCd
			arrParam(7) = lgSubInternalCd
			arrParam(8) = lgAuthUsrID

		Case 4,5
			arrParam(0) = frm1.txtFromReqDt.Text
			arrParam(1) = frm1.txtToReqDt.Text

			' 권한관리 추가 
			arrParam(5) = lgAuthBizAreaCd
			arrParam(6) = lgInternalCd
			arrParam(7) = lgSubInternalCd
			arrParam(8) = lgAuthUsrID

	End Select
	
	Select Case iWhere
		Case 2,3
			arrRet = window.showModalDialog("a5101ra1.asp", Array(window.parent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	
		Case 4,5
			arrRet = window.showModalDialog("A5104RA1.asp", Array(window.parent, arrParam), _
		     "dialogWidth=660px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")	
		Case Else    
			arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select		
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPopUp(arrRet, iWhere)
	End If	

	Call FocusAfterPopup (iWhere)
End Function

'=======================================================================================================
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtFromReqDt.text				'  Code Condition
   	arrParam(1) = frm1.txtToReqDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' 결의일자 상태 Condition  

	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID
		
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		Call SetDept(arrRet)
		frm1.txtDeptCd.focus
	End If	
End Function
'========================================================================================================= 

Function SetDept(Byval arrRet)
	frm1.hOrgChangeId.value = arrRet(2)	
	frm1.txtDeptCd.value    = arrRet(0)
	frm1.txtDeptNm.value    = arrRet(1)		
	frm1.txtFromReqDt.text  = arrRet(4)
	frm1.txtToReqDt.text    = arrRet(5)
End Function

'========================================================================================================= 
Function OpenPopupTempGL ()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5130ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	With frm1.vspdData
		.Row = .ActiveRow
		.Col =  C_TempGlNo
		
		arrParam(0) = Trim(.Text)	'결의전표번호 
		arrParam(1) = ""			'Reference번호 
	End With

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'========================================================================================================= 
Function OpenPopupGL()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(1)	

	If IsOpenPop = True Then Exit Function

	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	With frm1.vspdData
		.Row = .ActiveRow
		.Col =  C_GlNo
		
		arrParam(0) = Trim(.Text)	'회계전표팝업 
		arrParam(1) = ""			'Reference번호 
	End With

	IsOpenPop = True
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
End Function

'========================================================================================================= 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)			
			Case 1
				.txtGlInputType.value = arrRet(0)
					.txtGlInputTypeNm.value = arrRet(1)		
			Case 2		'결의전표번호 
				frm1.txtTempGlNoFr.value = UCase(Trim(arrRet(0)))
			Case 3		'결의전표번호 
				frm1.txtTempGlNoTo.value = UCase(Trim(arrRet(0)))
			Case 4		'회계전표번호 
				frm1.txtGlNoFr.value = UCase(Trim(arrRet(0)))
			Case 5		'회계전표번호 
				frm1.txtGlNoTo.value = UCase(Trim(arrRet(0)))
		End Select
	End With	
End Function

'=======================================================================================================
Function FocusAfterPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtDeptCd.focus
			Case 1  
				.txtGlInputType.focus
			Case 2  
				.txtTempGlNoFr.focus
			Case 3 
				.txtTempGlNoTo.focus
		End Select    
	End With
End Function

'========================================================================================================= 
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.txtDeptCd.focus
			Case 1  
				.txtGlInputType.focus
			Case 2  
				.txtTempGlNoFr.focus
			Case 3 
				.txtTempGlNoTo.focus
		End Select    
	End With
End Function


'----------------------------------------  OpenBizAreaCd()  -------------------------------------------------
'	Name : BizAreaCd()
'	Description : Business PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizAreaCd(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "사업장 팝업"				' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = ""							' Where Condition
	
	' 권한관리 추가 
	If lgAuthBizAreaCd <>  "" Then
		arrParam(4) = " BIZ_AREA_CD=" & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	
	arrParam(5) = "사업장 코드"			

    arrField(0) = "BIZ_AREA_CD"					' Field명(0)
    arrField(1) = "BIZ_AREA_NM"					' Field명(1)

    arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
'		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,iWhere)
	End If
End Function

'=======================================================================================================
'	Name : SetReturnVal()
'	Description : 
'=======================================================================================================
Function SetReturnVal(byval arrRet,Field_fg)
	Select Case Field_fg
		case 0
			frm1.txtBizAreaCd.Value	= arrRet(0)
			frm1.txtBizAreaNm.Value	= arrRet(1)
			frm1.txtBizAreaCd.focus
		case 1
			frm1.txtBizAreaCd1.Value	= arrRet(0)
			frm1.txtBizAreaNm1.Value	= arrRet(1)
			frm1.txtBizAreaCd1.focus
	End Select
	
	lgBlnFlgChgValue = True
End Function

'========================================================================================================= 
Sub fnBttnConf()	
	Dim ii 
	
	With frm1
		For ii = 1 To .vspddata.MaxRows
			.vspddata.row = ii
			.vspddata.col = C_Confirm
			.vspddata.value = "1"
		Next	
	End With		
End Sub

'========================================================================================================= 
Function fnBttnUnConf()
	Dim ii 
	
	With frm1
		For ii = 1 To .vspddata.MaxRows
			.vspddata.row = ii
			.vspddata.col = C_Confirm
			.vspddata.value = "0"
		Next	
	End With		
End Function

'========================================================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : 
'========================================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	Dim iGridPos

	Select Case UCase(Trim(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			iGridPos = "A"
		Case "VSPDDATA2"			
			iGridPos = "B"
	End Select			

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(iGridPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		Exit Function
	Else
		Call ggoSpread.SaveXMLData(iGridPos,arrRet(0),arrRet(1))
		Call InitVariables()
		Call InitSpreadSheet()       
   End If
End Function

'========================================================================================================
' Function Name : CompareTempGlNoByDB
' Function Desc : 
'========================================================================================================
'==========================================================================================
Function CompareTempGlNoByDB(ByVal FromNo , ByVal ToNo)
	Dim strSelect,strFrom,strWhere
	Dim iFlag,iRs

	CompareTempGlNoByDB = False

    If FromNo.value <> "" And ToNo.value <> "" Then
        strSelect = ""
        strSelect = "  Case When  " & FilterVar(UCase(FromNo.value), "''", "S") & " "
        strSelect = strSelect & "  >  " & FilterVar(UCase(ToNo.value), "''", "S") & "  Then " & FilterVar("N", "''", "S") & "  "
        strSelect = strSelect & " When  " & FilterVar(UCase(FromNo.value), "''", "S") & " "
        strSelect = strSelect & "  <=  " & FilterVar(UCase(ToNo.value), "''", "S") & "  Then " & FilterVar("Y", "''", "S") & "  End "
        strFrom = ""
        strWhere = ""
        If CommonQueryRs2by2(strSelect, strFrom, strWhere, iRs) = True Then
            iFlag = Split(iRs, Chr(11))
            If Trim(iFlag(1)) = "N" Then
                Exit Function
            End If
        Else
            Exit Function
        End If
    End If
    
    CompareTempGlNoByDB = True
End Function


'#########################################################################################################
'												3. Event부 
'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'==========================================================================================
Sub txtDeptCD_OnChange()
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj

	If frm1.txtDeptCd.value = "" Then
		frm1.txtDeptNm.value = ""
	End If
	
    lgBlnFlgChgValue = True
	
	If TRim(frm1.txtDeptCd.value) <>"" Then
		strSelect = "dept_cd, ORG_CHANGE_ID, DEPT_NM"
		strFrom   =  " B_ACCT_DEPT "
		strWhere  = " ORG_CHANGE_DT >= "
		strWhere  = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtFromReqDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ")"
		strWhere  = strWhere & " AND ORG_CHANGE_DT <= " 
		strWhere  = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtToReqDt.Text, gDateFormat,Parent.gServerDateType), "''", "S") & ") "
		strWhere  =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

		' 권한관리 추가 
		If lgInternalCd <> "" Then
			strWhere  = strWhere & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  
		End If
	
		If lgSubInternalCd <> "" Then
			strWhere  = strWhere & " AND INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
		End If

		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus
		Else 
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			

			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				frm1.txtDeptNm.value = Trim(arrVal2(3))
			Next	
		End If
	End If
End Sub

'========================================================================================================= 
Sub txtFromReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtToReqDt.focus
		Call FncQuery
	End If
End Sub
'========================================================================================================= 

Sub txtToReqDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then 
		frm1.txtFromReqDt.focus
		Call FncQuery
	End If
End Sub
'========================================================================================================= 

Sub txtBizAreaCd_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		exit sub
	ElseIf KeyAscii = 13 Then 
		Call FncQuery
	End If
End Sub

'========================================================================================================= 
Sub txtBizAreaCd1_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		exit sub
	ElseIf KeyAscii = 13 Then 
		Call FncQuery
	End If
End Sub

'========================================================================================================= 
Sub vspdData_Click(ByVal Col , ByVal Row)
	Dim StrConf1

    Call SetPopupMenuItemInf(lgGridPoupMenu)

    gMouseClickStatus = "SPC"   

	lgPageNo_B = ""
    
    With frm1
		Set gActiveSpdSheet = .vspdData

		If .vspdData.MaxRows <= 0 Then                                                    'If there is no data.
			Exit Sub
   		End If

		ggoSpread.Source  = .vspdData
   	
		Select Case Col
			Case C_Confirm 	
				.vspdData.col = C_Confirm
				StrConf1  = .cboConfFg.value
				If .vspdData.text = "0" Then
					If StrConf1 = "C" Then
						StrConf1 = "U"
					Else
						StrConf1 = "C"
					End If	
				End If	
				
				.vspdData.col =	C_Conf_nm			

				If .vspdData.Value <> StrConf1 Then
					lgBlnFlgChgValue = True						
				Else
					ggoSpread.EditUndo                                                  '☜: Protect system from crashing				
				End If	
		End Select 
	End With    
   		    
	If Row <= 0 Then
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey
	        lgSortKey = 1
	    End If
	    Exit Sub
	End If

	Call DbQuery("2",Row)    
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData	

    lgPageNo_B		 = ""                               'initializes Previous Key for spreadsheet #2    
    lgSortKey_B      = "1"	
End Sub

'========================================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub  vspdData2_Click(ByVal Col, ByVal Row)
	Call SetPopUpMenuItemInf("00000000001")
	
    gMouseClickStatus = "SP2C"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData2            
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 1
        End If    
        Exit Sub
    End If
    
	Call SetSpreadColumnValue("A",frm1.vspdData2,Col,Row)	
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    If Col <= C_Confirm Or NewCol <= C_Confirm Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'==========================================================================================
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub

'==========================================================================================
Sub txtFromReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromReqDt.focus
    End If
End Sub

Sub txtFromReqDt_Change()
    lgBlnFlgChgValue = True
End Sub

'==========================================================================================
Sub txttoReqDt_DblClick(Button)
    If Button = 1 Then
        frm1.txttoReqDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txttoReqDt.focus
    End If
End Sub
'========================================================================================================= 

Sub txttoReqDt_Change(Button)
    lgBlnFlgChgValue = True
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : 
'=======================================================================================================
Sub  vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    gMouseClickStatus = "SPC"	'Split 상태코드 
	Set gActiveSpdSheet = frm1.vspdData

    If Row <> NewRow And NewRow > 0 Then
	    If NewRow = 0 Then
		    Exit Sub
	    End If

		Call DbQuery("2",NewRow)

		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
	
		lgPageNo_B		 = ""
		lgSortKey_B      = 1
    End If
End Sub

'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================
Sub vspdData_KeyPress(index , KeyAscii )
     lgBlnFlgChgValue = True													'⊙: Indicates that value changed
End Sub

'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    Dim LngLastRow    
    Dim LngMaxRow     
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
		If lgStrPrevKeyTempGlNo <> "" Then                         
			Call DbQuery("1",frm1.vspddata.row)
		End If
	End If		
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub  vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then	'☜: 재쿼리 체크'
		If lgPageNo_B <> "" Then													'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
           Call DbQuery("2",frm1.vspddata.ActiveRow)
		End If
   End if
End Sub

'#########################################################################################################
'												4. Common Function부 
'=========================================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False																		'⊙: Processing is NG

    Err.Clear																				'☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
    With frm1
	    ggoSpread.Source = .vspdData
	    If  ggoSpread.SSCheckChange = True Then
			IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")					'데이타가 변경되었습니다. 조회하시겠습니까?
	    	If IntRetCD = vbNo Then
		      	Exit Function
	    	End If
	    End If

		'-----------------------
	    'Check condition area
	    '-----------------------
		If Not chkFieldByCell(.txtFromReqDt, "A", "1") Then Exit Function
		If Not chkFieldByCell(.txtToReqDt, "A", "1") Then Exit Function
		If Not chkFieldByCell(.cboConfFg, "A", "1") Then Exit Function    

	'    If Not chkField(Document, "1") Then												'⊙: This function check indispensable field     
	'		Exit Function
	'    End If

	    If CompareDateByFormat(.txtFromReqDt.text,.txtToReqDt.text,.txtFromReqDt.Alt,.txtToReqDt.Alt, _
	                        "970025",.txtFromReqDt.UserDefinedFormat,parent.gComDateType,True) = False Then		
			Exit Function	
	    End If

		If CompareTempGlNoByDB(.txtTempGlNoFr,.txtTempGlNoTo) = False Then
		    Call DisplayMsgBox("970025", "X", .txtTempGlNoFr.Alt, .txtTempGlNoTo.Alt)
		    frm1.txtTempGlNoFr.focus
			Exit Function
		End If		

		If Trim(frm1.txtBizAreaCd.value) <> "" and Trim(frm1.txtBizAreaCd1.value) <> "" Then				
			If UCase(Trim(frm1.txtBizAreaCd.value)) > UCase(Trim(frm1.txtBizAreaCd1.value)) Then
				IntRetCd = DisplayMsgBox("970025", "X", frm1.txtBizAreaCd.Alt, frm1.txtBizAreaCd1.Alt)
				frm1.txtBizAreaCd.focus
				Exit Function
			End If
		End If
		
		If frm1.txtBizAreaCd.value = "" Then
			frm1.txtBizAreaNm.value = ""
		End If
	
		If frm1.txtBizAreaCd1.value = "" Then
			frm1.txtBizAreaNm1.value = ""
		End If
		
	    '-----------------------
	    'Erase contents area
	    '-----------------------
'	    Call ggoOper.ClearField(Document, "2")												'⊙: Clear Contents  Field
	    ggoSpread.Source = frm1.vspdData
	    ggoSpread.ClearSpreadData
	    
	    ggoSpread.Source = frm1.vspdData2
	    ggoSpread.ClearSpreadData

	    Call InitVariables 																	'⊙: Initializes local global variables

		If .txtDeptCd.value = "" Then
			.txtDeptNm.value = ""
		End If

		lgQueryFlag = "New"		' 신규조회 및 추가조회 구분 Flag (현재는 신규임)

	    '-----------------------
	    'Query function call area
	    '-----------------------
		If  DbQuery("1",frm1.vspddata.row) = False Then									'☜: Query db data
			Exit Function	
		End If

	    FncQuery = True																		'⊙: Processing is OK
	End With	    
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False																		'⊙: Processing is NG
    
    Err.Clear																			'☜: Protect system from crashing
    'On Error Resume Next																'☜: Protect system from crashing
    
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분    
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
     
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "1")												'⊙: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")												'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    
	Call LockObjectField(.txtFromReqDt,"R")
	Call LockObjectField(.txtToReqDt,"R")
	Call LockHTMLField(.cboConfFg,"R")
	Call LockHTMLField(.txtDeptNm,"P")
	Call LockHTMLField(.txtGlInputTypeNm,"P")        				    
    
'    Call ggoOper.LockField(Document, "N")												'⊙: Lock  Suitable  Field
    Call SetDefaultVal
    Call InitVariables																	'⊙: Initializes local global variables
    
    FncNew = True																		'⊙: Processing is OK
End Function


'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    FncSave = False																		'⊙: Processing is NG
    
    Err.Clear																			'☜: Protect system from crashing
    On Error Resume Next																'☜: Protect system from crashing

    '-----------------------
    'Precheck area
    '-----------------------
    With frm1

	    ggoSpread.Source = .vspdData

	    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False   Then				'⊙: Check If data is chaged
	        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")								'⊙: Display Message(There is no changed data.)
	        Exit Function
	    End If

		If Not chkFieldByCell(.txtFromReqDt, "A", "1") Then Exit Function
		If Not chkFieldByCell(.txtToReqDt, "A", "1") Then Exit Function
		If Not chkFieldByCell(.cboConfFg, "A", "1") Then Exit Function   

	'    If Not chkField(Document, "1") Then													'⊙: Check required field(Single area)
	'		Exit Function
	'    End If

		'-----------------------
	    'Check content area
	    '----------------------- 
	    ggoSpread.Source = .vspdData
	    If Not ggoSpread.SSDefaultCheck Then												'⊙: Check contents area
			Exit Function
	    End If
	End With    
    '-----------------------
    'Save function call area
    '-----------------------
    IF  DbSave	= False Then			                                                '☜: Save db data 
		Exit Function	
    End If    
   	
    FncSave = True																		'⊙: Processing is OK
End Function

'========================================================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo																	'☜: Protect system from crashing    
End Function

'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()																'☜: Protect system from crashing
End Function

'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)												'☜: 화면 유형 
End Function


'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)											'☜:화면 유형, Tab 유무 
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'=======================================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")								'데이타가 변경되었습니다. 종료 하시겠습니까?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True
End Function

'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================
Function DbQuery(ByVal iOpt,ByVal Row) 
	Dim strVal

    DbQuery = False    
    Call LayerShowHide(1)
    frm1.btnConf.disabled  = True
	frm1.btnUnCon.disabled = True

	On Error Resume Next
    Err.Clear																						'☜: Protect system from crashing
    
	With frm1
		Select Case iOpt 
			Case "1"  	
				If lgIntFlgMode = parent.OPMD_UMODE Then
					strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001								'☜:조회표시 
					strVal = strVal & "&lgStrPrevKeyTempGlDt=" & lgStrPrevKeyTempGlDt
					strVal = strVal & "&lgStrPrevKeyTempGlNo=" & lgStrPrevKeyTempGlNo
					strVal = strVal & "&txtDeptCd=" & Trim(.hDeptcd.value)
					strVal = strVal & "&cboConfFg=" & Trim(.hcboConfFg.value)
				'	if Trim(.hcboConfFg.value) = "최종승인" Then
				'		strVal = strVal & "&cboConfFg=최종결재"
				'	else
				'		strVal = strVal & "&cboConfFg=" & Trim(.hcboConfFg.value)
				'	end if
					strVal = strVal & "&txtGlInputType=" & Trim(.txtGlInputType.value)			
					strVal = strVal & "&txtFromReqDt=" & (.txtFromReqDt.text)
					strVal = strVal & "&txtToReqDt=" & (.txtToReqDt.text)
					strVal = strVal	& "&txtTempGlNoFr=" & .txtTempGlNoFr.value
					strVal = strVal & "&txtTempGlNoTo=" & .txtTempGlNoTo.value
					strVal = strVal	& "&txtTempGlNoFr=" & .txtTempGlNoFr.value
					strVal = strVal & "&txtGlNoTo=" & .txtGlNoTo.value
					strVal = strVal	& "&txtGlNoFr=" & .txtGlNoFr.value
					strVal = strVal	& "&txtRefNo=" & .txtRefNo.value
					strVal = strVal & "&txtBizAreaCd=" & Trim(.htxtBizAreaCd.value)
					strVal = strVal & "&txtBizAreaCd1=" & Trim(.htxtBizAreaCd1.value) 			
					strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
					strVal = strVal & "&hOrgChangeId=" & Trim(.hOrgChangeId.Value)
				Else
					strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001								'☜:조회표시 
					strVal = strVal & "&lgStrPrevKeyTempGlDt=" & lgStrPrevKeyTempGlDt
					strVal = strVal & "&lgStrPrevKeyTempGlNo=" & lgStrPrevKeyTempGlNo
					strVal = strVal & "&txtDeptCd=" & Trim(.txtDeptcd.value)
					strVal = strVal & "&cboConfFg=" & Trim(.cboConfFg.value)
				'	if Trim(.cboConfFg.value) = "최종승인" Then
				'		strVal = strVal & "&cboConfFg=최종결재"
				'	else
				'		strVal = strVal & "&cboConfFg=" & Trim(.cboConfFg.value)
				'	end if
					strVal = strVal & "&txtGlInputType=" & Trim(.txtGlInputType.value)			
					strVal = strVal & "&txtFromReqDt=" & (.txtFromReqDt.text)
					strVal = strVal & "&txtToReqDt=" & (.txtToReqDt.text)
					strVal = strVal	& "&txtTempGlNoFr=" & .txtTempGlNoFr.value
					strVal = strVal & "&txtTempGlNoTo=" & .txtTempGlNoTo.value
					strVal = strVal & "&txtGlNoTo=" & .txtGlNoTo.value
					strVal = strVal	& "&txtGlNoFr=" & .txtGlNoFr.value
					strVal = strVal	& "&txtRefNo=" & .txtRefNo.value
					strVal = strVal & "&txtBizAreaCd=" & Trim(.txtBizAreaCd.value)
					strVal = strVal & "&txtBizAreaCd1=" & Trim(.txtBizAreaCd1.value)
					strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
					strVal = strVal & "&hOrgChangeId=" & Trim(.hOrgChangeId.Value)
				End If
			Case "2"
				.vspddata.col = C_TempGlNo
				.vspddata.row = Row

				strVal = BIZ_PGM_ID3 & "?txtTempGlNo=" & .vspddata.value                     
				strVal = strVal	& "&lgPageNo=" & lgPageNo_B											'☜: Next key tag				
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
				strVal = strVal & "&lgTailList=" & MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList=" & EnCoding(GetSQLSelectList("A"))		
		End Select 
	End With						

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd=" & lgAuthBizAreaCd				' 사업장			
	strVal = strVal & "&lgInternalCd=" & lgInternalCd			' 내부부서 
	strVal = strVal & "&lgSubInternalCd=" & lgSubInternalCd		' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID=" & lgAuthUsrID				' 개인 

	Call RunMyBizASP(MyBizASP, strVal)																'☜: 비지니스 ASP 를 가동 
   
    DbQuery = True
End Function

'========================================================================================
Function DbQueryOk(ByVal iOpt)																		'☆: 조회 성공후 실행로직 
Dim i
Dim IntRetCD 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE																'⊙: Indicates that current mode is Update mode

	With frm1

		'If iOpt = 1 Then
                'For i = 1 To .vspdData.MaxRows
    
	'		.vspdData.Row = i
	'		.vspdData.Col = C_TempGlNo
        '               IntRetCD = CommonQueryRs(" (SELECT USR_NM FROM Z_USR_MAST_REC WHERE USR_ID =  PROJECT_NO) , CASE ISNULL(DIST_TYPE,'') WHEN  '' THEN '최초등록' WHEN 'Y' THEN '상신' WHEN 'N' THEN '반려' WHEN 'E' THEN '최종승인' END ", "A_TEMP_GL", "TEMP_GL_NO =" & FilterVar(.vspdData.value, "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) 
        '                If IntRetCD = False Then
        '                Else
        '                .vspdData.Col = C_USER
        '                .vspdData.Value = Trim(Replace(lgF0,Chr(11),""))
        '                .vspdData.Col = C_Conf_Fg
        '                .vspdData.Value = Trim(Replace(lgF1,Chr(11),""))
        '                End If
        '        NEXT
        '		End If			
  
		If iOpt = 1 Then
			Call vspdData_Click(1,1)
			.vspdData.focus
		End If			

'		Call LockObjectField(.txtFromReqDt,"R")
'		Call LockObjectField(.txtToReqDt,"R")
'		Call LockHTMLField(.cboConfFg,"R")		
		Call LockHTMLField(.txtDeptNm,"P")
		Call LockHTMLField(.txtGlInputTypeNm,"P")        

'		Call ggoOper.LockField(Document, "Q")															'⊙: This function lock the suitable field
		Call LayerShowHide(0)
		Call SetToolbar("110010000001111")																'⊙: 버튼 툴바 제어 
    
		lgGridPoupMenu  =   GRID_POPUP_MENU_UPD
		SetSpreadColor 1, .vspddata.maxrows
	
		If .vspdData.MaxRows > 0 Then	
			.btnConf.disabled	= False
			.btnUnCon.disabled	= False
		End If
	End With		
End Function

'======================================================================================================
Function SetGridFocus()
	with frm1
		.vspdData.Row = 1
		.vspdData.Col = 1
		.vspdData.Action = 1
	end with 
End Function 


'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	Dim iSelectCnt
		
    DbSave = False                                                          '⊙: Processing is NG
    Call LayerShowHide(1)
    
    On Error Resume Next                                                   '☜: Protect system from crashing

	With frm1
		.txtMode.value = parent.UID_M0002
		.txtUpdtUserId.value = parent.gUsrID
		.txtInsrtUserId.value = parent.gUsrID
		
		'-----------------------
		'Data manipulate area
		'-----------------------
		lGrpCnt = 1
		strVal = ""
		iSelectCnt = 0
		lgAllSelect = False
		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col = C_Confirm
			
			IF frm1.vspdData.text = "1" THEN
			
				strVal = strVal & "U" & parent.gColSep				'☜: U=Update
		
				.vspdData.Col = C_TempGlNo		'4
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep										
				.vspdData.Col = C_GlDt
				strVal = strVal & Trim(.vspdData.Text) & parent.gColSep										
				IF Trim(.cboConfFg.value) = "U" Then 
					strVal = strVal & "C" & parent.gColSep
				Else
					strVal = strVal & "U" & parent.gColSep
				END IF	
				
				.vspdData.Col = C_GlNo
				strVal = strVal & "" & Trim(.vspdData.Text) & parent.gRowSep
				
				lGrpCnt = lGrpCnt + 1
				iSelectCnt = iSelectCnt + 1	                
			END IF
		Next
	
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal			
		
		If iSelectCnt = .vspdData.MaxRows Then
			lgAllSelect = True
		End If

		.txthAuthBizAreaCd.value = lgAuthBizAreaCd
		.txthInternalCd.value    = lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value     = lgAuthUsrID
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)									'☜: 비지니스 ASP 를 가동 
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
End Function


'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
    Call LayerShowHide(0)
	
	frm1.vspdData.ReDraw = False
	ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData    
	frm1.vspdData.ReDraw = True
	Call InitVariables
	If lgAllSelect = True Then
		IF frm1.cboConfFg.value = "C" Then
			frm1.cboConfFg.value = "U"
		Else
			frm1.cboConfFg.value = "C"
		End If
	End If
	
	Call DbQuery("1",frm1.vspdData.row)
End Function

'########################################################################################
'# Area Name   : User-defined Method Part
'=======================================================================================================
Sub GIDate_DblClick(Button)
    If Button = 1 Then
        frm1.GIDate.Action = 7
        Call SetFocusToDocument("M")
        frm1.GIDate.focus    
    End If
End Sub

'=======================================================================================================
Sub  GIDate_Change()
	Dim gDate
	Dim IRow
	Dim gCnt

	gDate = frm1.GIDate.Text
	gCnt = frm1.vspdData.MaxRows

	frm1.vspdData.Col = C_GlDt

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If

	for IRow = 1 to frm1.vspdData.MaxRows
		frm1.vspdData.Row = IRow
		frm1.vspdData.Text	= gDate
	Next

	lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
Sub GIDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

'=======================================================================================================
Sub  cboConfFg_OnChange()  
    lgBlnFlgChgValue = True
End Sub
