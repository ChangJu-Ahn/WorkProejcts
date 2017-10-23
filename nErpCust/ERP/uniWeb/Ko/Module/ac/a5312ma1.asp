<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5312ma1
'*  4. Program Name         : a5312ma1
'*  5. Program Desc         : 외환평가 전표 처리 
'*  6. Component List       :
'*  7. Modified date(First) : 2002/01/15
'*  8. Modified date(Last)  : 2003/05/30
'*  9. Modifier (First)     : park jai hong
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<Script Language="VBScript">
Option Explicit																	'☜: indicates that All variables must be declared in advance
	

'========================================================================================================
Const BIZ_PGM_ID = "a5312mb1.asp"												'Biz Logic ASP
Const BIZ_SAVE_PGM_ID = "a5312mb2.asp"
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233

Dim C_MODULE_CD 
Dim C_MODULE_NM
Dim C_REF_NO
Dim C_TRAN_DATE
Dim C_ACCT_CD 
Dim C_ACCT_NM 
Dim C_DOC_CUR 
Dim C_XCH_RATE 
Dim C_ITEM_AMT 
Dim C_ITEM_LOC_AMT
Dim C_EVAL_XCH_RATE 
Dim C_EVAL_LOC_AMT 
Dim C_EVAL_LOSS_AMT 
Dim C_EVAL_PROFIT_AMT 

Dim IscookieSplit 

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop 

'========================================================================================================
Sub InitSpreadPosVariables()

	C_MODULE_CD = 1
	C_MODULE_NM = 2
	C_REF_NO = 3
	C_TRAN_DATE = 4
	C_ACCT_CD = 5
	C_ACCT_NM = 6	
	C_DOC_CUR = 7	
	C_XCH_RATE = 8															'Spread Sheet의 Column별 상수 
	C_ITEM_AMT = 9													
	C_ITEM_LOC_AMT = 10
	C_EVAL_XCH_RATE = 11
	C_EVAL_LOC_AMT = 12
	C_EVAL_LOSS_AMT = 13
	C_EVAL_PROFIT_AMT = 14

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE
	lgBlnFlgChgValue  = False
	lgIntGrpCount     = 0
    lgStrPrevKey      = ""
    lgStrPrevKeyIndex = ""
    lgSortKey         = 1		
    frm1.btnReflect.disabled= True
	frm1.btnCancle.disabled= True												'⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

	Dim StartDate
	Dim EndDate
	Dim strYear, strMonth, strDay

	StartDate	= "<%=GetSvrDate%>"                           'Get Server DB Date

	Call ExtractDateFrom(StartDate,Parent.gServerDateFormat,Parent.gServerDateType,strYear,strMonth,strDay)
	frm1.txtYyyymm.Text	= UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat)
	
	EndDate = UniConvDateAToB(UNIGetLastDay(StartDate,parent.gClientDateFormat) ,parent.gServerDateFormat,parent.gDateFormat)	
	frm1.txtGLDt.text = EndDate
	
	Call ggoOper.FormatDate(frm1.txtYyyymm, Parent.gDateFormat,2)

	frm1.hOrgChangeid.value = parent.gChangeOrgId
	frm1.txtYyyymm.focus
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>

ggAmtOfMoney.DecPoint  = 2
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
    'WriteCookie CookieSplit , lsConcd
	'FncQuery()
	Dim strCookie, i

	Const CookieSplit = 4877						

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	WriteCookie CookieSplit , IsCookieSplit

End Function

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()

End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()      '1.3 [initSpreadPosVariables] 호출 Logic 추가 

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20060424",,Parent.gAllowDragDropSpread            '2.1 [Spreadinit]에 Source Version No. 와 DragDrop지원여부를 설정 
	
	With frm1.vspdData
	
		.ReDraw = false
	
    	.MaxCols   = C_EVAL_PROFIT_AMT + 1                                                  ' ☜:☜: Add 1 to Maxcols
	    .Col       = .MaxCols                                                        ' ☜:☜: Hide maxcols
        .ColHidden = True           
       
		ggoSpread.Source= frm1.vspdData
		ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos("A")
        Call AppendNumberPlace("6","3","0")
       
        ggoSpread.SSSetEdit   C_MODULE_CD ,		"모듈구분코드"  ,5,,, 10
        ggoSpread.SSSetEdit   C_MODULE_NM ,		"모듈구분"      ,10,2,, 20
        ggoSpread.SSSetEdit   C_REF_NO ,		"관련번호"      ,20,,, 80
        ggoSpread.SSSetDate   C_TRAN_DATE ,		"발생일자"      ,12, 2, parent.gDateFormat
        ggoSpread.SSSetEdit   C_ACCT_CD ,		"계정코드"      ,10,,, 20
        ggoSpread.SSSetEdit   C_ACCT_NM ,		"계정명"        ,16,,, 30
		ggoSpread.SSSetEdit   C_DOC_CUR,        "거래통화"      , 8,2,, 10, 2
		ggoSpread.SSSetFloat  C_XCH_RATE,       "발생환율"      , 8, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_ITEM_AMT,       "잔액"          ,15, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_ITEM_LOC_AMT,   "잔액(자국)"    ,15, parent.ggAmTofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_EVAL_XCH_RATE,  "평가환율"      , 8, parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat  C_EVAL_LOC_AMT,   "평가금액(자국)",15, parent.ggAmTofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_EVAL_LOSS_AMT,  "평가손"        ,15, parent.ggAmTofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
	    ggoSpread.SSSetFloat  C_EVAL_PROFIT_AMT,"평가익"        ,15, parent.ggAmTofMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

	   Call ggoSpread.SSSetColHidden(C_MODULE_CD,C_MODULE_CD,True)
       
	   call ggoSpread.MakePairsColumn(C_MODULE_CD,C_MODULE_NM)	
	   call ggoSpread.MakePairsColumn(C_ACCT_CD,C_ACCT_NM)	
	   
	   .ReDraw = true
	
       Call SetSpreadLock 
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock      C_MODULE_CD , -1, C_EVAL_XCH_RATE
		ggoSpread.SpreadLock      C_EVAL_LOC_AMT , -1, C_EVAL_PROFIT_AMT
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
		.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub  SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow

    iPosArr = Split(iPosArr,Parent.gColSep)

    If IsNumeric(iPosArr(0)) Then
		iRow = CInt(iPosArr(0))
		For iDx = 1 To  frm1.vspdData.MaxCols - 1
		    Frm1.vspdData.Col = iDx
		    Frm1.vspdData.Row = iRow

		    If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
				Frm1.vspdData.Col = iDx
				Frm1.vspdData.Row = iRow
				Frm1.vspdData.Action = 0 ' go to 
				Exit For
		    End If
		Next
    End If   
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_MODULE_CD			= iCurColumnPos(1)
			C_MODULE_NM			= iCurColumnPos(2)
			C_REF_NO			= iCurColumnPos(3)    
			C_TRAN_DATE			= iCurColumnPos(4)
			C_ACCT_CD			= iCurColumnPos(5)    
			C_ACCT_NM			= iCurColumnPos(6)
			C_DOC_CUR			= iCurColumnPos(7)
			C_XCH_RATE			= iCurColumnPos(8)
			C_ITEM_AMT			= iCurColumnPos(9)
			C_ITEM_LOC_AMT		= iCurColumnPos(10)
			C_EVAL_XCH_RATE		= iCurColumnPos(11)
			C_EVAL_LOC_AMT		= iCurColumnPos(12)
			C_EVAL_LOSS_AMT		= iCurColumnPos(13)
			C_EVAL_PROFIT_AMT	= iCurColumnPos(14)
    End Select    
End Sub

'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()
	With frm1
		'Call ReFormatSpreadCellByCellByCurrency(.vspdData,1,.vspddata.MaxRows,C_DOC_CUR,C_ITEM_AMT,"A", "Q" ,"X","X")


		'Call ReFormatSpreadCellByCellByCurrency(.vspdData,1,.vspddata.MaxRows,C_DOC_CUR,C_ITEM_AMT ,"A","I","X","X") 

'		Call ReFormatSpreadCellByCellByCurrency(.vspdData,1,.vspddata.MaxRows,C_DOC_CUR,C_EVAL_XCH_RATE,"D", "Q" ,"X","X")
'		Call ReFormatSpreadCellByCellByCurrency(.vspdData,1,.vspddata.MaxRows,C_DOC_CUR,C_XCH_RATE,"D", "Q" ,"X","X")			
	End With	
End Sub

'========================================================================================================
Sub Form_Load()
    Err.Clear

	Call LoadInfTB19029

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call SetDefaultVal
    Call SetToolbar("1100000000001111")										        '버튼 툴바 제어 
   
	' 권한관리 추가 
	Dim xmlDoc

	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 

	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text

	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text

	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text

	Set xmlDoc = Nothing   
   
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub


'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field

    If Not chkField(Document, "1") Then									         '☜: This function check required field
		Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables

    If DbQuery = False Then
		Exit Function
    End If

    FncQuery = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If

    ggoSpread.Source = frm1.vspdData

    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
		Exit Function
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
		Exit Function
    End If
    
    If DbSave = False Then
        Exit Function
    End If    
    
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow() 

	If IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	Else
	    imRow = AskSpdSheetAddRowCount()
    
		If imRow = "" Then
		    Exit Function
		End If
	End If		

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imrow -1
       .vspdData.ReDraw = True
    End With
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
    
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================================
Function FncPrev() 
    On Error Resume Next                                                  '☜: Protect system from crashing
End Function

'========================================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================================
Function FncNext() 
    On Error Resume Next                                                   '☜: Protect system from crashing
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
Function DbQuery() 
    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    Dim strVal
    Dim iRetCd,IntRetCD

    DbQuery = False
call     CurFormatNumSprSheet
    Err.Clear																			'☜: Clear err status

	if LayerShowHide(1) = False then
	   Exit Function
	end if

    Call ExtractDateFrom(frm1.txtYYYYMM.Text,frm1.txtYYYYMM.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth

    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="           & Parent.UID_M0001
        strVal = strVal     & "&txtYYYYMM="         & strYYYYMM							'☜: Query Key
        strVal = strVal     & "&txtModuleCd="       & .txtModuleCd.Value					'☜: Query Key
        strVal = strVal     & "&txtBizAreaCd="      & .txtBizAreaCd.Value					'☜: Query Key
        strVal = strVal     & "&txtMaxRows="        & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex                 '☜: Next key tag
    End With

	' 권한관리 추가 
	strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
	strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
	strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
	strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

	Call RunMyBizASP(MyBizASP, strVal)													  '☜: Run Biz Logic
    
    DbQuery = True
End Function



'========================================================================================================
' Name : ExeReflect
' Desc : This function is data query and display
'========================================================================================================
Function ExeReflect() 

    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    Dim strVal
    Dim iRetCd,IntRetCD
    Dim strFirstDay
	
	ExeReflect = false
	err.Clear
	
	If Not chkField(Document, "1") Then									         '☜: This function check required field
		Exit Function
    End If
    
    if Trim(frm1.txtDeptCd.value)= "" then
		Call DisplayMsgBox("140621", parent.VB_INFORMATION, "X", "X")
		Exit Function
    End if

    iRetCd = DisplayMsgBox("900018", parent.VB_YES_NO,"전표생성 ","X")             'Will you destory previous data"

	If iRetCd = vbNo Then
		Exit Function
	End If

    Call ExtractDateFrom(frm1.txtYYYYMM.Text,frm1.txtYYYYMM.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth

	If LayerShowHide(1) = False Then
	   Exit Function
	End If

	With Frm1
		strVal = BIZ_SAVE_PGM_ID & "?txtMode="      & Parent.UID_M0002
        strVal = strVal     & "&txtYYYYMM="         & strYYYYMM							'☜: Query Key
        strVal = strVal     & "&txtModuleCd="       & .txtModuleCd.Value					'☜: Query Key
        strVal = strVal     & "&txtBizAreaCd="      & .txtBizAreaCd.Value					'☜: Query Key
        strVal = strVal     & "&txtOrgChangeId="    & .hOrgChangeid.Value					'☜: Query Key
        strVal = strVal     & "&txtDeptCd="         & .txtDeptCd.Value					'☜: Query Key
        strVal = strVal     & "&txtGLDt="			& UNIConvDate(Trim(.txtGLDt.Text)) 	'☜: Query Key

		Call ExtractDateFrom(Trim(.txtGLDt.Text),parent.gDateFormat,parent.gComDateType,strYear,strMonth,strDay)		
		strFirstDay = strYear + strMonth + strDay
		strFirstDay = UniDateAdd("m", 1, strFirstDay, parent.gServerDateFormat)
        strFirstDay = UniConvDateAToB(UNIGetFirstDay(strFirstDay,parent.gClientDateFormat),parent.gClientDateFormat,parent.gDateFormat)

        strVal = strVal     & "&txtRevGLDt="		& strFirstDay                    	'☜: Query Key
        strVal = strVal     & "&txtMaxRows="        & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex                 '☜: Next key tag
        
		'권한관리추가 start
		.txthAuthBizAreaCd.value =  lgAuthBizAreaCd
		.txthInternalCd.value =  lgInternalCd
		.txthSubInternalCd.value = lgSubInternalCd
		.txthAuthUsrID.value = lgAuthUsrID		
		'권한관리추가 end        
    End With												  

	Call RunMyBizASP(MyBizASP, strVal)													'☜: Run Biz Logic

    ExeReflect = True
End Function

'========================================================================================================
' Name : ExeReflectOk
' Desc : This function is data query and display
'========================================================================================================
Function ExeReflectOk() 
	Call DisplayMsgBox("183114", parent.VB_INFORMATION, "X", "X")
	Call fncQuery()
End Function
'========================================================================================================
' Name : ExeReflect
' Desc : This function is data query and display
'========================================================================================================
Function ExeCancle() 

    Dim strYYYYMM
    Dim strYear,strMonth,strDay
    Dim strVal
    Dim iRetCd,IntRetCD
    Dim strFirstDay
	
	ExeCancle = false
	err.Clear

    iRetCd = DisplayMsgBox("990008", parent.VB_YES_NO,"X","X")             'Will you destory previous data"
	If iRetCd = vbNo Then
		Exit Function
	End If
	    
	If LayerShowHide(1) = False Then
	   Exit Function
	End if
	
	Call ExtractDateFrom(frm1.txtYYYYMM.Text,frm1.txtYYYYMM.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)

    strYYYYMM = strYear & strMonth

	With Frm1
		strVal = BIZ_SAVE_PGM_ID & "?txtMode="      & Parent.UID_M0003
        strVal = strVal     & "&txtTempGLNo="       & .txtTempGLNo.Value				'☜: Query Key
        strVal = strVal     & "&txtGLNo="			& .txtGLNo.Value					'☜: Query Key
        strVal = strVal     & "&txtYYYYMM="         & strYYYYMM							'☜: Query Key
        strVal = strVal     & "&txtModuleCd="       & .txtModuleCd.Value				'☜: Query Key
        strVal = strVal     & "&txtBizAreaCd="      & .txtBizAreaCd.Value					'☜: Query Key
        strVal = strVal     & "&txtOrgChangeId="    & .hOrgChangeid.Value					'☜: Query Key        
        strVal = strVal     & "&txtDeptCd="         & .txtDeptCd.Value					'☜: Query Key
        strVal = strVal     & "&txtGLDt="			& UNIConvDate(Trim(.txtGLDt.Text)) 	'☜: Query Key        
        strFirstDay = UniConvDateAToB(UNIGetFirstDay(DateAdd("M",1,Trim(.txtGLDt.Text)),parent.gClientDateFormat),parent.gClientDateFormat,parent.gDateFormat)
        strVal = strVal     & "&txtRevGLDt="		& strFirstDay                    	'☜: Query Key
        strVal = strVal     & "&txtRevTempGLNo="    & .txtRevTempGLNo.Value				'☜: Query Key
        strVal = strVal     & "&txtRevGLNo="		& .txtRevGLNo.Value					'☜: Query Key
        strVal = strVal     & "&txtMaxRows="        & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex                 '☜: Next key tag
    End With												  '☜: Run Biz Logic

	Call RunMyBizASP(MyBizASP, strVal)		
	
    ExeCancle = True
    
End Function
'========================================================================================================
' Name : ExeReflectOk
' Desc : This function is data query and display
'========================================================================================================
Function ExeCancleOk() 
	Call DisplayMsgBox("183114", parent.VB_INFORMATION, "X", "X")
	Call fncQuery()
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                            '☆:
        Exit Function
    End If
    
    '-----------------------
    'Delete function call area
    '-----------------------
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()				
	 
	lgBlnFlgChgValue = False
	lgIntFlgMode = parent.OPMD_UMODE
    '-----------------------
    'Reset variables area
    '-----------------------
	If Trim(frm1.txtDeptCd.value) <> "" then	'전표생성되었다.
		Call ggoOper.SetReqAttr(frm1.txtDeptCd, "Q")'N:Required, Q:Protected, D:Default
		Call ggoOper.SetReqAttr(frm1.txtGLDt, "Q")
		frm1.btnReflect.disabled= True
		frm1.btnCancle.disabled= False	
	Else
		frm1.txtGLDt.text = UniConvDateAToB(UNIGetLastDay(frm1.txtYyyymm.Text,parent.gClientDateFormat) ,parent.gServerDateFormat,parent.gDateFormat)	
		Call ggoOper.SetReqAttr(frm1.txtDeptCd, "N")
		Call ggoOper.SetReqAttr(frm1.txtGLDt, "N")
		frm1.btnReflect.disabled= False
		frm1.btnCancle.disabled= True	
	End if                                           '☆: Developer must customize
	
	Call CurFormatNumSprSheet()
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables

	Call DisableToolBar(Parent.TBC_QUERY)
	
	If DBQuery = false Then
	    Call RestoreToolBar()
   	    Exit Function
	End If

End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'========================================================================================================
' Name : OpenCondAreaPopup()
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	With frm1
		If IsOpenPop = True Then Exit Function 

		Select Case iWhere
			Case 1
				arrParam(0) = "사업장 팝업"		    	<%' 팝업 명칭 %>
				arrParam(1) = "B_BIZ_AREA"					<%' TABLE 명칭 %>
				arrParam(2) = frm1.txtBizAreaCd.value			<%' Code Condition%>
				arrParam(3) = "" 		            		<%' Name Cindition%>
				arrParam(4) = ""							<%' Where Condition%>
				arrParam(5) = "사업장"			
	
				arrField(0) = "BIZ_AREA_CD"					<%' Field명(0)%>
				arrField(1) = "BIZ_AREA_NM"	     			<%' Field명(1)%>
    
				arrHeader(0) = "사업장코드"				<%' Header명(0)%>
				arrHeader(1) = "사업장명"				<%' Header명(1)%>

			Case 2
				arrParam(0) = "모듈구분팝업"										    ' 팝업 명칭 
				arrParam(1) = " b_minor "													' TABLE 명칭 
				arrParam(2) = frm1.txtModuleCd.value											' Code Condition
				arrParam(3) = ""															' Name Cindition
				arrParam(4) = " MAJOR_CD = " & FilterVar("A1045","''","S")
				arrParam(5) = "모듈구분"												' 조건필드의 라벨 명칭 

				arrField(0) = "MINOR_CD"													' Field명(0)
				arrField(1) = "MINOR_NM"													' Field명(1)
			 
				arrHeader(0) = "모듈구분"												' Header명(0)
				arrHeader(1) = "모듈구분명"												' Header명(1)

		End Select    
	End With
	
	IsOpenPop = True
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1
				frm1.txtBizAreaCd.focus
			Case 2
				frm1.txtModuleCd.focus
		End Select    
		Exit Function
	Else
		Call SetMajor(arrRet, iWhere)
	End If	

End Function

Function txtBizAreaCd_OnChange()
	frm1.hTxtBizAreaCd.value = frm1.txtBizAreaCd.value  

End Function


'======================================================================================================
'	Name : SetMajor()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================%>
Function SetMajor(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.txtBizAreaCd.focus
				.txtBizAreaCd.value = arrRet(0)
				.txtBizAreaNm.value = arrRet(1)
				.htxtBizAreaCd.value = arrRet(0)
			Case 2
				.txtModuleCd.focus
				.txtModuleCd.value = arrRet(0)
				.txtModuleName.value = arrRet(1)		
		End Select    
	End With
End Function


Function OpenDept(Byval strCode)
	Dim arrRet
	Dim arrParam(5),arrField(5),arrHeader(5)
	Dim strBizAreaCd

	' 선택한 사업장에 속한 부서만 PopUp
	strBizAreaCd = frm1.htxtBizAreaCd.value
	
	If Trim(strBizAreaCd) = "" then
		strBizAreaCd = "%"
	End If
		
	arrParam(0) = "부서코드팝업"			' 팝업 명칭 
	arrParam(1) = "B_ACCT_DEPT A, B_COST_CENTER B, B_BIZ_AREA C "    				' TABLE 명칭 
	arrParam(2) = frm1.txtDeptCd.value 						' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "A.ORG_CHANGE_ID = (select distinct org_change_id" & _
	              " from b_acct_dept where org_change_dt = ( select max(org_change_dt)" & _
	              " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))" & _
	              " and C.biz_area_cd LIKE " & FilterVar(strBizAreaCd, "''", "S")  & _
	              " AND B.cost_cd = A.cost_cd " & _
	              " AND C.biz_area_cd = B.biz_area_cd "
	arrParam(5) = "부서코드"													' 조건필드의 라벨 명칭 
			
	arrField(0) = "A.DEPT_CD"	     												' Field명(0)
	arrField(1) = "A.DEPT_NM"			    										' Field명(1)
	arrField(2) = "A.ORG_CHANGE_ID"			    									' Field명(1)
	arrField(3) = "C.BIZ_AREA_CD"			    									' Field명(2)
	arrField(4) = "C.BIZ_AREA_NM"			    									' Field명(3)
    
	arrHeader(0) = "부서코드"													' Header명(0)
	arrHeader(1) = "부서명"														' Header명(1)						
	arrHeader(2) = "조직개편아이디"												' Header명(1)						
	arrHeader(3) = "사업장코드"													' Header명(2)		
	arrHeader(4) = "사업장명"													' Header명(3)	

	' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    If arrRet(0) <> "" Then
		Call SetDept(arrRet)
	End If

	frm1.txtDeptCd.focus
	
End Function

'========================================================================================================= 
Function SetDept(Byval arrRet)
	With frm1
		.txtDeptCd.Value = arrRet(0)
		.txtDeptNm.Value = arrRet(1)
		.hOrgChangeid.Value = arrRet(2)

		 Call txtDeptCd_OnBlur()  
	End With
End Function       

'========================================================================================================= 
Sub txtDeptCd_OnBlur()
    Dim strSelect
	Dim strFrom
	Dim strWhere 	
    Dim IntRetCD 
	Dim arrVal1
	Dim arrVal2
	Dim ii
	Dim jj

	If Trim(frm1.txtGLDt.Text = "") Or Trim(frm1.txtDeptCd.value) = "" Then    
		Exit sub
    End If

    lgBlnFlgChgValue = True

	strSelect	=			 " dept_cd, org_change_id, internal_cd "    		
	strFrom		=			 " b_acct_dept(NOLOCK) "		
	strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.Value)), "''", "S") 
	strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
	strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
	strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UniConvDateToYYYYMMDD(frm1.txtGLDt.Text, parent.gDateFormat,""), "''", "S") & "))"			
		
	If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
		IntRetCD = DisplayMsgBox("124600","X","X","X")  
		frm1.txtDeptCd.Value = ""
		frm1.txtDeptNm.Value = ""
		frm1.hOrgChangeId.Value = ""
	Else 
		arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
		jj = Ubound(arrVal1,1)

		For ii = 0 to jj - 1
			arrVal2 = Split(arrVal1(ii), chr(11))			
			frm1.hOrgChangeId.Value = Trim(arrVal2(2))
		Next	
	End If
End Sub

Sub txtGLDt_onBlur()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
 
  	lgBlnFlgChgValue = True

	With frm1
		If LTrim(RTrim(.txtDeptCd.value)) <> "" and Trim(.txtGLDt.Text <> "") Then
			strSelect	=			 " Distinct org_change_id "    		
			strFrom		=			 " b_acct_dept(NOLOCK) "		
			strWhere	=			 " dept_Cd =  " & FilterVar(LTrim(RTrim(.txtDeptCd.value)), "''", "S") 
			strWhere	= strWhere & " and org_change_id = (select distinct org_change_id "			
			strWhere	= strWhere & " from b_acct_dept where org_change_dt = ( select max(org_change_dt)"
			strWhere	= strWhere & " from b_acct_dept where org_change_dt <= " & FilterVar(UNIConvDateToYYYYMMDD(.txtGLDt.Text, gDateFormat,""), "''", "S") & "))"			

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			If IntRetCD = False  Or Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
				.txtDeptCd.value = ""
				.txtDeptNm.value = ""
				.hOrgChangeId.value = ""
				.txtDeptCd.focus
			End If
		End If
	End With
End Sub

'=======================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtYyyymm_DblClick(Button) 
    If Button = 1 Then
        frm1.txtYyyymm.Action = 7
 		Call SetFocusToDocument("M")
		Frm1.txtYyyymm.Focus
   End If
End Sub

'=======================================================================================================
'   Event Name : txtYyyymm_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtYyyymm_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

'=======================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtYyyymm_Change() 
	frm1.txtGLDt.text = UniConvDateAToB(UNIGetLastDay(frm1.txtYyyymm.Text,parent.gClientDateFormat) ,parent.gServerDateFormat,parent.gDateFormat)	
End Sub

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub


'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
    		End If
    	End If
    End if
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")    

    gMouseClickStatus = "SPC" 

	Set gActiveSpdSheet = frm1.vspdData


    If frm1.vspdData.MaxRows <= 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)	
	
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
    
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
	
	If Row<=0 then
		Exit Sub
	End If
	If Frm1.vspdData.MaxRows =0 then
		Exit Sub
	End if
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    'If Col <= C_SNm Or NewCol <= C_SNm Then
    '    Cancel = True
    '    Exit Sub
    'End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()

	Call InitData()
End Sub

'========================================================================================================
'   Event Name : vspdData_LeaveCell
'   Event Desc : Focus 이동 
'========================================================================================================
Sub vspdData_LeaveCell(Col, Row, NewCol, NewRow, Cancel)
    
 '   frm1.vspdData.OperationMode = 3             
	
End Sub



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>


<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>작업년월</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript>ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtYyyymm" CLASS=FPDTYYYYMMDD tag="12X1" ALT="작업년월" Title="FPDATETIME"></OBJECT>')</script></TD>
								
			            		<TD CLASS="TD5" NOWRAP>모듈구분</TD>
			            		<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtModuleCd" SIZE=10 MAXLENGTH=20 tag="12XXXU" ALT="모듈구분"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnModuleCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup(2)">
									<INPUT TYPE="Text" NAME="txtModuleName" SiZE=22 MAXLENGTH=50 tag="14XXXU" ALT="모듈구분명">
			            		</TD>
			            	</TR>
			            	<TR>
			            		<TD CLASS="TD5" NOWRAP>사업장</TD>
			            		<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtBizAreaCd" SIZE=10 MAXLENGTH=20 tag="121XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMajorCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondAreaPopup(1)">
									<INPUT TYPE="Text" NAME="txtBizAreaNm" SiZE=22 MAXLENGTH=50 tag="14XXXU" ALT="사업장명">
			            		</TD>
			            		<TD CLASS="TD5" NOWRAP></TD>
			            		<TD CLASS="TD6" NOWRAP></TD>
			            	</TR> 
			            </TABLE>
			    	    </FIELDSET>
			        </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD WIDTH=90% HEIGHT=40>
						<TABLE <%=LR_SPACE_TYPE_60%>>					
							<TR>
								<TD CLASS=TD5 NOWRAP>전표일자</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript>ExternalWrite('<OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtGLDt" CLASS=FPDTYYYYMMDD tag="23X1" ALT="전표일자" Title="FPDATETIME"></OBJECT>')</script></TD>
<!--								<TD CLASS=TD5 NOWRAP></TD>								
								<TD CLASS=TD6 NOWRAP></TD>-->
								<TD CLASS=TD5 NOWRAP>부서</TD>								
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" MAXLENGTH="10" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(frm1.txtDeptCd.Value)">&nbsp;
													 <INPUT NAME="txtDeptNm" ALT="부서명"   MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN: left" tag="24X"></TD>
													 <INPUT NAME="txtInternalCd" ALT="내부부서코드" TYPE="hidden" SiZE=7 MAXLENGTH=30  tag="24"  TABINDEX="-1">
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTempGLNo" ALT="결의전표번호" MAXLENGTH="30" SIZE="20" tag="24X" >
								                     <INPUT NAME="txtGLNo" ALT="회계전표번호" MAXLENGTH="30" SIZE="20" tag="24X" ></TD>
								<TD CLASS=TD5 NOWRAP>역분개전표번호</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRevTempGLNo" ALT="결의전표번호" MAXLENGTH="30" SIZE="20" tag="24X" >
								                     <INPUT NAME="txtRevGLNo" ALT="회계전표번호" MAXLENGTH="30" SIZE="20" tag="24X" ></TD>
							</TR>							
			            </TABLE>
			        </TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT="100%" NAME=vspdData width="100%" tag="23" TITLE=SPREAD id=OBJECT12> <PARAM NAME=MaxCols VALUE=0><PARAM NAME=MaxRows VALUE=0></OBJECT>')</script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD NOWRAP  <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=20>
		<TD NOWRAP >
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD NOWRAP  WIDTH=10>&nbsp;</TD>
					<TD NOWRAP ><BUTTON NAME="btnReflect" CLASS="CLSSBTN" onclick="ExeReflect()" >전표생성</BUTTON>&nbsp;
					            <BUTTON NAME="btnCancle" CLASS="CLSSBTN" onclick="ExeCancle()">전표취소</BUTTON></TD>
                    <TD NOWRAP  WIDTH=*>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="24" TABINDEX="-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hOrgChangeid"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hTxtBizAreaCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=hidden NAME="txthAuthBizAreaCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthSubInternalCd"	tag="24" Tabindex="-1">
<INPUT TYPE=hidden NAME="txthAuthUsrID"		tag="24" Tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>

