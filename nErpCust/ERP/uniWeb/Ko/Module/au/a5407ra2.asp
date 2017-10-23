<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>


<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : a5117ma1
'*  4. Program Name         : 결의전표조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : +
'*  7. Modified date(First) : 2000/11/27
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ahn Hae Jin
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'*                            
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">		</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">			</SCRIPT>
<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
	

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
top.document.title = "미결연결팝업"
Const BIZ_PGM_ID 		= "a5407rb2.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 20                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 
Const C_MaxKey          = 21					                          '☆: SpreadSheet의 키의 갯수 

Const C_AcctCd = 1
Const C_DocCur = 2
Const C_MgntCd1 = 3
Const C_MgntCd2 = 4

Const C_MgntCD1_Grid		= 1
Const C_MgntCD2_Grid		= 2
Const C_GlNo_Grid			= 3
Const C_GlDt_Grid			= 4
Const C_GlDesc_Grid			= 5
Const C_OpenAmt_Grid		= 6
Const C_Temp_ClsAmt_Grid	= 7
Const C_BalAmt_Grid			= 8
Const C_GlSeq_Grid			= 9
Const C_AcctCD_Grid			= 10
Const C_AcctNm_Grid			= 11
Const C_DrCrFg_Grid			= 12
Const C_DrCrNm_Grid			= 13
Const C_XchRate_Grid		= 14
Const C_MgntFg_Grid			= 15
Const C_DocCur_Grid			= 16

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          

Dim lgSelectList                                         
Dim lgSelectListDT                                       

Dim lgSortFieldNm                                        
Dim lgSortFieldCD                                         

Dim lgMaxFieldCount

Dim lgPopUpR                                              

Dim lgKeyPos                                              
Dim lgKeyPosVal												
Dim lgCookValue 

Dim lgSaveRow 
Dim IsOpenPop 

Dim lgAuthorityFlag

Dim lgArrReturn

Dim lgArrParent
Dim lgGlNoSeq
Dim lgDocCur
Dim lgtodate
Dim arrParam

Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

lgArrParent = window.dialogArguments
Set PopupParent = lgArrParent(0)
lgGlNoSeq = lgArrParent(1)
lgtodate = UNIConvDateToYYYYMMDD(lgArrParent(2),PopupParent.gDateFormat,"")
arrParam  = lgArrParent(3)

ReDim lgArrReturn(0,0)
Self.Returnvalue = lgArrReturn	

'------ Set Parameters from Parent ASP -----------------------------------------------------------------------

Dim BaseDate,LastDate,FirstDate
                                                 
   BaseDate     = "<%=GetSvrDate%>"                                                           'Get DB Server Date
'  BaseDate     = Date(You must not code like this!!!!)                                       'Get AP Server Date

   LastDate     = UNIGetLastDay (BaseDate,PopupParent.gServerDateFormat)                                  'Last  day of this month
   FirstDate    = UNIGetFirstDay(BaseDate,PopupParent.gServerDateFormat)                                  'First day of this month


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0

 	' 권한관리 추가 
	If UBound(arrParam) > 5 Then
		lgAuthBizAreaCd		= arrParam(5)
		lgInternalCd		= arrParam(6)
		lgSubInternalCd		= arrParam(7)
		lgAuthUsrID			= arrParam(8)
	End If
 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()

	Dim strYear, strMonth, strDay
	Dim StartDate,EndDate

	Call ExtractDateFrom(lgtodate, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)

	StartDate= UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, "01")		'☆: 초기화면에 뿌려지는 시작 날짜 
	EndDate= UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)		'☆: 초기화면에 뿌려지는 마지막 날짜 

    frm1.txtFromDt.text	= StartDate
	frm1.txtToDt.Text	= EndDate

End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "*","NOCOOKIE","RA") %>                                '☆: 
	<% Call LoadBNumericFormatA("Q", "*","NOCOOKIE","RA") %>
End Sub



'========================================================================================================
'	Name : CookiePage()
'	Description : JUMP시 Load화면으로 조건부로 Value
'========================================================================================================
Function CookiePage(ByVal Kubun)

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
		Dim strTemp, arrVal

	Const CookieSplit = 4877						

	If Kubun = 0 Then                                              ' Called Area
       strTemp = ReadCookie(CookieSplit)

       If strTemp = "" then Exit Function

       arrVal = Split(strTemp, PopupParent.gRowSep)


       WriteCookie CookieSplit , ""
	
	ElseIf Kubun = 1 then                                         ' If you want to call
		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		WriteCookie CookieSplit , lgCookValue		
		Call PgmJump(BIZ_PGM_JUMP_ID2)
	End IF

	
End Function

'=============================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'============================================================================================================
Sub InitComboBox()	
	Err.clear	
	 
End Sub



'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    
	frm1.vspdData.OperationMode = 5
	Call SetZAdoSpreadSheet("A5407RA2", "S", "A", "V20030324", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
	Call SetSpreadLock() 
End Sub

'========================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================================
Sub SetSpreadLock()
	  ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub




'++++++++++++++++++++++++++++++++++++++++++  2.3 개발자 정의 함수  ++++++++++++++++++++++++++++++++++++++
'+	개발자 정의 Function, Procedure																		+
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'==========================================  2.3.1 OkClick()  ===========================================
'=	Name : OkClick()																					=
'=	Description : Return Array to Opener Window when OK button click									=
'========================================================================================================


Function OKClick()

	Dim intColCnt, intRowCnt, intInsRow, ii
	if frm1.vspdData.SelModeSelCount > 0 Then 			

		intInsRow = 0

		Redim lgArrReturn(frm1.vspdData.SelModeSelCount -1, C_MaxKey)  
		For intRowCnt = 0 To frm1.vspdData.MaxRows

			frm1.vspdData.Row = intRowCnt + 1
			If frm1.vspdData.SelModeSelected Then
				For ii = 0 to GetKeyPos("A",C_MaxKey) - 1
					frm1.vspdData.Col	= GetKeyPos("A",ii + 1)
					lgArrReturn(intInsRow,ii)		= frm1.vspdData.Text
				Next
				intInsRow = intInsRow + 1
			End IF
		Next
		
	End if
	Self.Returnvalue = lgArrReturn
	Self.Close()
					
End Function




'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'=	Description : Return Array to Opener Window for Cancel button click 								=
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function
	
	
'===========================================================================
' Function Name : OpenOrderBy
' Function Desc : OpenOrderBy Reference Popup
'===========================================================================
Function OpenSortPopup()
   
   	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()
   End If

End Function
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    
    
    Call ggoOper.LockField(Document, "N")
	'call SetAuthorityFlag	'권한관리 

	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()

	If lgDocCur <> "" then

		frm1.txtDocCur.value = lgDocCur
		Call ggoOper.SetReqAttr(frm1.txtDocCur,	"Q")
	End If
	frm1.txtDocCur.value = "KRW"

    frm1.txtFromDt.focus
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
End Sub
'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub


'**************************  3.2 HTML Form Element & Object Event처리  **********************************
'	Document의 TAG에서 발생 하는 Event 처리	
'	Event의 경우 아래에 기술한 Event이외의 사용을 자제하며 필요시 추가 가능하나 
'	Event간 충돌을 고려하여 작성한다.
'********************************************************************************************************* 

 '******************************  3.2.1 Object Tag 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'********************************************************************************************************* 

Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFromDt.Focus
    End If
End Sub

Sub txtFromDt_Change() 
    lgBlnFlgChgValue = True
End Sub

Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.Focus
    End If
End Sub

Sub txtToDt_Change() 
    lgBlnFlgChgValue = True
End Sub

Sub txtFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub


'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 

    FncQuery = False                                                
    Err.Clear                                                   

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    If CompareDateByFormat(frm1.txtFromDt.text,frm1.txtToDt.text,frm1.txtFromDt.Alt,frm1.txtToDt.Alt, _
                        "970025",frm1.txtFromDt.UserDefinedFormat,PopupParent.gComDateType,True) = False Then			
		Exit Function
    End If


    '-----------------------
    'Query function call area
    '-----------------------
    
    IF DbQuery	 = False Then															'☜: Query db data
       Exit Function
    End IF
       
    FncQuery = True												

End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function


'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status    
   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCancel
' Desc : This function is called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
    FncInsertRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncInsertRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : This function is called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrint
' Desc : This function is called by MainPrint in Common.vbs
'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncPrev
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncNext
' Desc : This function is called by MainPrev in Common.vbs
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncExcel
' Desc : This function is called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(PopupParent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(PopupParent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function

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


'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    Err.Clear                                                       
    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1
        strVal = BIZ_PGM_ID
        
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
        If lgIntFlgMode = PopupParent.OPMD_CMODE Then   ' This means that it is first search
			strVal = strVal & "?txtMode="		& PopupParent.UID_M0001	
			strVal = strVal & "&txtFromDt="		& UNIConvDateToYYYYMMDD(Trim(.txtFromDt.text),PopupParent.gDateFormat,"")	
			strVal = strVal & "&txtToDt="		& UNIConvDateToYYYYMMDD(Trim(.txtToDt.text),PopupParent.gDateFormat,"")								
			strVal = strVal & "&txtDocCur="		& UCase(Trim(.txtDocCur.value))				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtDeptCd="		& UCase(Trim(.txtDeptCd.value))
			strVal = strVal & "&txtCardCoCd="	& UCase(Trim(.txtCardCoCd.value))
			strVal = strVal & "&txtCardNo="		& UCase(Trim(.txtCardNo.value))
			strVal = strVal & "&txtMaxRows="	& frm1.vspdData.MaxRows
        Else
            strVal = strVal & "?txtMode="		& PopupParent.UID_M0001	
			strVal = strVal & "&txtFromDt="		& UNIConvDateToYYYYMMDD(Trim(.hFromDt.value),PopupParent.gDateFormat,"")		
			strVal = strVal & "&txtToDt="		& UNIConvDateToYYYYMMDD(Trim(.hToDt.value),PopupParent.gDateFormat,"")		
			strVal = strVal & "&txtDocCur="		& UCase(Trim(.hDocCur.value))				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtDeptCd="		& UCase(Trim(.hDeptCd.value))
			strVal = strVal & "&txtCardCoCd="	& UCase(Trim(.hCardCoCd.value))
			strVal = strVal & "&txtCardNo="		& UCase(Trim(.hCardNo.value))
			strVal = strVal & "&txtMaxRows="	& .vspdData.MaxRows
        End If  
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&lgPageNo="       & lgPageNo         
        strVal = strVal & "&lgMaxCount="     & C_SHEETMAXROWS_D
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		strVal = strVal & "&lgAuthorityFlag="   & EnCoding(lgAuthorityFlag)            '권한관리 추가		
		
		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 
		
        
    End With
    
    frm1.txtGlNoSeq.value =lgGlNoSeq
    frm1.lgSelectListDT.value = GetSQLSelectListDataType("A")
    'msgbox strVal
     Call ExecMyBizASP(frm1, strVal)	
    DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncQuery에 있는것을 옮김 
'========================================================================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    'lgSaveRow        = 1
	'CALL vspdData_Click(1, 1)
End Function


'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'------------------------------------------  OpenItemInfo()  ---------------------------------------------
'	Name : OpenItemInfo()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 

Function OpenPopUp(Byval strCode, Byval iWhere)


Dim arrRet
Dim arrParam(5), arrField(6), arrHeader(6)
Dim arrStrRet				'권한관리 추가  
Dim IntRetCD, IntRetCD1

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.hOrgChangeId.value = PopupParent.gChangeOrgId

	Select Case iWhere
	CASE 0
		arrParam(0) = "카드사팝업"									' 팝업 명칭 
		arrParam(1) = "b_card_co A"						' TABLE 명칭 
		arrParam(2) = strCode													' Code Condition
		arrParam(3) = ""														' Name Cindition
		arrParam(4) = ""														' Where Condition			
		arrParam(5) = frm1.txtCardCoCd.Alt										' 조건필드의 라벨 명칭 

		arrField(0) = "A.CARD_CO_CD"						' Field명(0)
		arrField(1) = "A.CARD_CO_NM"						' Field명(1)
   
		arrHeader(0) = frm1.txtCardCoCd.Alt					' Header명(0)
		arrHeader(1) = frm1.txtCardCoNm.Alt					' Header명(1)
	Case 1
		arrParam(0) = "카드번호팝업"								' 팝업 명칭 
		arrParam(1) = "B_CREDIT_CARD"	 									' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = ""												' Where Condition
		arrParam(5) = frm1.txtCardNo.Alt								' 조건필드의 라벨 명칭 

		arrField(0) = "CREDIT_NO"										' Field명(0)
		arrField(1) = "USE_USER_ID"
		arrField(2) = "CREDIT_NM"										' Field명(0)

		arrHeader(0) = frm1.txtCardNo.Alt									' Header명(0)
		arrHeader(1) = "사용자"
		arrHeader(2) = "카드명"									' Header명(1)
	CASE 2
		arrParam(0) = "부서팝업"								' 팝업 명칭 
		arrParam(1) = "b_acct_dept"	 									' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = ""												' Where Condition
		arrParam(5) = frm1.txtDeptCd.Alt								' 조건필드의 라벨 명칭 

		arrField(0) = "Dept_Cd"										' Field명(0)
		arrField(1) = "Dept_Nm"									' Field명(1)

		arrHeader(0) = frm1.txtDeptCd.Alt									' Header명(0)
		arrHeader(1) = frm1.txtDeptNm.Alt									' Header명(1)
	Case 6
		arrParam(0) = "거래통화팝업"								' 팝업 명칭 
		arrParam(1) = "B_CURRENCY"	 									' TABLE 명칭 
		arrParam(2) = strCode											' Code Condition
		arrParam(3) = ""												' Name Cindition
		arrParam(4) = ""												' Where Condition
		arrParam(5) = frm1.txtDocCur.Alt								' 조건필드의 라벨 명칭 

		arrField(0) = "CURRENCY"										' Field명(0)
		arrField(1) = "CURRENCY_DESC"									' Field명(1)

		arrHeader(0) = frm1.txtDocCur.Alt									' Header명(0)
		arrHeader(1) = "통화코드명"									' Header명(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

'------------------------------------------  OpenDeptOrgPopup()  ---------------------------------------
'	Name : OpenDeptOrgPopup()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenDeptOrgPopup()
	Dim arrRet
	Dim arrParam(8)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtFromDt.text								'  Code Condition
   	arrParam(1) = frm1.txtToDt.Text
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "F"									' 결의일자 상태 Condition  
	
	' 권한관리 추가 
	arrParam(5) = lgAuthBizAreaCd
	arrParam(6) = lgInternalCd
	arrParam(7) = lgSubInternalCd
	arrParam(8) = lgAuthUsrID	
	
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(popupparent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDeptCd.focus
		Exit Function
	Else
		
		frm1.txtDeptCd.value =arrRet(0)
		frm1.txtDeptNm.value =arrRet(1)
	End If	
End Function


'------------------------------------------  SetCostCenterInfo()  ----------------------------------------
'	Name : SetAcct()
'	Description : Account Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere				
			Case 0
				.txtCardCoCd.Value = arrRet(0)
				.txtCardCoNm.value = arrRet(1)
			Case 1
				.txtCardNo.value = arrRet(0)
				.txtCardNm.value = arrRet(2)
			Case 2
				.txtDeptCd.value =  arrRet(0)
				.txtDeptNm.value =  arrRet(1)		
			Case 6
				.txtDocCur.value =  arrRet(0)		
		End Select

		'lgBlnFlgChgValue = True
	End With
End Function

'========================================================================================================
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'========================================================================================================
Function OpenGroupPopup()

	Dim arrRet
	Dim arrParam
	Dim TInf(5)
	Dim ii
	
	On Error Resume Next
	
	ReDim arrParam(PopupParent.C_MaxSelList * 2 - 1 )

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    TInf(0) = PopupParent.gMethodText
  
	For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
      arrParam(ii + 0 ) = lgPopUpR(ii / 2  , 0)
      arrParam(ii + 1 ) = lgPopUpR(ii / 2  , 1)
    Next  
      
  
	arrRet = window.showModalDialog("../../ComAsp/ADOGrpSortPopup.asp",Array(lgSortFieldCD,lgSortFieldNm,arrParam,TInf),"dialogWidth=420px; dialogHeight=250px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "0" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	
	   For ii = 0 to PopupParent.C_MaxSelList * 2 - 1 Step 2
           lgPopUpR(ii / 2 ,0) = arrRet(ii + 1)  
           lgPopUpR(ii / 2 ,1) = arrRet(ii + 2)
       Next    
	   
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	If Row < 1 Then Exit Sub

	'frm1.vspdData.Row = Row
	'lsPoNo=frm1.vspdData.Text
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)

End Sub
	
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then		   
    	If lgPageNo <> "" Then    	
           Call DbQuery()           
    	End If
    End If
    
End Sub

'========================================================================================================
'   Event Name : fpdtFromEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtFromEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtFromEnterDt.Focus
	End If
End Sub
'========================================================================================================
'   Event Name : fpdtToEnterDt
'   Event Desc : Date OCX Double Click
'========================================================================================================
Sub fpdtToEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtToEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtToEnterDt.Focus
	End If
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub fpdtFromEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'========================================================================================================
Sub fpdtToEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub


'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					
					<TR>
						<TD CLASS=TD5 NOWRAP>기간</TD>
						<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime1 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtFromDt CLASSID=<%=gCLSIDFPDT%> ALT="시작일자" tag="11"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
											 <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtToDt CLASSID=<%=gCLSIDFPDT%> ALT="종료일자" tag="14XXXU"></OBJECT>');</SCRIPT></TD>
						<TD CLASS=TD5 NOWRAP>거래통화</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.Value, 6)"></TD>
					</TR>
					<TR>				
						<TD CLASS="TD5" NOWRAP>부서</TD>
						<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDeptCd" ALT="부서코드" Size= "10" MAXLENGTH="10"  tag="11X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenDeptOrgPopup()">
											   <INPUT NAME="txtDeptNm" ALT="부서명" SIZE = "20" STYLE="TEXT-ALIGN: left" tag="14X"></TD>
                        <TD CLASS=TD5 NOWRAP>카드사</TD>
                        <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCardCoCd"  SIZE="10" MAXLENGTH="10" TAG="11xxxU" ALT="카드사"><IMG SRC="../../image/btnPopup.gif" NAME="txtCardCoCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtCardCoCd.value,0)">
											<INPUT TYPE=TEXT NAME="txtCardCoNm"  SIZE=20   TAG="14xxxU" ALT="카드사명"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>
						<TD CLASS=TD5 NOWRAP>카드번호</TD>
                        <TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCardNo"  SIZE=20 MAXLENGTH=20 TAG="11XXXU" ALT="신용카드번호"><IMG SRC="../../image/btnPopup.gif" NAME="txtCardNo" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtCardNo.value,1)">
                        <INPUT TYPE=TEXT NAME="txtCardNm"  SIZE=20   TAG="14xxxU" ALT="카드명"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR HEIGHT=100%>
					<TD WIDTH=100%>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vspdData tag="2"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"></OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>&nbsp;<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD> 
		<!-- <TD WIDTH=100% HEIGHT=80%><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=80% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>  -->

	</TR>
</TABLE>

<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24">
<INPUT TYPE=hidden NAME="hFromDt"   tag="24">
<INPUT TYPE=hidden NAME="hToDt"   tag="24">
<INPUT TYPE=hidden NAME="hDocCur"   tag="24">
<INPUT TYPE=hidden NAME="hDeptCd"   tag="24">
<INPUT TYPE=hidden NAME="hCardCoCd"   tag="24">
<INPUT TYPE=hidden NAME="hCardNo"   tag="24">
<INPUT TYPE=hidden NAME="lgSelectListDT"   tag="24">
<INPUT TYPE=hidden NAME="txtGlNoSeq"   tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


