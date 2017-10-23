<%@ LANGUAGE="VBSCRIPT" %>
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
Const BIZ_PGM_ID 		= "a5403rb2_ko441.asp"                              '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_MaxKey          = 24					                          '☆: SpreadSheet의 키의 갯수 

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
Dim Strflag
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

Dim lgArrParam

Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 >>AIR

lgArrParent = window.dialogArguments
Set PopupParent = lgArrParent(0)
lgGlNoSeq = lgArrParent(1)
lgDocCur = lgArrParent(2)
lgtodate = UNIConvDateToYYYYMMDD(lgArrParent(3),PopupParent.gDateFormat,"")
lgArrParam = lgArrParent(4)
ReDim lgArrReturn(0,0)
Self.Returnvalue = lgArrReturn	



'------ Set Parameters from Parent ASP -----------------------------------------------------------------------

Dim BaseDate,LastDate,FirstDate,FromDateOfDB,ToDateOfDB

                                                 
   BaseDate     = "<%=GetSvrDate%>"                                                           'Get DB Server Date

   LastDate     = UNIGetLastDay (BaseDate,PopupParent.gServerDateFormat)                                  'Last  day of this month
   FirstDate    = UNIGetFirstDay(BaseDate,PopupParent.gServerDateFormat)                                  'First day of this month

   FromDateOfDB = UNIDateAdd("yyyy", -410, BaseDate,PopupParent.gServerDateFormat)
   ToDateOfDB   = UNIDateAdd("yyyy",  410, BaseDate,PopupParent.gServerDateFormat)
 
   FromDateOfDB  = UniConvDateAToB(FromDateOfDB ,PopupParent.gServerDateFormat,PopupParent.gDateFormat)               'Convert DB date type to Company
   ToDateOfDB    = UniConvDateAToB(ToDateOfDB   ,PopupParent.gServerDateFormat,PopupParent.gDateFormat)               'Convert DB date type to Company



'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0

	If Not isNull(lgArrParam(5)) And lgArrParam(5) <> "" Then
		lgAuthBizAreaCd	= lgArrParam(5)
	End If	
 
End Sub


'========================================================================================================
Sub SetDefaultVal()

	Dim strYear, strMonth, strDay
	Dim StartDate,EndDate

	Call	ExtractDateFrom(lgtodate, PopupParent.gServerDateFormat, PopupParent.gServerDateType, strYear, strMonth, strDay)	

	StartDate= UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, "01")		'☆: 초기화면에 뿌려지는 시작 날짜 

	EndDate= UniConvYYYYMMDDToDate(PopupParent.gDateFormat, strYear, strMonth, strDay)		'☆: 초기화면에 뿌려지는 마지막 날짜 

    frm1.txtFromDt.text		= StartDate
	frm1.txtToDt.Text		= EndDate

End Sub


'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "A","NOCOOKIE","RA") %> 
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "RA") %>                               '☆: 

End Sub



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


'============================================================================================================
Sub InitComboBox()	
	Err.clear	
	 
End Sub



'========================================================================================================
Sub InitSpreadSheet()
    
	frm1.vspdData.OperationMode = 5
	Call SetZAdoSpreadSheet("A5403RA2_KO441", "S", "A", "V20041210", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
	Call SetSpreadLock() 
End Sub


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
	Dim intColCnt, intRowCnt, intInsRow
	If frm1.vspdData.SelModeSelCount > 0 Then 			
		intInsRow = 0

		Redim lgArrReturn(frm1.vspdData.SelModeSelCount -1, 16)  
		For intRowCnt = 0 To frm1.vspdData.MaxRows
			frm1.vspdData.Row = intRowCnt + 1
			If frm1.vspdData.SelModeSelected Then
				frm1.vspdData.Col	= GetKeyPos("A",7)	'C_GlNo_Grid
				lgArrReturn(intInsRow,0)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",13)	'C_GlSeq_Grid
				lgArrReturn(intInsRow,1)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",14)	'C_AcctCD_Grid
				lgArrReturn(intInsRow,2)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",15)	'C_AcctNm_Grid
				lgArrReturn(intInsRow,3)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",16)	'C_DrCrFg_Grid
				lgArrReturn(intInsRow,4)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",17)	'C_DrCrNm_Grid
				lgArrReturn(intInsRow,5)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",12)	'C_BalAmt_Grid
				lgArrReturn(intInsRow,6)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",9)	'C_GlDesc_Grid
				lgArrReturn(intInsRow,7)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",18)	'C_XchRate_Grid
				lgArrReturn(intInsRow,8)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",19)	'C_MgntFg_Grid
				lgArrReturn(intInsRow,9)		= frm1.vspdData.Text
				'frm1.vspdData.Col	= GetKeyPos("A",11)	'C_DocCur_Grid
				lgArrReturn(intInsRow,10)		= frm1.htxtDocCur.value
				frm1.vspdData.Col	= GetKeyPos("A",8)	'C_GlDT_Grid
				lgArrReturn(intInsRow,11)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",20)	'C_deptcd
				lgArrReturn(intInsRow,12)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",21)	'C_deptnm
				lgArrReturn(intInsRow,13)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",22)	'C_internalcd
				lgArrReturn(intInsRow,14)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",23)	'C_costcd
				lgArrReturn(intInsRow,15)		= frm1.vspdData.Text
				frm1.vspdData.Col	= GetKeyPos("A",24)	'C_orgchangeid
				lgArrReturn(intInsRow,16)		= frm1.vspdData.Text
			
				intInsRow = intInsRow + 1
			End If
		Next
	End If
	
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
Sub Form_Load()
    Call LoadInfTB19029														

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
     
    Call ggoOper.LockField(Document, "N")

	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()

	If lgDocCur <> "" then
		frm1.txtDocCur.value = lgDocCur
		Call ggoOper.SetReqAttr(frm1.txtDocCur,		"Q")
	End If
	
	Call ggoOper.SetReqAttr(frm1.txtMgntCd1,		"Q")
	Call ggoOper.SetReqAttr(frm1.txtMgntCd2,		"Q")
	Call ggoOper.SetReqAttr(frm1.txtMgntCd2,		"Q")
	Call ggoOper.SetReqAttr(frm1.txtMgntCd1Nm,		"Q")
	Call ggoOper.SetReqAttr(frm1.txtMgntCd2Nm,		"Q")
	Strflag = "2"
    frm1.txtFromDt.focus
    Set gActiveElement = document.activeElement		
End Sub

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

Sub txtToGl_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub


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

    If frm1.txtAcctCd.value = "" Then
		frm1.txtAcctNm.value = ""
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
Function FncCancel() 
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
End Function


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
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function


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
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncExport(PopupParent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call Parent.FncFind(PopupParent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
End Function


'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


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
			strVal = strVal & "&txtAcctCd="		& Trim(.txtAcctCd.value)
			strVal = strVal & "&txtFromDt="		& UNIConvDateToYYYYMMDD(Trim(.txtFromDt.text),PopupParent.gDateFormat,"")
			strVal = strVal & "&txtToDt="		& UNIConvDateToYYYYMMDD(Trim(.txtToDt.text),PopupParent.gDateFormat,"")								
			strVal = strVal & "&txtDocCur="		& UCase(Trim(.txtDocCur.value))				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtBizCd="		& Trim(.txtBizCd.value)		'>>air
		    strVal = strVal & "&txtGlNo="		& Trim(.txtGlNo.value)
			strVal = strVal & "&txtMgntCd1="	& Trim(.txtMgntCd1.value)
			strVal = strVal & "&txtMgntCd2="	& Trim(.txtMgntCd2.value)
			'strVal = strVal & "&txtGlNoSeq="	& lgGlNoSeq			
			strVal = strVal & "&txtMaxRows="	& frm1.vspdData.MaxRows
			strVal = strVal & "&txtAcctCd_Alt=" & Trim(.txtAcctCd.Alt)
			strVal = strVal & "&txtDocCur_Alt=" & Trim(.txtDocCur.Alt)
			strVal = strVal & "&txtGlNo_Alt="   & Trim(.txtGlNo.Alt)
			strVal = strVal & "&txtMgntCd1_Alt=" & Trim(.txtMgntCd1.Alt)
			strVal = strVal & "&txtMgntCd2_Alt=" & Trim(.txtMgntCd2.Alt)
			strVal = strVal & "&txtFrDueDt="	&  UNIConvDateToYYYYMMDD(Trim(.txtFrDueDt.text),PopupParent.gDateFormat,"")
			strVal = strVal & "&txtToDueDt="	& UNIConvDateToYYYYMMDD(Trim(.txtToDueDt.text),PopupParent.gDateFormat,"")		
        Else
            strVal = strVal & "?txtMode="		& PopupParent.UID_M0001	
			strVal = strVal & "&txtAcctCd="		& Trim(.htxtAcctCd.value)
			strVal = strVal & "&txtFromDt="		& UNIConvDateToYYYYMMDD(Trim(.htxtFromDt.value),PopupParent.gDateFormat,"")
			strVal = strVal & "&txtToDt="		& UNIConvDateToYYYYMMDD(Trim(.htxtToDt.value),PopupParent.gDateFormat,"")								
			strVal = strVal & "&txtDocCur="		& UCase(Trim(.htxtDocCur.value))				'☆: 조회 조건 데이타 
			strVal = strVal & "&txtBizCd="		& Trim(.htxtBizCd.value)	'>>air
		    strVal = strVal & "&txtGlNo="		& Trim(.htxtGlNo.value)
			strVal = strVal & "&txtMgntCd1="	& Trim(.htxtMgntCd1.value)
			strVal = strVal & "&txtMgntCd2="	& Trim(.htxtMgntCd2.value)
			'strVal = strVal & "&txtGlNoSeq="	& lgGlNoSeq
			strVal = strVal & "&txtMaxRows="	& .vspdData.MaxRows
			strVal = strVal & "&txtAcctCd_Alt=" & Trim(.txtAcctCd.Alt)
			strVal = strVal & "&txtDocCur_Alt=" & Trim(.txtDocCur.Alt)
			strVal = strVal & "&txtGlNo_Alt="   & Trim(.txtGlNo.Alt)
			strVal = strVal & "&txtMgntCd1_Alt=" & Trim(.txtMgntCd1.Alt)
			strVal = strVal & "&txtMgntCd2_Alt=" & Trim(.txtMgntCd2.Alt)			
			strVal = strVal & "&txtFrDueDt="	& UNIConvDateToYYYYMMDD(Trim(.txtFrDueDt.text),PopupParent.gDateFormat,"")' UniConvDateAToB(Trim(.txtFrDueDt.text),popupparent.gDateFormat,popupparent.gServerDateFormat)
			strVal = strVal & "&txtToDueDt="	& UNIConvDateToYYYYMMDD(Trim(.txtToDueDt.text),PopupParent.gDateFormat,"") 'UniConvDateAToB(Trim(.txtToDueDt.text),popupparent.gDateFormat,popupparent.gServerDateFormat)			
        End If  
    '--------- Developer Coding Part (End) ------------------------------------------------------------
        strVal = strVal & "&Strflag="        & Strflag
        strVal = strVal & "&lgPageNo="       & lgPageNo         
	'	 strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgSelectList=" & replace(GetSQLSelectList("A"),"+","%2B")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")         
		 'strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		 strVal = strVal & "&lgAuthorityFlag="   & EnCoding(lgAuthorityFlag)            '권한관리 추가		
		 
	'Call RunMyBizASP(MyBizASP, strVal)							
	
    End With
    'msgbox strVal
 frm1.txtGlNoSeq.value = lgGlNoSeq	         
    frm1.lgSelectListDT.value= GetSQLSelectListDataType("A")


    Call ExecMyBizASP(frm1, strVal)
    DbQuery = True

End Function


'========================================================================================
Function DbQueryOk()												
    
	lgBlnFlgChgValue = False
	lgIntFlgMode     = PopupParent.OPMD_UMODE	
    call txtDocCur_OnChange()											'⊙: Indicates that current mode is Update mode
 
End Function


'--------------------------------------------------------------------------------------------------------- 

Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim arrStrRet				'권한관리 추가  
	Dim IntRetCD, IntRetCD1
	Dim strFrom, strWhere, strFrom1, strWhere1
	Dim arrVal, arrVal1, arrVal2, arrVal3, arrVal4, arrVal5, arrVal6, arrVal7
	DIm stbl_id, scol_id, sdata_id,sMajor_id, stbl_id2, scol_id2, sdata_id2,sMajor_id2	 							  
	Dim strgChangeOrgId

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	frm1.hOrgChangeId.value = PopupParent.gChangeOrgId

	Select Case iWhere			
		Case C_AcctCd
			arrParam(0) = "계정코드팝업"											' 팝업 명칭 
			arrParam(1) = "A_Acct, A_ACCT_GP" 											' TABLE 명칭 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Condition
			arrParam(4) = "A_ACCT.GP_CD=A_ACCT_GP.GP_CD AND A_ACCT.DEL_FG <> " & FilterVar("Y", "''", "S") & "  AND A_ACCT.MGNT_FG = " & FilterVar("Y", "''", "S") & "  and A_ACCT.mgnt_type = " & FilterVar("9", "''", "S") & " "    ' Where Condition
			arrParam(5) = "계정코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "A_ACCT.Acct_CD"												' Field명(0)
			arrField(1) = "A_ACCT.Acct_NM"												' Field명(1)
    		arrField(2) = "A_ACCT_GP.GP_CD"												' Field명(2)
			arrField(3) = "A_ACCT_GP.GP_NM"												' Field명(3)
			
			arrHeader(0) = "계정코드"												' Header명(0)
			arrHeader(1) = "계정코드명"												' Header명(1)
			arrHeader(2) = "그룹코드"												' Header명(2)
			arrHeader(3) = "그룹명"													' Header명(3)
		Case C_DocCur
		   If frm1.txtDocCur.readOnly = true then
				IsOpenPop = False
				Exit Function
			End If
	
			arrParam(0) = "통화코드 팝업"											' 팝업 명칭			
			arrParam(1) = "B_Currency"	    											' TABLE 명칭 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Cindition
			arrParam(4) = ""															' Where Condition
			arrParam(5) = "통화코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "Currency"	    											' Field명(0)
			arrField(1) = "Currency_desc"	    										' Field명(1)
    
			arrHeader(0) = "통화코드"												' Header명(0)
			arrHeader(1) = "통화코드명"												' Header명(1)
		Case C_MgntCd1
		
			If frm1.txtMgntCd1.readOnly = true then
				IsOpenPop = False
				Exit Function
			End If
		
		    Call QueryCtrlVal()
			
			stbl_id = frm1.hTblId.value
			scol_id = frm1.hDataColmID.value
			arrVal3 = frm1.hDataColmNm.value
			sMajor_id = frm1.hMajorCd.value

			if stbl_id = "" then
			IsOpenPop = False
			Exit Function
			End if
				
			strFrom = " A_OPEN_ACCT A, " & stbl_id & " B "
			if Trim(frm1.txtAcctCd.value) <> ""  then
				strWhere = " ACCT_CD =  " & FilterVar(frm1.txtAcctCd.value, "''", "S") & ""
				strWhere = strWhere  & " AND A.MGNT_VAL1 = B."&scol_id & " AND STATUS <> " & FilterVar("C", "''", "S") & " "
			else
				strWhere = " A.MGNT_VAL1 = B."&scol_id
				
			end if

			If sMajor_id  <> "" Then	
				strWhere = strWhere  & " and major_cd = '" & sMajor_id & "'"
			End If
						 
			arrParam(0) = "미결코드1팝업"											' 팝업 명칭 
			arrParam(1) = strFrom		    											' TABLE 명칭 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Condition
			arrParam(4) = strWhere														' Where Condition
			arrParam(5) = "미결코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "A.MGNT_VAL1"	    											' Field명(0)
			arrField(1) = "B."&arrVal3 	    											' Field명(1)

			arrHeader(0) = "미결관리1"												' Header명(0)
			arrHeader(1) = "미결코드"												' Header명(1)
		Case C_MgntCd2			

			If frm1.txtMgntCd2.readOnly = true then
				IsOpenPop = False
				Exit Function
			End If
			Call QueryCtrlVal2()
			
			stbl_id2 = frm1.hTblId2.value
			scol_id2 = frm1.hDataColmID2.value
			arrVal7 = frm1.hDataColmNm2.value
			sMajor_id2 = frm1.hMajorCd2.value
			
			if stbl_id2 = "" then
			IsOpenPop = False
			Exit Function
			End if
			
			'strFrom1 = " A_OPEN_ACCT A, " & stbl_id2 & " B "
			'strWhere1 = " ACCT_CD =  " & FilterVar(frm1.txtAcctCd.value, "''", "S") & ""
			'strWhere1 = strWhere1  & " AND  A.MGNT_VAL2 = B."&scol_id2 & " AND STATUS <> " & FilterVar("C", "''", "S") & " "

			strFrom1 = " A_OPEN_ACCT A, " & stbl_id2 & " B "
			strWhere1 = " ACCT_CD =  " & FilterVar(frm1.txtAcctCd.value, "''", "S") & ""
			strWhere1 = strWhere1  & " AND A.MGNT_VAL2 = B."&scol_id2
			
			If sMajor_id2  <> "" Then	
				strWhere1 = strWhere1  & " and major_cd = '" & sMajor_id2 & "'"
			End If

			arrParam(0) = "미결코드2팝업"											' 팝업 명칭 
			arrParam(1) = strFrom1	    												' TABLE 명칭 
			arrParam(2) = strCode														' Code Condition
			arrParam(3) = ""															' Name Condition
			arrParam(4) = strWhere1														' Where Condition
			arrParam(5) = "미결코드"												' 조건필드의 라벨 명칭 

			arrField(0) = "MGNT_VAL2"	    											' Field명(0)
			arrField(1) = "B."&arrVal7 	    											' Field명(1)
   
			arrHeader(0) = "미결관리1"												' Header명(0)
			arrHeader(1) = "미결코드"
	End Select
	
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = C_AcctCd Then
			frm1.txtAcctCd.focus
		ElseIf iWhere = C_DocCur Then
			frm1.txtDocCur.focus  
		ElseIf iWhere = C_MgntCd1 Then
			frm1.txtMgntCd1.focus
		ElseIf iWhere = C_MgntCd2 Then
			frm1.txtMgntCd2.focus     
		End If  
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function



'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere				
			Case C_AcctCd
				.txtAcctCd.focus
				.txtAcctCd.Value = arrRet(0)
				.txtAcctNm.value = arrRet(1)
				Call txtAcctcd_Onchange()
				
                       
			Case C_DocCur
				.txtDocCur.focus 
				.txtDocCur.value = arrRet(0)
			Case C_MgntCd1
				.txtMgntCd1.focus
				.txtMgntCd1.value =  arrRet(0)	
				.txtMgntCd1Nm.value =  arrRet(1)	
			Case C_MgntCd2
				.txtMgntCd2.focus
				.txtMgntCd2.value =  arrRet(0)		
				.txtMgntCd2Nm.value =  arrRet(1)		
		End Select

		'lgBlnFlgChgValue = True
	End With
End Function



'------------------------------------------  OpenBizCd()  -------------------------------------------------
'	Name : OpenBizCd()
'	Description : Cost PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtBizCd.className = "protected" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "사업장팝업"					' 팝업 명칭
	arrParam(1) = "B_BIZ_AREA"						' TABLE 명칭
	arrParam(2) = Trim(frm1.txtBizCd.Value)			' Code Condition
	arrParam(3) = ""								' Name Cindition
		' 권한관리 추가
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If
	arrParam(5) = "사업장"			
	
    arrField(0) = "BIZ_AREA_CD"						' Field명(0)
    arrField(1) = "BIZ_AREA_NM"						' Field명(1)
    
    arrHeader(0) = "사업장"						' Header명(0)
    arrHeader(1) = "사업장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	IF 	arrRet(0) <> "" then		
		Call SetBizCd(arrRet)
	Else
		frm1.txtBizCd.focus
		Exit Function
	end if
	
End Function

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

'------------------------------------------  SetBizCd()  --------------------------------------------------
'	Name : SetBizCd()
'	Description : Item Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetBizCd(Byval arrRet)
	
	frm1.txtBizCd.value = arrRet(0)		
	frm1.txtBizNm.value = arrRet(1)
	frm1.txtBizCd.focus
	lgBlnFlgChgValue = True				
	
End Function


Function QueryCtrlVal()

    Dim ArrRet

    
    Call CommonQueryRs("TBL_ID, DATA_COLM_ID, DATA_COLM_NM,COLM_DATA_TYPE,ISNULL(LTRIM(RTRIM(MAJOR_CD)),'')", _
                       "A_ACCT A, A_CTRL_ITEM B", _
                       "A.mgnt_cd1 = B.CTRL_CD AND A.ACCT_CD= " & FilterVar(frm1.txtAcctCd.value, "''", "S"),_
                       lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	ArrRet 	= REPLACE(lgF0,Chr(11),"") 
	
	IF Trim(REPLACE(lgF0,Chr(11),"")) <> "" then
		Strflag = "1"
		frm1.hTblId.value  = REPLACE(lgF0,Chr(11),"") 
		
		'ArrRet 	= Split(lgF1,Chr(11))
		frm1.hDataColmID.value  = REPLACE(lgF1,Chr(11),"") 
		'ArrRet 	= Split(lgF2,Chr(11))
		frm1.hDataColmNm.value = REPLACE(lgF2,Chr(11),"") 
		'ArrRet 	= Split(lgF4,Chr(11))
		frm1.hMajorCd.value = REPLACE(lgF4,Chr(11),"") 

	ELSE
		Strflag = "2"
		if replace(lgF3,Chr(11),"") = "D" then
			 frm1.txtMgntCd1Nm.value = "YYYY-MM-DD"
		Elseif replace(lgF3,Chr(11),"") = "N" then
			 frm1.txtMgntCd1Nm.value = "숫자는 구분자없이"
		End if	 
				
		
	END IF

End Function

Function QueryCtrlVal2()

    Dim ArrRet



    Call CommonQueryRs("TBL_ID, DATA_COLM_ID, DATA_COLM_NM,COLM_DATA_TYPE,ISNULL(LTRIM(RTRIM(MAJOR_CD)),'')", _
                       "A_ACCT A, A_CTRL_ITEM B", _
                       "A.mgnt_cd2 = B.CTRL_CD AND A.ACCT_CD= " & FilterVar(frm1.txtAcctCd.value, "''", "S"),_
                       lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
	ArrRet 	= REPLACE(lgF0,Chr(11),"") 

	IF Trim(REPLACE(lgF0,Chr(11),"") ) <> "" then
		Strflag = "1"
		frm1.hTblId2.value  = REPLACE(lgF0,Chr(11),"") 
		
		'ArrRet 	= Split(lgF1,Chr(11))
		frm1.hDataColmID2.value  = REPLACE(lgF1,Chr(11),"") 
		'ArrRet 	= Split(lgF2,Chr(11))
		frm1.hDataColmNm2.value = REPLACE(lgF2,Chr(11),"") 
		'ArrRet 	= Split(lgF4,Chr(11))
		frm1.hMajorCd2.value = REPLACE(lgF4,Chr(11),"") 


	ELSE
		Strflag = "2"		
		if replace(lgF3,Chr(11),"") = "D" then
			 frm1.txtMgntCd2Nm.value = "YYYY-MM-DD"
		Elseif replace(lgF3,Chr(11),"") = "N" then
			 frm1.txtMgntCd2Nm.value = "숫자는 구분자없이"
		End if	 
				
		
	END IF

End Function

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
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    


'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub


'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	

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

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)

End Sub
	

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


'==========================================================================================
Sub txtDocCur_OnChange()
    lgBlnFlgChgValue = True
    IF CommonQueryRs( "CURRENCY_DESC" , "B_CURRENCY" ,  " CURRENCY = " & FilterVar(frm1.txtDocCur.value , "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then   			        							
'		Call CurFormatNumericOCX()
		Call CurFormatNumSprSheet()
	END IF	    
End Sub

Function txtAcctCd_Onchange()
	With frm1
		
    
		Call CommonQueryRs("distinct A_ACCt.ACCT_CD, ACCT_NM ","A_ACCT, A_ACCT_CTRL_ASSN","A_ACCT.ACCT_CD = '" & .txtAcctCd.value & "' AND A_ACCT.acct_cd = a_acct_ctrl_assn.acct_cd" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
		If (lgF0 <> "X") And (Trim(lgF0) <> "") Then 
			.txtAcctNm.value = Left(lgF1, Len(lgF1)-1)    
			.txtMgntCd1.value = ""
			.txtMgntCd1Nm.value = ""
			.txtMgntCd2.value = ""
			.txtMgntCd2Nm.value = ""
			
			Call CommonQueryRs("CTRL_NM", _
                       "A_ACCT A, A_CTRL_ITEM B", _
                       "A.mgnt_cd1 = B.CTRL_CD AND A.ACCT_CD= " & FilterVar(frm1.txtAcctCd.value, "''", "S"),_
                       lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
			if lgF0 <> "" then
				CtrlCd.innerHTML = REPLACE(lgF0,Chr(11),"") 
			Else
				CtrlCd.innerHTML = "미결코드1" 
			End if
			
			Call CommonQueryRs("CTRL_NM", _
                       "A_ACCT A, A_CTRL_ITEM B", _
                       "A.mgnt_cd2 = B.CTRL_CD AND A.ACCT_CD= " & FilterVar(frm1.txtAcctCd.value, "''", "S"),_
                       lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			if lgF0 <> "" then
				CtrlCd2.innerHTML = REPLACE(lgF0,Chr(11),"")
			Else
				CtrlCd2.innerHTML = "미결코드2"
			End if
			
			Call ggoOper.SetReqAttr(frm1.txtMgntCd1,		"D")
			Call ggoOper.SetReqAttr(frm1.txtMgntCd2,		"D")
		'	Call ggoOper.SetReqAttr(frm1.txtMgntCd2,		"D")
			Call ggoOper.SetReqAttr(frm1.txtMgntCd1Nm,		"Q")
			Call ggoOper.SetReqAttr(frm1.txtMgntCd2Nm,		"Q")
			.txtAcctCd.focus
			
		Else       
			.txtAcctCd.value = ""
			.txtAcctNm.value = ""
			.txtMgntCd1.value = ""
			.txtMgntCd1Nm.value = ""
			.txtMgntCd2.value = ""
			.txtMgntCd2Nm.value = ""
			CtrlCd.innerHTML = "미결코드1"
			CtrlCd2.innerHTML = "미결코드2"
			Call ggoOper.SetReqAttr(frm1.txtMgntCd1,		"Q")
			Call ggoOper.SetReqAttr(frm1.txtMgntCd2,		"Q")
		'	Call ggoOper.SetReqAttr(frm1.txtMgntCd2,		"Q")
			Call ggoOper.SetReqAttr(frm1.txtMgntCd1Nm,		"Q")
			Call ggoOper.SetReqAttr(frm1.txtMgntCd2Nm,		"Q")      
			'.txtCtrlVal.value = ""
			'.txtCtrlValNm.value = ""       
			.txtAcctCd.focus       
		End If   
	End With
	
    txtAcctCd_OnChange = True
End Function



Sub  txtFrDueDt_DblClick(Button)
    If Button = 1 Then
        txtFrDueDt.Action = 7                        
        Call SetFocusToDocument("P")
		txtFrDueDt.Focus 
    End If
End Sub


Sub txtFrDueDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub

  
Sub  txtToDueDt_DblClick(Button)
    If Button = 1 Then
        txtToDueDt.Action = 7                        
        Call SetFocusToDocument("P")
		txtToDueDt.Focus 
    End If
End Sub


Sub txtToDueDt_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	ElseIf KeyAscii = 13 Then 
		Call Fncquery()
	End IF
End Sub


'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

	With frm1

		ggoSpread.Source = frm1.vspdData
		'금액 
		ggoSpread.SSSetFloatByCellOfCur GetKeyPos("A",10),-1, .txtDocCur.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur GetKeyPos("A",11),-1, .txtDocCur.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur GetKeyPos("A",12),-1, .txtDocCur.value, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetFloatByCellOfCur GetKeyPos("A",18),-1, .txtDocCur.value, PopupParent.ggExchRateNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gComNum1000, PopupParent.gComNumDec
	End With

End Sub


'========================================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtFromEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtFromEnterDt.Focus
	End If
End Sub


'========================================================================================================
Sub fpdtToEnterDt_DblClick(Button)
	If Button = 1 Then
       frm1.fpdtToEnterDt.Action = 7
       Call SetFocusToDocument("M")	
       frm1.fpdtToEnterDt.Focus
	End If
End Sub

'========================================================================================================
Sub fpdtFromEnterDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub


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
						<TD CLASS=TD5 NOWRAP>발생일자</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5403ra2_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
											 <script language =javascript src='./js/a5403ra2_fpDateTime2_txtToDt.js'></script></TD>							
						<TD CLASS=TD5 NOWRAP>거래통화</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDocCur" ALT="거래통화" MAXLENGTH="3" SIZE=10 STYLE="TEXT-ALIGN: left" tag  ="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtDocCur.Value, C_DocCur)"></TD>
								
					</TR>
					<TR>				
						<TD CLASS=TD5 NOWRAP>계정코드</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAcctCd" ALT="계정코드" MAXLENGTH="10" SIZE=11 tag ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtAcctCd.Value,C_AcctCd)">&nbsp;&nbsp;
											 <INPUT NAME="txtAcctNm" ALT="계정명"   MAXLENGTH="20" SIZE=18 tag ="14XXXU"></TD>
						<TD CLASS=TD5 ID="CtrlCd" NOWRAP>미결코드1</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMgntCd1" ALT="미결코드1" MAXLENGTH="30" SIZE=20 tag ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtMgntCd1.Value,C_MgntCd1)">&nbsp;						
											 <INPUT NAME="txtMgntCd1Nm" ALT="미결코드명1"   MAXLENGTH="30" SIZE=18 tag ="14XXXU"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>전표번호</TD>				
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtGlNo" SIZE=13 MAXLENGTH=18 tag="1XXXXU" ALT="전표번호"></TD>											 
						<TD CLASS=TD5 ID="CtrlCd2" NOWRAP>미결코드2</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtMgntCd2" ALT="미결코드2" MAXLENGTH="30" SIZE=20 tag ="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtMgntCd2.Value,C_MgntCd2)">&nbsp;								
											 <INPUT NAME="txtMgntCd2Nm" ALT="미결코드명2"   MAXLENGTH="30" SIZE=18 tag ="14XXXU"></TD>
					</TR>
					<TR>				
						<TD CLASS=TD5 NOWRAP>지급일</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5403ra2_ko441_fpDateTime1_txtFrDueDt.js'></script>&nbsp;~&nbsp;
											 <script language =javascript src='./js/a5403ra2_ko441_fpDateTime2_txtToDueDt.js'></script></TD>		
						<TD CLASS=TD5 NOWRAP>사업장</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 STYLE="TEXT-ALIGN: Left" tag=11NXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizCd()"> <INPUT TYPE=TEXT NAME="txtBizNm" SIZE=20 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="14" ALT="사업장명"></TD>					
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
				     <!-- 100 -->
						<script language =javascript src='./js/a5403ra2_vspdData_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<!-- 20 -->
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
		<!--	<TD WIDTH=100% HEIGHT=80%><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=80% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>  -->
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="hOrgChangeId"   tag="24">
<INPUT TYPE=hidden NAME="htxtAcctCd"   tag="24">
<INPUT TYPE=hidden NAME="htxtFromDt"   tag="24">
<INPUT TYPE=hidden NAME="htxtToDt"   tag="24">
<INPUT TYPE=hidden NAME="htxtDocCur"   tag="24">
<INPUT TYPE=HIDDEN NAME="htxtBizCd"		tag="24">
<INPUT TYPE=hidden NAME="htxtGlNo"   tag="24">
<INPUT TYPE=hidden NAME="htxtMgntCd1"   tag="24">
<INPUT TYPE=hidden NAME="htxtMgntCd2"   tag="24">
<INPUT TYPE=hidden NAME="lgSelectListDT"   tag="24">
<INPUT TYPE=hidden NAME="txtGlNoSeq"   tag="24">
<INPUT TYPE=hidden NAME="hTblId"   tag="24">
<INPUT TYPE=hidden NAME="hDataColmID"   tag="24">
<INPUT TYPE=hidden NAME="hDataColmNm"   tag="24">
<INPUT TYPE=hidden NAME="hMajorCd"   tag="24">

<INPUT TYPE=hidden NAME="hTblId2"   tag="24">
<INPUT TYPE=hidden NAME="hDataColmID2"   tag="24">
<INPUT TYPE=hidden NAME="hDataColmNm2"   tag="24">
<INPUT TYPE=hidden NAME="hMajorCd2"   tag="24">
</form>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


