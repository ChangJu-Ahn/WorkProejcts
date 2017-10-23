<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 회계관리 
*  2. Function Name        :
*  3. Program ID           : a5401ba1
*  4. Program Name         : 환평가작업 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2006/03/28
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Jeong Yong Kyun
* 10. Modifier (Last)      : 
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<!-- #Include file="../../inc/uni2kcm.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<Script Language="VBScript">

Option Explicit																			'☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID    = "a5310mb1.asp"
Const COOKIE_SPLIT  =  4877																'Cookie Split String

Dim C_EXCH_MODULE_CD
dIM C_EXCH_MODULE_CD_PB
Dim C_EXCH_MODULE_NM
Dim C_EXCH_MODULE_USP
Dim C_EXCH_ACCT_CD
Dim C_EXCH_ACCT_CD_PB
Dim C_EXCH_ACCT_NM
Dim C_MASTER_REFLECT_FG
Dim C_USP_FG

<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgIsOpenPop
Dim IsOpenPop

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE												'⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False															'⊙: Indicates that no value changed
	lgIntGrpCount     = 0																'⊙: Initializes Group View Size
    lgStrPrevKey      = ""																'⊙: initializes Previous Key
    lgSortKey         = 1																'⊙: initializes sort direction

    lgStrPrevKey = ""																	'initializes Previous Key
    lgLngCurRows = 0																	'initializes Deleted Rows Count
End Sub

'========================================================================================================
Sub SetDefaultVal()

End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) --------------------------------------------------------------

	<% Call loadInfTB19029A("Q", "G", "NOCOOKIE", "MA") %>

	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) --------------------------------------------------------------
   '------ Developer Coding part (End   ) --------------------------------------------------------------
End Sub

'========================================================================================================
Sub InitComboBox()
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	ggoSpread.source = frm1.vspdData
	ggoSpread.SetCombo "Y" & vbtab & "N" , C_MASTER_REFLECT_FG
End Sub

'Function InitComboBoxGrid()
'    ggoSpread.Source = frm1.vspdData
'
'	Call CommonQueryRs("minor_cd, minor_nm", "b_minor", "major_cd=" & FilterVar("A1045", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
'
'	lgF0 = "" & chr(11) & lgF0
'	lgF1 = "" & chr(11) & lgF1
'
'	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_EXCH_MODULE_CD
'   ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_EXCH_MODULE_NM
'
'End Function

'========================================================================================================
Sub InitData()

End Sub

'========================================================================================================
Sub InitSpreadPosVariables()
	C_EXCH_MODULE_CD	= 1
	C_EXCH_MODULE_CD_PB = 2	
	C_EXCH_MODULE_NM	= 3
	C_EXCH_MODULE_USP	= 4
	C_EXCH_ACCT_CD		= 5
	C_EXCH_ACCT_CD_PB	= 6
	C_EXCH_ACCT_NM		= 7
	C_MASTER_REFLECT_FG	= 8
	C_USP_FG			= 9
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()

	With frm1.vspdData
		.MaxCols = C_USP_FG + 1                                                      ' ☜:☜: Add 1 to Maxcols

		.Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
		.ColHidden = True                                                            ' ☜:☜:

        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021128", ,parent.gAllowDragDropSpread

		ggoSpread.ClearSpreadData

		.ReDraw = False

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit   C_EXCH_MODULE_CD,	    "모듈구분",   8,0,,2,2
		ggoSpread.SSSetButton C_EXCH_MODULE_CD_PB
		ggoSpread.SSSetEdit   C_EXCH_MODULE_NM,		"모듈구분명",  10
		ggoSpread.SSSetEdit	  C_EXCH_MODULE_USP,	"실행USP",			30,0,,100,2
		ggoSpread.SSSetEdit	  C_EXCH_ACCT_CD,		"환평가계정코드",	15,0,,15,2
		ggoSpread.SSSetButton C_EXCH_ACCT_CD_PB
		ggoSpread.SSSetEdit	  C_EXCH_ACCT_NM,	    "환평가계정코드명",	20,0,,30,2
		ggoSpread.SSSetCombo  C_MASTER_REFLECT_FG,  "마스터반영",       10,2
		ggoSpread.SSSetEdit	  C_USP_FG,			    "",						5,0,,5,2

'		Call ggoSpread.SSSetColHidden(C_EXCH_MODULE_CD,C_EXCH_MODULE_CD,True)
		Call ggoSpread.SSSetColHidden(C_MASTER_REFLECT_FG,C_MASTER_REFLECT_FG,True)
		Call ggoSpread.SSSetColHidden(C_USP_FG,C_USP_FG,True)	   

		.ReDraw = True
		Call SetSpreadLock
    End With
End Sub

'========================================================================================================
Sub SetSpreadLock()
	Dim ii

    With frm1
		ggoSpread.Source = .vspdData
        .vspdData.ReDraw = False

		For ii = 1 To .vspdData.MaxRows
			ggoSpread.SSSetProtected    C_EXCH_MODULE_CD    , ii    ,ii
			ggoSpread.SSSetProtected    C_EXCH_MODULE_CD_PB , ii    ,ii
			ggoSpread.SSSetProtected    C_EXCH_MODULE_NM	, ii    ,ii
			.vspddata.col = C_USP_FG
			.vspddata.row = ii
			If Trim(.vspddata.Text) = "USP" Then
				ggoSpread.SSSetRequired  C_EXCH_MODULE_USP  , ii    ,ii
				ggoSpread.SSSetProtected C_EXCH_ACCT_CD		, ii    ,ii
				ggoSpread.SSSetProtected C_EXCH_ACCT_CD_PB	, ii    ,ii
				ggoSpread.SSSetProtected C_EXCH_ACCT_NM		, ii    ,ii				
			Else
				ggoSpread.SSSetProtected C_EXCH_MODULE_USP , ii    ,ii
				ggoSpread.SSSetProtected C_EXCH_ACCT_CD		, ii    ,ii
				ggoSpread.SSSetProtected C_EXCH_ACCT_CD_PB	, ii    ,ii
				ggoSpread.SSSetProtected C_EXCH_ACCT_NM		, ii    ,ii								
			End If
			
			ggoSpread.SSSetProtected	C_MASTER_REFLECT_FG	, ii    ,ii								
			ggoSpread.SSSetProtected	.vspdData.MaxCols   , ii	,ii
		Next

        .vspdData.ReDraw = True
    End With
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired     C_EXCH_MODULE_CD    , pvStartRow, pvEndRow
		ggoSpread.SpreadUnLock      C_EXCH_MODULE_CD_PB , pvStartRow, C_EXCH_MODULE_CD_PB , pvEndRow
		ggoSpread.SSSetProtected    C_EXCH_MODULE_NM    , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_EXCH_MODULE_USP   , pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_EXCH_ACCT_CD		, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_EXCH_ACCT_CD_PB	, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected    C_EXCH_ACCT_NM		, pvStartRow, pvEndRow				
		ggoSpread.SSSetProtected    C_MASTER_REFLECT_FG , pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    End With
End Sub

'========================================================================================
Sub SetSpreadColorAfterExchModuleSelect(ByVal lRow)
	Dim strCd
	Dim strColNm,strTableid,strwhere
	Dim	strUspAcctfg,iFlag
	Dim IntRetCD

    With frm1
		.vspddata.col = C_EXCH_MODULE_CD
		.vspddata.row = .vspddata.activerow
		strCd = .vspddata.Text

		If strCd <> "" Then
			strColNm = " REFERENCE "
			strTableid = " B_CONFIGURATION "
			strwhere = " MAJOR_CD = " & FilterVar("A1045","''","S") & " AND MINOR_CD = " & FilterVar(strCd,"''","S") & " AND  SEQ_NO=1 "

			If CommonQueryRs2by2( strColNm , strTableid ,  strwhere , lgF2By2) = False Then
				.vspddata.col = C_EXCH_MODULE_USP
				.vspddata.text = ""
				.vspddata.col = C_EXCH_ACCT_CD
				.vspddata.text = ""
				.vspddata.col = C_EXCH_ACCT_NM
				.vspddata.text = ""						

'				ggoSpread.SpreadLock  		C_EXCH_MODULE_NM	, lRow, C_EXCH_MODULE_NM   , lRow		    		    			
				ggoSpread.SpreadLock  		C_EXCH_MODULE_USP	, lRow, C_EXCH_MODULE_USP   , lRow		    		    			
				ggoSpread.SpreadLock    	C_EXCH_ACCT_CD		, lRow, C_EXCH_ACCT_CD		, lRow
				ggoSpread.SpreadLock   		C_EXCH_ACCT_CD_PB   , lRow, C_EXCH_ACCT_CD_PB	, lRow
			    ggoSpread.SpreadLock		C_EXCH_ACCT_NM		, lRow, C_EXCH_ACCT_NM		, lRow

				IntRetCD = DisplayMsgBox("AM0040", "X", strCd, "X")		    
				Exit Sub
			Else	
				iFlag = Split(lgF2By2, Chr(11))
				strUspAcctfg = iFlag(1)
			End If		
		Else
			Exit Sub
		End If	

		ggoSpread.Source = .vspdData    
	    .vspdData.ReDraw = False

		.vspddata.col = C_EXCH_MODULE_USP
		.vspddata.text = ""
		.vspddata.col = C_EXCH_ACCT_CD
		.vspddata.text = ""
		.vspddata.col = C_EXCH_ACCT_NM
		.vspddata.text = ""			

		If Trim(strUspAcctfg) =  "USP" Then
'			ggoSpread.SpreadLock  		C_EXCH_MODULE_NM	, lRow, C_EXCH_MODULE_NM   , lRow		    		    					
			ggoSpread.SpreadUnLock	    C_EXCH_MODULE_USP   , lRow, C_EXCH_MODULE_USP	, lRow		
		    ggoSpread.SSSetRequired  	C_EXCH_MODULE_USP	, lRow, lRow		    
			ggoSpread.SpreadLock    	C_EXCH_ACCT_CD		, lRow, C_EXCH_ACCT_CD		, lRow
			ggoSpread.SpreadLock   		C_EXCH_ACCT_CD_PB   , lRow, C_EXCH_ACCT_CD_PB	, lRow
		    ggoSpread.SpreadLock		C_EXCH_ACCT_NM		, lRow, C_EXCH_ACCT_NM		, lRow			
'			ggoSpread.SpreadUnLock	    C_MASTER_REFLECT_FG , lRow, C_MASTER_REFLECT_FG	, lRow
'			ggoSpread.SSSetRequired  	C_MASTER_REFLECT_FG	, lRow, lRow		    
		Else
'			ggoSpread.SpreadLock  		C_EXCH_MODULE_NM	, lRow, C_EXCH_MODULE_NM   , lRow		    		    					
			ggoSpread.SSSetProtected  	C_EXCH_MODULE_USP	, lRow, lRow		    		    
			ggoSpread.SpreadUnLock    	C_EXCH_ACCT_CD		, lRow, C_EXCH_ACCT_CD		, lRow
			ggoSpread.SpreadUnLock   	C_EXCH_ACCT_CD_PB   , lRow, C_EXCH_ACCT_CD_PB	, lRow
			ggoSpread.SSSetRequired    	C_EXCH_ACCT_CD		, lRow, lRow
'			ggoSpread.SpreadLock	    C_MASTER_REFLECT_FG , lRow, C_MASTER_REFLECT_FG	, lRow
		End If

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

			C_EXCH_MODULE_CD	= iCurColumnPos(1)
			C_EXCH_MODULE_CD_PB	= iCurColumnPos(2)					
			C_EXCH_MODULE_NM	= iCurColumnPos(3)
			C_EXCH_MODULE_USP	= iCurColumnPos(4)
			C_EXCH_ACCT_CD		= iCurColumnPos(5)
			C_EXCH_ACCT_CD_PB	= iCurColumnPos(6)
			C_EXCH_ACCT_NM		= iCurColumnPos(7)
			C_MASTER_REFLECT_FG	= iCurColumnPos(8)
			C_USP_FG			= iCurColumnPos(9)
    End Select    
End Sub

'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	With frm1
		If IsOpenPop = True Then Exit Function 

		Select Case iWhere
			Case 1
				arrParam(0) = "환평가버젼팝업"											' 팝업 명칭 
				arrParam(1) = " (select distinct(ver_cd) ver_cd from a_exchange_version) a " 
				arrParam(2) = Trim(.txtVerCd.value)											' Code Condition
				arrParam(3) = " "															' Name Cindition
				arrParam(4) = " "															' Where Condition
				arrParam(5) = "환평가버젼"												' 조건필드의 라벨 명칭 

				arrField(0) = "VER_CD"														' Field명(0)
'				arrField(1) = "VER_CD"														' Field명(1)

			 	arrHeader(0) = "환평가버젼"												' Header명(2)
'				arrHeader(1) = "환평가버젼"
			Case 2
				arrParam(0) = "계정코드팝업"										    ' 팝업 명칭 
				arrParam(1) = "A_Acct A, A_ACCT_GP B"										' TABLE 명칭 
				.vspddata.Col = C_EXCH_ACCT_CD
				.vspddata.Row = .vspddata.ActiveRow
				arrParam(2) = Trim(.vspddata.Text)											' Code Condition
				arrParam(3) = ""															' Name Cindition
				arrParam(4) = "A.GP_CD=B.GP_CD AND A.DEL_FG <> " & FilterVar("Y", "''", "S") & " and fx_eval_fg = "	& FilterVar("Y","''","S")
				arrParam(4) = arrParam(4) & " and a.bs_pl_fg= " & FilterVar("BS","''","S")							' Where Condition
				arrParam(5) = "계정코드"												' 조건필드의 라벨 명칭 

				arrField(0) = "A.Acct_CD"													' Field명(0)
				arrField(1) = "A.Acct_NM"													' Field명(1)
			    arrField(2) = "B.GP_CD"														' Field명(2)
				arrField(3) = "B.GP_NM"														' Field명(3)
			 
				arrHeader(0) = "계정코드"												' Header명(0)
				arrHeader(1) = "계정코드명"												' Header명(1)
				arrHeader(2) = "그룹코드"												' Header명(2)
				arrHeader(3) = "그룹명"
			Case 3
				arrParam(0) = "모듈구분팝업"										    ' 팝업 명칭 
				arrParam(1) = " b_minor "													' TABLE 명칭 
				.vspddata.Col = C_EXCH_MODULE_CD
				.vspddata.Row = .vspddata.ActiveRow
				arrParam(2) = Trim(.vspddata.Text)											' Code Condition
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
		Call EscPopup(iWhere)
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function

'======================================================================================================
'   Function Name : EscPopup(Byval iWhere)
'   Function Desc : 
'=======================================================================================================
Function EscPopup(Byval iWhere)
	With frm1
		Select Case iWhere
			Case 1
				.txtVerCd.focus 
			Case 2
				Call SetActiveCell(.vspdData,C_EXCH_ACCT_CD,.vspdData.ActiveRow ,"M","X","X")
			Case 3
				Call SetActiveCell(.vspdData,C_EXCH_MODULE_CD,.vspdData.ActiveRow ,"M","X","X")				
		End Select    
	End With
End Function

'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================
Function SetPopup(Byval arrRet,Byval pWhere)
	With frm1
		.vspddata.row = .vspddata.activerow

		Select Case pWhere
			Case 1
				.txtVerCd.Value = arrRet(0)
			Case 2	
				.vspdData.Col = C_EXCH_ACCT_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_EXCH_ACCT_NM
				.vspdData.Text = arrRet(1)
'				Call vspdData_Change(C_ACCT_GP, .vspddata.activerow )  ' 변경이 일어났다고 알려줌        
				Call SetActiveCell(.vspdData,C_EXCH_ACCT_CD,.vspdData.ActiveRow ,"M","X","X")
			Case 3
				.vspdData.Col = C_EXCH_MODULE_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_EXCH_MODULE_NM
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(C_EXCH_MODULE_CD, .vspddata.activerow )  ' 변경이 일어났다고 알려줌        
				Call SetActiveCell(.vspdData,C_EXCH_MODULE_CD,.vspdData.ActiveRow ,"M","X","X")				
				
		End Select    
	End With
 
	If pWhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If 
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format

	'------ Developer Coding part (Start ) --------------------------------------------------------------

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                            '⊙: Lock  Suitable  Field

    Call InitVariables                                                               '⊙: Setup the Spread sheet
	Call SetDefaultVal

    Call InitSpreadSheet
    Call InitComboBox
'    Call InitComboBoxGrid

	Call SetToolbar("1100111100011111")                                              '☆: Developer must customize

	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
Function FncQuery()
    Dim IntRetCD
    FncQuery = False																		'☜: Processing is NG
    Err.Clear																				'☜: Clear err status

    ggoSpread.Source = frm1.vspdData

    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")						'☜: Data is changed.  Do you want to display it?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")													'☜: Clear Contents  Field

    If Not chkField(Document, "1") Then														'☜: This function check required field
		Exit Function
    End If

    Call InitVariables																		'⊙: Initializes local global variables

    If DbQuery = False Then
		Exit Function
    End If																					'☜: Query db data

    Set gActiveElement = document.ActiveElement
    FncQuery = True																			'☜: Processing is OK
End Function

'========================================================================================================
Function FncNew()
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
Function FncDelete()
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
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

	    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then				'⊙: Check If data is chaged
	    	IntRetCD = DisplayMsgBox("900001", "X", "X", "X")								'⊙: Display Message(There is no changed data.)
	    	Exit Function
	    End If

		If Not chkFieldByCell(.txtVerCd, "A", "1") Then Exit Function

		'-----------------------
	    'Check content area
	    '----------------------- 
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
   	
    FncSave = True		                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncCopy()
    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncCancel()
    Dim lRow
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    ggoSpread.EditUndo
    FncCancel = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    Dim imRow2
	Dim iCurRowPos

    On Error Resume Next
    Err.Clear
    
    FncInsertRow = False

	With frm1
		If .txtVerCd.value = "" Then 
			IntRetCD = DisplayMsgBox ("173133", "X", .txtVerCd.Alt, "X")
			Exit Function
		End If
	    
		If IsNumeric(Trim(pvRowCnt)) Then
	        imRow = CInt(pvRowCnt)
	    Else
	        imRow = AskSpdSheetAddRowCount()
			If imRow = "" Then
	            Exit Function
	        End If
	    End If

        .vspdData.ReDraw = False
		.vspdData.focus
        ggoSpread.Source = .vspdData
		iCurRowPos       = .vspdData.ActiveRow

        For imRow2 = 1 To imRow
            ggoSpread.InsertRow ,1         
            Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow)
        Next

        .vspdData.ReDraw = True
    End With
    
    If Err.number = 0 Then
		FncInsertRow = True
    End If
    
    Set gActiveElement = document.ActiveElement
End Function

'========================================================================================================
Function FncDeleteRow()
	Dim lDelRows
    Dim DelItemSeq

    FncDeleteRow = False                                                         '☜: Processing is NG
	Err.Clear
	
	With frm1
		ggoSpread.Source = .vspdData 

		.vspdData.Row = .vspdData.ActiveRow
		.vspdData.Col = 0 

		If .vspdData.MaxRows < 1 Or .vspdData.Text = ggoSpread.InsertFlag Then Exit Function

		.vspdData.Col = 1 
		DelItemSeq = .vspdData.Text

		lDelRows = ggoSpread.DeleteRow
	End With

    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
	Err.Clear
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
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncNext()
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    FncNext = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncExcel()
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	Call Parent.FncExport(Parent.C_SINGLE)
    FncExcel = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncFind()
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	Call Parent.FncFind(Parent.C_SINGLE, True)
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
	FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
	FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function DbQuery()
	Dim strVal

    Err.Clear                                                                    '☜: Clear err status
    DbQuery = False                                                              '☜: Processing is NG

	If LayerShowHide(1) = False Then
	    Exit Function
	End If

	'------ Developer Coding part (Start)  --------------------------------------------------------------
    With frm1
		strVal = BIZ_PGM_ID & "?txtMode="  & Parent.UID_M0001
        strVal = strVal     & "&txtVerCd=" & .txtVerCd.value 
    End With

	'------ Developer Coding part (End )   --------------------------------------------------------------
	Call RunMyBizASP(MyBizASP, strVal)													'☜: Run Biz Logic

    DbQuery = True
End Function

'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel

    DbSave = False																		'⊙: Processing is NG
    Call LayerShowHide(1)
    
    On Error Resume Next																'☜: Protect system from crashing
    Err.Clear       

    DbSave = False																		'☜: Processing is NG

	With frm1
		.txtMode.value = parent.UID_M0002

		lGrpCnt = 1
		strVal = ""
		strDel = ""

		'-----------------------
		'Data manipulate area
		'-----------------------
		For lRow = 1 To .vspdData.MaxRows
		    .vspdData.Row = lRow
			.vspdData.Col = 0
	
		    Select Case .vspdData.Text
  				Case ggoSpread.InsertFlag	   											'☜: 신규 
					strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep					
				    .vspdData.Col = C_EXCH_MODULE_CD
				    strVal = strVal & UCASE(Trim(.vspdData.Text)) & Parent.gColSep
		            .vspdData.Col = C_EXCH_MODULE_USP
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep				    
		            .vspdData.Col = C_EXCH_ACCT_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MASTER_REFLECT_FG
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep			            
				Case ggoSpread.UpdateFlag												'☜: 수정 
					strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep					
				    .vspdData.Col = C_EXCH_MODULE_CD
				    strVal = strVal & UCASE(Trim(.vspdData.Text)) & Parent.gColSep
		            .vspdData.Col = C_EXCH_MODULE_USP
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep				    
		            .vspdData.Col = C_EXCH_ACCT_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MASTER_REFLECT_FG
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep			            
		        Case ggoSpread.DeleteFlag												'☜: 삭제 
					strVal = strVal & "D" & Parent.gColSep & lRow & Parent.gColSep					
				    .vspdData.Col = C_EXCH_MODULE_CD
				    strVal = strVal & UCASE(Trim(.vspdData.Text)) & Parent.gColSep
		            .vspdData.Col = C_EXCH_MODULE_USP
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep				    
		            .vspdData.Col = C_EXCH_ACCT_CD
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
		            .vspdData.Col = C_MASTER_REFLECT_FG
		            strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep			            
		    End Select
		Next
	
		.txtSpread.value = strDel & strVal
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)															'☜: 비지니스 ASP 를 가동 
	End With		

    DbSave = True																					'☜: Processing is OK
End Function

'========================================================================================================
Function DbDelete()
    Err.Clear																						'☜: Clear err status
    DbDelete = False																				'☜: Processing is NG
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
    DbDelete = True																					'☜: Processing is OK
End Function

'========================================================================================================
Sub DbQueryOk()
	lgIntFlgMode = parent.OPMD_UMODE							'Indicates that current mode is Update mode

	Call SetToolbar("110011110001111")																'⊙: 버튼 툴바 제어 
	Call SetSpreadLock()
End Sub

'========================================================================================================
Sub DbSaveOk()

    Call InitVariables()									'Initializes local global variables
	ggoSpread.Source = frm1.vspdData        				
	ggoSpread.ClearSpreadData
					
	Call DBquery()  
End Sub

'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  --------------------------------------------------------------
	'------ Developer Coding part (End )   --------------------------------------------------------------
End Sub

'========================================================================================================
Function ExeCopy() 
	Dim IntRetCD
	Dim strVal

	On Error Resume Next
	Err.Clear 

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If

	If Not chkField(Document, "2") Then
		Exit Function
	End If	

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")

	If IntRetCD = vbNo Then
		Exit Function
	End If

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	ExeCopy = False

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003
	strVal = strVal & "&txtVerCd=" & Trim(frm1.htxtVerCd.value)
	strVal = strVal & "&txtNewVerCd=" & Trim(frm1.txtNewVerCd.value)

	Call RunMyBizASP(MyBizASP, strVal)

	ExeCopy = True
End Function

'========================================================================================================
Private Sub vspdData_Change(ByVal Col , ByVal Row )
	With frm1 	
	    ggoSpread.Source = .vspdData
	    ggoSpread.UpdateRow Row

	    Select Case Col
			Case C_EXCH_MODULE_CD
				Call SetSpreadColorAfterExchModuleSelect(Row)
	    End Select
	End With	    
End Sub


'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1
	    .vspdData.Row = Row
		Select Case Col
			Case C_EXCH_ACCT_CD_PB
				Call OpenPopup(2)
			Case C_EXCH_MODULE_CD_PB
				Call OpenPopup(3)				
	    End Select
	End With    
End Sub

'========================================================================================================
sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
	ggoSpread.source = frm1.vspdData
	Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 	
End Sub

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgStrPrevKeyIndex <> "" Then
      		Call DisableToolBar(Parent.TBC_QUERY)
      		If DBQuery = False Then
      			Call RestoreToolBar()
      			Exit Sub
      		End If
    	End If
    End If
End Sub

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col,lgSortKey
            lgSortKey = 1
        End If
    End If
End Sub

'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
    End If
End Sub

</SCRIPT>
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%= LR_SPACE_TYPE_00 %>>
	<TR>
		<TD NOWRAP  <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD NOWRAP >
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD NOWRAP  WIDTH=10>&nbsp;</TD>
					<TD NOWRAP  CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<TD NOWRAP  background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
                    <TD NOWRAP  WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD NOWRAP  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD NOWRAP  <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD NOWRAP  HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
						    <TR>
								<TD NOWRAP CLASS=TD5>환평가버젼</TD>
								<TD NOWRAP CLASS=TD656><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtVerCd" SIZE=10 MAXLENGTH=3 STYLE="TEXT-ALIGN: Left" tag="12XXXU" ALT="환평가버젼"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInputType2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup('1')"></TD>
							</TR>
    					</TABLE>
					</TD>
				</TR>
				<TR>
					<TD NOWRAP  <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD NOWRAP  WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD NOWRAP  HEIGHT="100%">
									<script language =javascript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"  id=vspdData><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>')</script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD NOWRAP  <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
				</TR>
				<TR HEIGHT=20>
					<TD WIDTH=100% HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">					
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD NOWRAP CLASS=TD5>신규버전</TD>
									<TD NOWRAP CLASS=TD656><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtNewVerCd" SIZE=10 MAXLENGTH=3 tag="22XXXU" ALT="신규버전">&nbsp;<BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeCopy()" Flag=1>복사실행</BUTTON></TD>
									<TD NOWRAP CLASS=TD656></TD>
									<TD NOWRAP CLASS=TD655></TD>
								</TR>
							</TABLE>
						</FIELDSET>	
					</TD>
				</TR>				
			</TABLE>
		</TD>
	</TR>

	<TR>
		<TD NOWRAP  WIDTH=100% HEIGHT=<%=BIZSIZE%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread"	tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows"		tag="24" TABINDEX = "-1"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode"			tag="24" TABINDEX = "-1">
<INPUT TYPE=HIDDEN NAME="htxtVerCd"			tag="24" TABINDEX = "-1">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
