<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : MULTI & SINGLE
*  3. Program ID           : ha203ma1
*  4. Program Name         : ha203ma1
*  5. Program Desc         : 퇴직금추계액조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/30f
*  8. Modified date(Last)  : 2003/06/16
*  9. Modifier (First)     : TGS 최용철 
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit

Const BIZ_PGM_ID      = "ha203mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg
Dim gtxtChargeType
Dim IsOpenPop
Dim lgOldRow
Dim lgBlnDataChgFlg

Dim C_HAA010T_NAME  
Dim C_EMP_NO        
Dim C_DEPT_CD       
Dim C_TOT_DUTY_MM   
Dim C_PAY_ESTI_AMT   
Dim C_BONUS_ESTI_AMT 
Dim C_YEAR_ESTI_AMT  
Dim C_AVR_WAGES_AMT 
Dim C_RETIRE_AMT    
Dim C_TOT_PROV_AMT   
Dim C_DUTY_YY        
Dim C_DUTY_MM    
Dim C_DUTY_DD       
Dim C_RETIRE_YYMM    

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
Sub initSpreadPosVariables()  
	C_HAA010T_NAME    = 1
	C_EMP_NO          = 2
	C_DEPT_CD         = 3
	C_TOT_DUTY_MM     = 4
	C_PAY_ESTI_AMT    = 5
	C_BONUS_ESTI_AMT  = 6
	C_YEAR_ESTI_AMT   = 7
	C_AVR_WAGES_AMT   = 8
	C_RETIRE_AMT      = 9
	C_TOT_PROV_AMT    = 10
	C_DUTY_YY         = 11
	C_DUTY_MM         = 12
	C_DUTY_DD         = 13
	C_RETIRE_YYMM     = 14	
End Sub
'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
	lgOldRow = 0

	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
	lgBlnDataChgFlg   = False
End Sub

'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
    Dim strLastDate
    Dim strYYYYMM
    Dim strDate, strMonth, strDay

    strDate		= UniConvYYYYMMDDToDate(parent.gDateFormat,frm1.txtpay_yymm_dt.Year,Right("0" & frm1.txtpay_yymm_dt.Month,2),frm1.txtpay_yymm_dt.Day)
    strYYYYMM	= UniConvDateToYYYYMM(strDate,parent.gDateFormat,"")

    strLastDate	= UNIGetLastDay (strDate,parent.gDateFormat)
    lgKeyStream	= strYYYYMM & parent.gColSep
    lgKeyStream	= lgKeyStream & Trim(Frm1.txtEmp_no.value) & parent.gColSep                     'You Must append one character(parent.gColSep)
    lgKeyStream	= lgKeyStream & lgUsrIntcd & parent.gColSep
    lgKeyStream	= lgKeyStream & strLastDate & parent.gColSep
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex
    If Frm1.vspdData.MaxRows > 0 Then
        Call vspdData_Click(1 , 1)
        Frm1.vspdData.focus
        Set gActiveElement = document.ActiveElement
	End If
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================

Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	With frm1.vspdData

		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021119",,parent.gAllowDragDropSpread        
		.ReDraw = false
		.MaxCols = C_RETIRE_YYMM + 1												
		.Col = .MaxCols															
		.ColHidden = True
		.MaxRows = 0
		ggoSpread.ClearSpreadData	
		Call GetSpreadColumnPos("A")		
		Call AppendNumberPlace("6","4","0")

		ggoSpread.SSSetEdit  C_HAA010T_NAME  , "성명"       , 8,,,30,2		'Lock/ Edit
		ggoSpread.SSSetEdit  C_EMP_NO        , "사번"       , 11,,,13,2		'Lock/ Edit
		ggoSpread.SSSetEdit  C_DEPT_CD       , "부서"       , 14,,,40,2		'Lock/ Edit
		ggoSpread.SSSetFloat C_TOT_DUTY_MM   , "근속개월"   , 10,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_PAY_ESTI_AMT  , "급여총액"   , 12,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_BONUS_ESTI_AMT, "상여총액"   , 12,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_YEAR_ESTI_AMT , "연월차총액" , 12,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_AVR_WAGES_AMT , "평균임금"   , 12,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_RETIRE_AMT    , "퇴직금"     , 12,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetFloat C_TOT_PROV_AMT  , "총지급액"   , 12,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		ggoSpread.SSSetEdit  C_DUTY_YY       , "근속YY"     , 10,,,20,2		'Lock/ Edit
		ggoSpread.SSSetEdit  C_DUTY_MM       , "근속MM"     , 10,,,20,2		'Lock/ Edit
		ggoSpread.SSSetEdit  C_DUTY_DD       , "근속DD"     , 10,,,20,2		'Lock/ Edit
		ggoSpread.SSSetEdit  C_RETIRE_YYMM   , "정산년월"   , 10,,,20,2		'Lock/ Edit
		
       Call ggoSpread.SSSetColHidden(C_DUTY_YY,C_DUTY_YY,True)	
       Call ggoSpread.SSSetColHidden(C_DUTY_MM,C_DUTY_MM,True)	
       Call ggoSpread.SSSetColHidden(C_DUTY_DD,C_DUTY_DD,True)	
       Call ggoSpread.SSSetColHidden(C_RETIRE_YYMM,C_RETIRE_YYMM,True)	                     	
	
		Call SetSpreadLock

    End With

End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

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

			C_HAA010T_NAME    = iCurColumnPos(1)
			C_EMP_NO          = iCurColumnPos(2)
			C_DEPT_CD         = iCurColumnPos(3)
			C_TOT_DUTY_MM     = iCurColumnPos(4)
			C_PAY_ESTI_AMT    = iCurColumnPos(5)
			C_BONUS_ESTI_AMT  = iCurColumnPos(6)
			C_YEAR_ESTI_AMT   = iCurColumnPos(7)
			C_AVR_WAGES_AMT   = iCurColumnPos(8)
			C_RETIRE_AMT      = iCurColumnPos(9)
			C_TOT_PROV_AMT    = iCurColumnPos(10)
			C_DUTY_YY         = iCurColumnPos(11)
			C_DUTY_MM         = iCurColumnPos(12)
			C_DUTY_DD         = iCurColumnPos(13)
			C_RETIRE_YYMM     = iCurColumnPos(14)           
    End Select    
End Sub
'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to
              Exit For
           End If
        Next
     End If
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

	Dim strYear,strMonth,strDay

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

	Call ggoOper.FormatNumber(frm1.txtTot_duty_mm, "999999", "0", False, 0)	                '총근속개월수Format
    Call ggoOper.FormatDate(frm1.txtpay_yymm_dt, parent.gDateFormat, 2)                    '싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.
    Call ggoOper.FormatDate(frm1.txtRetire_yymm_dt, parent.gDateFormat, 2)                    '싱글에서 년월말 입력하고 싶은경우 다음 함수를 콜한다.

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                                 'Initializes local global variables

    Call FuncGetAuth("HA203MA1", parent.gUsrID, lgUsrIntCd)

	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
	
	Call ExtractDateFrom("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gServerDateType,strYear,strMonth,strDay)
    frm1.txtpay_yymm_dt.focus
    frm1.txtpay_yymm_dt.Year	= strYear
    frm1.txtpay_yymm_dt.Month	= strMonth
    frm1.txtpay_yymm_dt.Day		= strDay

	Call CookiePage (0)                                                             '☜: Check Cookie

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
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If  txtEmp_no_Onchange()  then
        Exit Function
    End If 

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables

    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If

    Call MakeKeyStream("X")

    If DbQuery = False Then
        Exit Function
    End If

    FncQuery = True	                                                             '☜: Processing is OK

End Function

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
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel()
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
	Call Parent.FncFind(parent.C_SINGLE, True)
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
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True

End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

    If LayerShowHide(1) = False Then
		Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

    Frm1.txtEMP_NO.focus

	Call SetToolbar("1100000000011111")												'⊙: Set ToolBar

    Call InitData()
    Call ggoOper.LockField(Document, "Q")
    frm1.vspdData.focus
    lgBlnFlgChgValue = False
End Function

'========================================================================================================
' Name : OpenEmptName()
' Desc : developer describe this line
'========================================================================================================

Function OpenEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	End If

	arrParam(2) = lgUsrIntcd

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_EMP_NO
			frm1.vspdData.action = 0
		End If
		Exit Function
	Else
		Call SubSetCondEmp(arrRet, iWhere)
	End If

End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_DEPT_CD
			.vspdData.Text = arrRet(2)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action = 0
		End If
	    .txtEmp_no.focus
	End With
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If

	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("0000111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row

	If lgOldRow <> Row Then

		frm1.vspdData.Col = 1
		frm1.vspdData.Row = row

		lgOldRow = Row

		With frm1
		.vspdData.Row = .vspdData.ActiveRow

		.vspdData.Col = C_EMP_NO
		.txtEmp_no2.value = .vspdData.Text

		.vspdData.Col = C_HAA010T_NAME
		.txtName2.value = .vspdData.Text

		.vspdData.Col = C_TOT_PROV_AMT
		.txtTot_prov_amt.value = .vspdData.Text

		.vspdData.Col = C_DEPT_CD
		.txtDept_cd2.value = .vspdData.Text

		.vspdData.Col = C_PAY_ESTI_AMT
		.txtPay_esti_amt.value = .vspdData.Text

		.vspdData.Col = C_AVR_WAGES_AMT
		.txtAvr_wages_amt.value = .vspdData.Text

		.vspdData.Col = C_BONUS_ESTI_AMT
		.txtBonus_esti_amt.value = .vspdData.Text

		.vspdData.Col = C_RETIRE_AMT
		.txtRetire_amt.value = .vspdData.Text

		.vspdData.Col = C_YEAR_ESTI_AMT
		.txtYear_esti_amt.value = .vspdData.Text

		.vspdData.Col = C_DUTY_YY
		.txtDuty_yy.value = .vspdData.Text

		.vspdData.Col = C_DUTY_MM
		.txtDuty_mm.value = .vspdData.Text

		.vspdData.Col = C_DUTY_DD
		.txtDuty_dd.value = .vspdData.Text

		.vspdData.Col = C_TOT_DUTY_MM
		.txtTot_duty_mm.value = .vspdData.Text

		.vspdData.Col = C_RETIRE_YYMM
		.txtRetire_yymm_dt.Text = .txtpay_yymm_dt.Text

		End With
	End If	     
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
          gMouseClickStatus = "SPCR"
        End If
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
End Sub
'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc : This function is called when cursor leave cell
'========================================================================================================
Sub vspdData_ScriptLeaveCell(Col,Row,NewCol,NewRow,Cancel)
	If NewRow <= 0 Or NewCol < 0 Then
		Exit Sub
	End If

		frm1.vspdData.Col = 1
		frm1.vspdData.Row = NewRow

		With frm1

		.vspdData.Col = C_EMP_NO
		.txtEmp_no2.value = .vspdData.Text

		.vspdData.Col = C_HAA010T_NAME
		.txtName2.value = .vspdData.Text

		.vspdData.Col = C_TOT_PROV_AMT
		.txtTot_prov_amt.value = .vspdData.Text

		.vspdData.Col = C_DEPT_CD
		.txtDept_cd2.value = .vspdData.Text

		.vspdData.Col = C_PAY_ESTI_AMT
		.txtPay_esti_amt.value = .vspdData.Text

		.vspdData.Col = C_AVR_WAGES_AMT
		.txtAvr_wages_amt.value = .vspdData.Text

		.vspdData.Col = C_BONUS_ESTI_AMT
		.txtBonus_esti_amt.value = .vspdData.Text

		.vspdData.Col = C_RETIRE_AMT
		.txtRetire_amt.value = .vspdData.Text

		.vspdData.Col = C_YEAR_ESTI_AMT
		.txtYear_esti_amt.value = .vspdData.Text

		.vspdData.Col = C_DUTY_YY
		.txtDuty_yy.value = .vspdData.Text

		.vspdData.Col = C_DUTY_MM
		.txtDuty_mm.value = .vspdData.Text

		.vspdData.Col = C_DUTY_DD
		.txtDuty_dd.value = .vspdData.Text

		.vspdData.Col = C_TOT_DUTY_MM
		.txtTot_duty_mm.value = .vspdData.Text

		.vspdData.Col = C_RETIRE_YYMM
		.txtRetire_yymm_dt.Text = .txtpay_yymm_dt.Text
		End With
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
    	If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	

      	   Call DisableToolBar(parent.TBC_QUERY)
      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
      	   End If
    	End If
    End if
End Sub


'========================================================================================================
'   Event Name : txtEmp_no_Onchange
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    If  frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)

	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
            
            Call ggoOper.ClearField(Document, "2")
			frm1.txtName.value = ""
            txtEmp_no_Onchange = true
        Else
			frm1.txtName.value = strName
        End if

        Frm1.txtEmp_no.focus
    End if
End Function

'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtpay_yymm_dt_DblClick(Button)
    If Button = 1 Then
        Call SetFocusToDocument("M")
        frm1.txtpay_yymm_dt.Action = 7
        Frm1.txtpay_yymm_dt.Focus
    End If
End Sub
'==========================================================================================
'   Event Name : txtpay_yymm_dt_KeyDown()
'   Event Desc : 조회조건부의 txtpay_yymm_dt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtpay_yymm_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Function Nodata()
'=============================================================================================
frm1.txtEmp_no2.value = ""
frm1.txtTot_prov_amt.value = ""
frm1.txtName2.value = ""
frm1.txtPay_esti_amt.value = ""
frm1.txtDept_cd2.value = ""
frm1.txtBonus_esti_amt.value = ""
frm1.txtAvr_wages_amt.value = ""
frm1.txtYear_esti_amt.value = ""
frm1.txtRetire_amt.value = ""
frm1.txtDuty_yy.value = ""
frm1.txtDuty_mm.value = ""
frm1.txtDuty_dd.value = ""
frm1.txtRetire_yymm_dt.value = ""
frm1.txtTot_duty_mm.value = ""
'=============================================================================================
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>퇴직추계액조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

    <TR HEIGHT=*>
		<TD width=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR>
			        <TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>정산년월</TD>
								<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/ha203ma1_txtpay_yymm_dt_txtpay_yymm_dt.js'></script></TD>
							    <TD CLASS=TD5 NOWRAP>사원</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" MAXLENGTH="13" SIZE="13" ALT ="사번" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: openEmptName(0)">
								                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="성명" tag="14XXXU"></TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            		    <TR>
									        <TD HEIGHT="100%"><script language =javascript src='./js/ha203ma1_vaSpread_vspdData.js'></script></TD>
									    </TR>
					            	</TABLE>
								</TD>
							</TR>
            			    <TR HEIGHT=30%>
            			    	<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
	            					<FIELDSET CLASS="CMSFLD"><LEGEND ALGIN="LEFT">퇴직추계액</LEGEND>
	            					<TABLE WIDTH="100%" HEIGHT=* CELLSPACING=0>
	            					     <TR>
							                <TD CLASS="TD5" NOWRAP>사번</TD>
							                <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="Emp_no" NAME="txtEmp_no2" MAXLENGTH="13" SIZE=13 tag="14XXXU" ALT="사번"></TD>
              						        <TD CLASS="TD5" NOWRAP>총지급액</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha203ma1_fpDoubleSingle2_txtTot_prov_amt.js'></script></TD>
              						    </TR>
              						    <TR>
              						        <TD CLASS="TD5" NOWRAP>성명</TD>
							                <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="HAA010T_NAME" NAME="txtName2" MAXLENGTH="30" SIZE=20 tag="14XXXU" ALT="성명"></TD>
                                            <TD CLASS="TD5" NOWRAP>급여추계총액</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha203ma1_fpDoubleSingle2_txtPay_esti_amt.js'></script></TD>
	                   					</TR>
              						    <TR>
              						        <TD CLASS="TD5" NOWRAP>부서</TD>
							                <TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID="Dept_cd" NAME="txtDept_cd2" MAXLENGTH="40" SIZE=20 tag="14XXXU" ALT="부서"></TD>
              						        <TD CLASS="TD5" NOWRAP>상여추계총액</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha203ma1_Bonus_esti_amt_txtBonus_esti_amt.js'></script></TD>
              						    </TR>
              						    <TR>
              						        <TD CLASS="TD5" NOWRAP>평균임금</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha203ma1_Avr_wages_amt_txtAvr_wages_amt.js'></script></TD>
              						        <TD CLASS="TD5" NOWRAP>연월차추계총액</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha203ma1_Year_esti_amt_txtYear_esti_amt.js'></script></TD>
              						    </TR>
              						    <TR>
              						        <TD CLASS="TD5" NOWRAP>퇴직금</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha203ma1_Retire_amt_txtRetire_amt.js'></script></TD>
              						        <TD CLASS="TD5" NOWRAP>근속</TD>
	                   						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ID=Duty_yy NAME=txtDuty_yy SIZE=3 STYLE="TEXT-ALIGN: RIGHT" tag="14" ALT="근속년">년
	                   						                       <INPUT TYPE=TEXT ID=Duty_mm NAME=txtDuty_mm SIZE=3 STYLE="TEXT-ALIGN: RIGHT" tag="14" ALT="근속월">개월
	                   						                       <INPUT TYPE=TEXT ID=Duty_dd NAME=txtDuty_dd SIZE=3 STYLE="TEXT-ALIGN: RIGHT" tag="14" ALT="근속일">일</TD>
              						    </TR>
              						    <TR>
              						        <TD CLASS="TD5" NOWRAP>정산년월</TD>
								            <TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha203ma1_Retire_yymm_dt_txtRetire_yymm_dt.js'></script></TD>
              						        <TD CLASS="TD5" NOWRAP>총근속개월수</TD>
	                   						<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/ha203ma1_txtTot_duty_mm_txtTot_duty_mm.js'></script>개월</TD>					            		</TR>
						            </TABLE>
						            </fieldset>
					            </TD>
                            </TR>
  						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</H
