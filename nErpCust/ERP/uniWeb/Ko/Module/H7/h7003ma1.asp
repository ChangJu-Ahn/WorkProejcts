<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 개인상여조정율등록 
*  3. Program ID           : H7003ma1
*  4. Program Name         : H7003ma1
*  5. Program Desc         : 상여관리/개인상여조정율등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/17
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : YBI
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID = "H7003mb1.asp"                                      'Biz Logic ASP
Const BIZ_PGM_JUMP_ID  = "H7003ba1"
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

Dim C_EMP_NO
Dim C_EMP_NO_POP
Dim C_NAME
Dim C_BONUS_RATE
Dim C_ADD_RATE
Dim C_MINUS1_RATE
Dim C_MINUS2_RATE
Dim C_PROV_RATE
Dim C_PROV_AMT
Dim C_GRAND_AMT
Dim C_GRAND_RATE

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

    C_EMP_NO = 1
    C_EMP_NO_POP = 2
    C_NAME = 3
    C_BONUS_RATE = 4
    C_ADD_RATE = 5
    C_MINUS1_RATE = 6
    C_MINUS2_RATE = 7
    C_PROV_RATE = 8
    C_PROV_AMT = 9
    C_GRAND_AMT = 10
    C_GRAND_RATE = 11

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	frm1.txtBonus_yymm_dt.focus

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtBonus_yymm_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtBonus_yymm_dt.Month = strMonth 
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
    On Error Resume Next
    Const CookieSplit = 4877	
	If flgs = 1 Then   
	    WriteCookie "BONUS_YYMM_DT" , frm1.txtBonus_yymm_dt.Text
	    WriteCookie "BONUS_TYPE" , frm1.txtBonus_type.Value
	End If
End Function

Function PgmJumpCheck()         
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
	   
    PgmJump(BIZ_PGM_JUMP_ID)
	    
End Function	  

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
	Dim strYYYY
	Dim strMM
	Dim strBonus_yymm_dt

    strYYYY = Frm1.txtBonus_yymm_dt.Year
    strMM = Frm1.txtBonus_yymm_dt.Month
    If len(strMM) = 1 Then
		strMM = "0" & strMM
	End if
    strBonus_yymm_dt = strYYYY & strMM

    lgKeyStream = Frm1.txtBonus_type.Value & Parent.gColSep                'You Must append one character(Parent.gColSep)
    lgKeyStream = lgKeyStream & strBonus_yymm_dt & Parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtemp_no.value & Parent.gColSep
    lgKeyStream = lgKeyStream & lgUsrIntCd & Parent.gColSep
End Sub        


'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0040", "''", "S") & " and ((minor_cd >= " & FilterVar("2", "''", "S") & " and minor_cd <= " & FilterVar("9", "''", "S") & ") or minor_cd=" & FilterVar("C", "''", "S") & "  or minor_cd=" & FilterVar("Q", "''", "S") & " ) ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = lgF0
    iNameArr = lgF1
    Call SetCombo2(frm1.txtBonus_type,iCodeArr, iNameArr,Chr(11))
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

	Call initSpreadPosVariables()   'sbk 

	With frm1.vspdData
        ggoSpread.Source = frm1.vspdData
	
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	    .ReDraw = false
	
        .MaxCols = C_GRAND_RATE + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
    
        .MaxRows = 0
        ggoSpread.ClearSpreadData

        Call GetSpreadColumnPos("A") 'sbk
	
	    Call AppendNumberPlace("6","3","2")
	    Call AppendNumberPlace("7","4","2")

        ggoSpread.SSSetEdit     C_EMP_NO,       "사번",          13,,,13
        ggoSpread.SSSetButton   C_EMP_NO_POP
        ggoSpread.SSSetEdit     C_NAME,         "성명",          10,,,10
        ggoSpread.SSSetFloat    C_BONUS_RATE,   "상여율",        09,"7",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
        ggoSpread.SSSetFloat    C_ADD_RATE,     "가산율",        09,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_MINUS1_RATE,  "근태차감율",    12,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_MINUS2_RATE,  "일할계산차감율",15,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_PROV_RATE,    "실지급율",      11,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_PROV_AMT,     "지급액",        13,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_GRAND_AMT,    "생산장려금",    12,Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
        ggoSpread.SSSetFloat    C_GRAND_RATE,   "생산장려율",    12,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec

        Call ggoSpread.MakePairsColumn(C_EMP_NO,C_EMP_NO_POP)    'sbk

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

    ggoSpread.SpreadLock        C_NAME, -1, C_NAME, -1
    ggoSpread.SpreadLock        C_EMP_NO_POP, -1, C_EMP_NO_POP, -1
    ggoSpread.SpreadLock        C_EMP_NO, -1, C_EMP_NO, -1

    ggoSpread.SSSetRequired		C_BONUS_RATE, -1, -1
    ggoSpread.SSSetRequired		C_ADD_RATE, -1, -1
    ggoSpread.SSSetRequired		C_MINUS1_RATE, -1, -1
    ggoSpread.SSSetRequired		C_MINUS2_RATE, -1, -1
    ggoSpread.SpreadLock		    C_PROV_RATE, -1, C_PROV_RATE, -1
    ggoSpread.SSSetRequired		C_PROV_AMT, -1, -1
    ggoSpread.SSSetRequired		C_GRAND_AMT, -1, -1
    ggoSpread.SSSetRequired		C_GRAND_RATE, -1, -1
    ggoSpread.SSSetProtected   .vspdData.MaxCols   , -1, -1

    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetProtected	    C_NAME, pvStartRow, pvEndRow

    ggoSpread.SSSetRequired		C_EMP_NO, pvStartRow, pvEndRow

    ggoSpread.SSSetRequired		C_BONUS_RATE, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_ADD_RATE, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_MINUS1_RATE, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_MINUS2_RATE, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	    C_PROV_RATE, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_PROV_AMT, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_GRAND_AMT, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_GRAND_RATE, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
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

            C_EMP_NO = iCurColumnPos(1)
            C_EMP_NO_POP = iCurColumnPos(2)
            C_NAME = iCurColumnPos(3)
            C_BONUS_RATE = iCurColumnPos(4)
            C_ADD_RATE = iCurColumnPos(5)
            C_MINUS1_RATE = iCurColumnPos(6)
            C_MINUS2_RATE = iCurColumnPos(7)
            C_PROV_RATE = iCurColumnPos(8)
            C_PROV_AMT = iCurColumnPos(9)
            C_GRAND_AMT = iCurColumnPos(10)
            C_GRAND_RATE = iCurColumnPos(11)
    End Select    
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
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call ggoOper.FormatDate(frm1.txtBonus_yymm_dt, Parent.gDateFormat, 2)
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call FuncGetAuth("H7003MA1", Parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")

    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어 
    
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
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.ClearSpreadData
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    if txtEmp_no_Onchange() then
		Exit Function
	end if
    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

	Call DisableToolBar(Parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
       Exit Function
    End If                                                                 '☜: Query db data
       
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
    Dim IntRetCD ,lRow
    
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
	With Frm1
       For lRow = 1 To .vspdData.MaxRows
           .vspdData.Row = lRow
           .vspdData.Col = 0
           if   .vspdData.Text = ggoSpread.InsertFlag OR .vspdData.Text = ggoSpread.UpdateFlag then
					.vspdData.Col = C_NAME

					If IsNull(Trim(.vspdData.Text)) OR Trim(.vspdData.Text) = "" Then
					    Call DisplayMsgBox("800048","X","X","X")
						Exit Function
					end if
            end if
        next
    end with
	
    Call MakeKeyStream("X")

	Call DisableToolBar(Parent.TBC_SAVE)
    If DbSave = False Then
		Call RestoreToolBar()
       Exit Function
    End If				                                                    '☜: Save db data
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
            SetSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	With Frm1.VspdData
           .Col  = C_NAME
           .Row  = .ActiveRow
           .Text = ""

           .Col  = C_EMP_NO
           .Row  = .ActiveRow
           .Text = ""
    End With

    Set gActiveElement = document.ActiveElement   

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
Function FncInsertRow(ByVal pvRowCnt) 
    Dim imRow, iCnt

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
 
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
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
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

        For iCnt = 1 To imRow 
            .vspdData.Row = .vspdData.ActiveRow + iCnt - 1

            .vspdData.Col = C_BONUS_RATE
            .vspdData.Text = 0
            .vspdData.Col = C_ADD_RATE
            .vspdData.Text = 0
            .vspdData.Col = C_MINUS1_RATE
            .vspdData.Text = 0
            .vspdData.Col = C_MINUS2_RATE
            .vspdData.Text = 0
            .vspdData.Col = C_PROV_RATE
            .vspdData.Text = 0
            .vspdData.Col = C_PROV_AMT
            .vspdData.Text = 0

            .vspdData.Col = C_GRAND_AMT
            .vspdData.Text = 0
            .vspdData.Col = C_GRAND_RATE
            .vspdData.Text = 0
        Next
        
       .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
    
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
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                                        '☜: Clear err status

	If LayerShowHide(1) = False then
    		Exit Function 
    End if
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
    
    DbQuery = True
    
End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel

    Dim dblBonus_rate
    Dim dblAdd_rate
    Dim dblMinus1_rate
    Dim dblMinus2_rate
    
	Dim strYYYY
	Dim strMM
	Dim strBonus_yymm_dt

    strYYYY = Frm1.txtBonus_yymm_dt.Year
    strMM = Frm1.txtBonus_yymm_dt.Month
    If len(strMM) = 1 Then
		strMM = "0" & strMM
	End if
	
	strBonus_yymm_dt = UniConvDateToYYYYMM(Frm1.txtBonus_yymm_dt.Text, Parent.gDateFormatYYYYMM, "")
    
	With Frm1

       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0

           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag                                      '☜: Update

                    .vspdData.Col = C_BONUS_RATE
                    dblBonus_rate = UNICDbl(.vspdData.Text)
                            
                    .vspdData.Col = C_ADD_RATE
                    dblAdd_rate = UNICDbl(.vspdData.Text)
                            
                    .vspdData.Col = C_MINUS1_RATE
                    dblMinus1_rate = UNICDbl(.vspdData.Text)

                    .vspdData.Col = C_MINUS2_RATE
                    dblMinus2_rate = UNICDbl(.vspdData.Text)
                            
                    .vspdData.Col = C_PROV_RATE
                    .vspdData.value = (dblBonus_rate + dblAdd_rate) - (dblMinus1_rate + dblMinus2_rate)
					'.vspdData.Text = UNIFormatNumber((dblBonus_rate + dblAdd_rate) - (dblMinus1_rate + dblMinus2_rate),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
           End Select
       Next

	End With

    DbSave = False                                                          
    
	If LayerShowHide(1) = False then
    		Exit Function 
    End if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                   strVal = strVal & "C" & Parent.gColSep
                                                   strVal = strVal & lRow & Parent.gColSep
                                                   strVal = strVal & .txtBonus_type.Value & Parent.gColSep
                                                   strVal = strVal & strBonus_yymm_dt & Parent.gColSep
                   .vspdData.Col = C_EMP_NO 	 : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_BONUS_RATE	 : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_ADD_RATE	 : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_MINUS1_RATE : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_MINUS2_RATE : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_PROV_RATE   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_PROV_AMT    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_GRAND_AMT   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_GRAND_RATE  : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                   strVal = strVal & "U" & Parent.gColSep
                                                   strVal = strVal & lRow & Parent.gColSep
                                                   strVal = strVal & .txtBonus_type.Value & Parent.gColSep
                                                   strVal = strVal & strBonus_yymm_dt & Parent.gColSep
                   .vspdData.Col = C_EMP_NO 	 : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_BONUS_RATE	 : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_ADD_RATE	 : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_MINUS1_RATE : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_MINUS2_RATE : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_PROV_RATE   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_PROV_AMT    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_GRAND_AMT   : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                   .vspdData.Col = C_GRAND_RATE  : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                   strDel = strDel & "D" & Parent.gColSep
                                                   strDel = strDel & lRow & Parent.gColSep
                                                   strDel = strDel & .txtBonus_type.Value & Parent.gColSep
                                                   strDel = strDel & strBonus_yymm_dt & Parent.gColSep
                   .vspdData.Col = C_EMP_NO 	 : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        = Parent.UID_M0002
       .txtUpdtUserId.value  = Parent.gUsrID
       .txtInsrtUserId.value = Parent.gUsrID
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
    DbSave = True                                                           
    
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
	Call DisableToolBar(Parent.TBC_DELETE)
    If DbDelete= False Then
		Call RestoreToolBar()
        Exit Function
    End If
    
    FncDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = Parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("110011110011111")									
	frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData
    ggoSpread.ClearSpreadData
    
    Call InitVariables															'⊙: Initializes local global variables

	Call MainQuery()
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================

Function OpenEmp(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	Else 'spread
	    frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
	End If
	arrParam(1) = ""'frm1.txtName.value			' Name Cindition
    arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_EMP_NO
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SetEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If	
	End With
End Sub

Sub btnAuto_proc_ButtonClicked(Byval ButtonDown)

    Dim arrParam(1)
    Dim arrRet
	arrRet = window.showModalDialog("H7003ma2.asp", Array(window.parent, arrParam), _
		"dialogWidth=460px; dialogHeight=220px; center: Yes; help: No; resizable: No; status: No;")

End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
			    Case C_EMP_NO_POP
   	                frm1.vspdData.Row = Row
                    frm1.vspdData.Col = Col
                    Call OpenEmp(1)
			End Select
		End If
    
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim dblBonus_rate
    Dim dblAdd_rate
    Dim dblMinus1_rate
    Dim dblMinus2_rate
    Dim IntRetCd

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
        Case C_BONUS_RATE, C_ADD_RATE, C_MINUS1_RATE, C_MINUS2_RATE
            Frm1.vspdData.Col = C_BONUS_RATE
            dblBonus_rate = UNICDbl(Frm1.vspdData.Text)
           
            Frm1.vspdData.Col = C_ADD_RATE
            dblAdd_rate = UNICDbl(Frm1.vspdData.Text)
            
            Frm1.vspdData.Col = C_MINUS1_RATE
            dblMinus1_rate = UNICDbl(Frm1.vspdData.Text)
			
            Frm1.vspdData.Col = C_MINUS2_RATE
            dblMinus2_rate = UNICDbl(Frm1.vspdData.Text)

            Frm1.vspdData.Col = C_PROV_RATE
            'Frm1.vspdData.Text = UNIFormatNumber((dblBonus_rate + dblAdd_rate) - (dblMinus1_rate + dblMinus2_rate),ggAmtOfMoney.DecPoint,-2,0,ggAmtOfMoney.RndPolicy,ggAmtOfMoney.RndUnit)
            Frm1.vspdData.value = (dblBonus_rate + dblAdd_rate) - (dblMinus1_rate + dblMinus2_rate)
            
        Case C_EMP_NO
            IntRetCd = CommonQueryRs(" NAME "," HAA010T "," EMP_NO =  " & FilterVar(Frm1.vspdData.Text , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
            If  IntRetCd = false then
		    	Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
                Frm1.vspdData.Col = C_NAME
		    	Frm1.vspdData.Text = ""
                Frm1.vspdData.Action = 0 ' go to 
                vspdData_Change = true
            Else
                Frm1.vspdData.Col = C_NAME
		    	Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
            End if 

    End Select    
             
   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Function

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("1101111111")

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
     End If
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
'======================================================================================================
'   Event Name : txtBonus_yymm_dt_DblClick
'   Event Desc : 달력 Popup을 호출 
'=======================================================================================================
Sub txtBonus_yymm_dt_DblClick(Button)
	If Button = 1 Then
		Call SetFocusToDocument("M")	
		frm1.txtBonus_yymm_dt.Action = 7
		frm1.txtBonus_yymm_dt.focus
	End If
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
    Dim strVal

	frm1.txtName.value = ""
    If  frm1.txtEmp_no.value = "" Then
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if
    
End Function

'==========================================================================================
'   Event Name : txtBonus_yymm_dt_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtBonus_yymm_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>개인상여조정율등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
							        <TD CLASS="TD5" NOWRAP>상여구분</TD>
							        <TD CLASS="TD6" NOWRAP><SELECT NAME="txtBonus_type" ALT="상여구분" CLASS ="cbonormal" TAG="12"></SELECT></TD>
							    	<TD CLASS=TD5>상여년월</TD>
							    	<TD CLASS=TD6><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtBonus_yymm_dt name=txtBonus_yymm_dt CLASS=FPDTYYYYMM title=FPDATETIME ALT="상여년월" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							    </TR>
							    <TR>	
							    	<TD CLASS=TD5>사원</TD>
							    	<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=13 MAXLENGTH=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmp('0')">
							    						 <INPUT NAME="txtName" ALT="성명" TYPE="Text" SiZE=20 MAXLENGTH=30 tag=14></TD>
							    	<TD CLASS=TD5></TD>
							    	<TD CLASS=TD6></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJumpCheck()" ONCLICK="VBSCRIPT:CookiePage 1">일괄적용</a>
				</TD>
				<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
	    </TD>   
	
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

