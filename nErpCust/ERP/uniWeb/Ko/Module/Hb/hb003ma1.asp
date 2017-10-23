<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 가족사항사항등록 
*  3. Program ID           : H2003ma1
*  4. Program Name         : H2003ma1
*  5. Program Desc         : 인사기본자료관리/가족사항등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/10
*  8. Modified date(Last)  : 2003/06/10
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
const	 CookieSplit = 1233
const	 C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>
Const BIZ_PGM_ID = "hb003mb1.asp"                                      'Biz Logic ASP
Const BIZ_PGM_REF_ID = "hb003mb2.asp" 
Const BIZ_PGM_CALC_ID = "hb003mb3.asp" 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          

Dim gUsrAuth    ' 자료권한관리 

Dim C_CHK
Dim C_EMP_NO
Dim C_EMP_NO_POP
Dim C_EMP_NM

Dim C_DEPT_CD
Dim C_DEPT_NM 

Dim C_DUTY_DAY 
Dim C_DAY_MONEY
Dim C_PROV_TOT_AMT
 
Dim C_SUB_TOT_AMT 
Dim C_INCOME_TAX
Dim C_RES_TAX
Dim C_REAL_PROV_AMT

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

	C_CHK			= 1
	C_EMP_NO		= 2
	C_EMP_NO_POP	= 3
	C_EMP_NM		= 4
	C_DEPT_CD		= 5
	C_DEPT_NM		= 6
	C_DUTY_DAY 		= 7
	C_DAY_MONEY		= 8
	C_PROV_TOT_AMT	= 9
		 
	C_SUB_TOT_AMT	= 10
	C_INCOME_TAX	= 11
	C_RES_TAX		= 12
	C_REAL_PROV_AMT	= 13

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
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %>
End Sub
'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit = 4877						
	Dim strTemp

	If flgs = 1 Then
		WriteCookie CookieSplit , frm1.txtEmp_no.Value
	ElseIf flgs = 0 Then

		strTemp = ReadCookie(CookieSplit)
		If strTemp = "" then Exit Function
			
		frm1.txtEmp_no.value =  strTemp

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		WriteCookie CookieSplit , ""
		
	Call MainQuery()
			
	End If
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream       = Replace(Frm1.txtPAY_YYMM.Text, "-", "") & parent.gColSep                                           'You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & Frm1.txtEmp_no.Value & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtDept_Cd.Value & parent.gColSep
End Sub        

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	frm1.txtPAY_YYMM.Year = strYear 
	frm1.txtPAY_YYMM.Month = strMonth 

End Sub


'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Dim strMaskYM
	strMaskYM = "999999-9999999"
	Call initSpreadPosVariables()  
	With frm1.vspdData

		Call AppendNumberPlace("6","6","0")
		
        ggoSpread.Source = frm1.vspdData
	
		ggoSpread.Spreadinit "V20021125",,parent.gForbidDragDropSpread    
	    .ReDraw = false    
        .MaxCols = C_REAL_PROV_AMT + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True	  
        .MaxRows = 0 
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData     

		Call GetSpreadColumnPos("A")  

		ggoSpread.SSSetCheck	C_CHK,			"급여계산", 10,,, 2
        ggoSpread.SSSetEdit		C_EMP_NO,		"사번", 13,,, 13,1
        ggoSpread.SSSetButton	C_EMP_NO_POP
        ggoSpread.SSSetEdit		C_EMP_NM,		"성명", 15,,, 30,1
        ggoSpread.SSSetEdit     C_DEPT_CD,		"부서", 10,,, 10,1
        ggoSpread.SSSetEdit     C_DEPT_NM,		"부서명", 15,,, 50,1

        ggoSpread.SSSetFloat	C_DUTY_DAY,		"근로일수", 10,		"6", ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec ,,,"Z" 
		ggoSpread.SSSetFloat	C_DAY_MONEY,	"일당금액", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  ,,,"Z"
		ggoSpread.SSSetFloat	C_PROV_TOT_AMT,	"총지급액", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  ,,,"Z"

		ggoSpread.SSSetFloat	C_SUB_TOT_AMT,	"총공제액", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  ,,,"Z"
		ggoSpread.SSSetFloat	C_INCOME_TAX,	"소득세", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  ,,,"Z"
		ggoSpread.SSSetFloat	C_RES_TAX,		"주민세", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  ,,,"Z"
		ggoSpread.SSSetFloat	C_REAL_PROV_AMT,"실지급액", 15,		Parent.ggAmtOfMoneyNo,		ggStrIntegeralPart,		ggStrDeciPointPart,		Parent.gComNum1000,		Parent.gComNumDec  ,,,"Z"

	   .ReDraw = true
	
       Call SetSpreadLock 
    
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

			C_CHK			= iCurColumnPos(1)
			C_EMP_NO		= iCurColumnPos(2)
			C_EMP_NO_POP	= iCurColumnPos(3)
			C_EMP_NM		= iCurColumnPos(4)

			C_DEPT_CD		= iCurColumnPos(5)
			C_DEPT_NM		= iCurColumnPos(6)
						
			C_DUTY_DAY 		= iCurColumnPos(7)
			C_DAY_MONEY		= iCurColumnPos(8)
			C_PROV_TOT_AMT	= iCurColumnPos(9)
					 
			C_SUB_TOT_AMT	= iCurColumnPos(10)
			C_INCOME_TAX	= iCurColumnPos(11)
			C_RES_TAX		= iCurColumnPos(12)
			C_REAL_PROV_AMT	= iCurColumnPos(13)
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
	Dim iMaxRows
	
    With frm1
    
    iMaxRows = .vspdData.MaxRows
    
    .vspdData.ReDraw = False
	ggoSpread.SpreadLock		C_EMP_NO, -1, C_EMP_NM
	ggoSpread.SpreadLock		C_DEPT_CD, -1, C_DEPT_NM
	ggoSpread.SSSetRequired		C_DUTY_DAY, 1, iMaxRows
	ggoSpread.SpreadLock		C_DAY_MONEY, -1, C_DAY_MONEY
	ggoSpread.SpreadLock		C_PROV_TOT_AMT, -1, C_PROV_TOT_AMT
	ggoSpread.SpreadLock		C_REAL_PROV_AMT, -1, C_REAL_PROV_AMT
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
    
    ggoSpread.SSSetRequired		C_EMP_NO, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_EMP_NM, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_DEPT_CD, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_DEPT_NM, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_DUTY_DAY, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_DAY_MONEY, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_PROV_TOT_AMT, pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_REAL_PROV_AMT, pvStartRow, pvEndRow
    
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

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
    Call ggoOper.FormatDate(frm1.txtPAY_YYMM, parent.gDateFormat,2)
        
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
	Call SetDefaultVal
	
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetToolbar("1100110100101111")										        '버튼 툴바 제어    
    frm1.txtPAY_YYMM.Focus
	Call CookiePage(0)                                                             '☜: Check Cookie
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
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	ggoSpread.ClearSpreadData    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
		Call RestoreToolBar()
        Exit Function
    End If

    FncQuery = True                                                              '☜: Processing is OK
    
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
    Dim lRow
    Dim strAdmi_dt
    Dim strGrudt_dt
    Dim strNat_cd
    Dim cnt00   ' 본인 
    Dim cnt01   ' 조부 
    Dim cnt02   ' 조모 
    Dim cnt03   ' 부 
    Dim cnt04   ' 모 
    Dim cnt09   ' 배우자 

    Dim res_no1, res_no2            ' 주민번호 
    Dim intChk, intMod, intDef      ' 주민번호 

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
    
    Call MakeKeyStream("X")
	Call DisableToolBar(parent.TBC_SAVE)
    If DbSave = False Then
		Call RestoreToolBar()
        Exit Function
    End If    
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

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
    Dim IntRetCD,imRow
    
    On Error Resume Next         
    FncInsertRow = False
    
    if IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
	End if
	With frm1
	    .vspdData.ReDraw = False
	    .vspdData.focus
	    ggoSpread.Source = .vspdData
	    ggoSpread.InsertRow,imRow
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
	   .vspdData.ReDraw = True
	End With
	Set gActiveElement = document.ActiveElement   
	If Err.number =0 Then
		FncInsertRow = True
	End if
	
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
End Sub
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
    Call parent.FncExport(parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
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
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	Dim strVal
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
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
    Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel

	Dim strRes_no

    DbSave = False                                                          
    
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                    strVal = strVal & "C" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO		: strVal = strVal & Trim(.vspdData.Text)  & parent.gColSep
                    .vspdData.Col = C_DUTY_DAY      : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_PROV_TOT_AMT	: strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_SUB_TOT_AMT	: strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_INCOME_TAX	: strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_RES_TAX		: strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_REAL_PROV_AMT	: strVal = strVal & Trim(.vspdData.value) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                    strVal = strVal & "U" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO		: strVal = strVal & Trim(.vspdData.Text)  & parent.gColSep
                    .vspdData.Col = C_DUTY_DAY      : strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_PROV_TOT_AMT	: strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_SUB_TOT_AMT	: strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_INCOME_TAX	: strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_RES_TAX		: strVal = strVal & Trim(.vspdData.value) & parent.gColSep
                    .vspdData.Col = C_REAL_PROV_AMT	: strVal = strVal & Trim(.vspdData.value) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	  : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep

                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
		.txtKeyStream.value = lgKeyStream
       .txtMode.value        = parent.UID_M0002
       .txtUpdtUserId.value  = parent.gUsrID
       .txtInsrtUserId.value = parent.gUsrID
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

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
	
	Call DisableToolBar(parent.TBC_DELETE)
    If DbDelete = False Then
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
	
	Dim strVal

    lgIntFlgMode = parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
	Call SetToolbar("110011110011111")
	Call SetSpreadLock
	Frm1.vspdData.focus	
End Function

Function DbRefOk()													     
	
	Dim strVal, i, iLen

	With frm1.vspdData 

	iLen = .MaxRows
	
	For i = 1 To iLen
		.Row = i
		.Col = 0
		.Value = ggoSpread.InsertFlag
	Next

	End With
	
	lgBlnFlgChgValue  = True
    lgIntFlgMode = parent.OPMD_CMODE
	Call SetToolbar("110011010011111")
	Call SetSpreadLock
	Frm1.vspdData.focus	
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
   	
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
       frm1.vspdData.Row = Row
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
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim duty_day, day_money
    Dim sub_tot_amt, prov_tot_amt, income_tax, res_tax
    Dim IntRetCd
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_DUTY_DAY
			duty_day = cdbl(Frm1.vspdData.value)

            Frm1.vspdData.Col = C_DAY_MONEY
            day_money = cdbl(Frm1.vspdData.value)

            prov_tot_amt = duty_day * day_money

            Frm1.vspdData.Col = C_PROV_TOT_AMT
            Frm1.vspdData.text = prov_tot_amt

            Frm1.vspdData.Col = C_SUB_TOT_AMT
			sub_tot_amt =  Frm1.vspdData.text

            Frm1.vspdData.Col = C_INCOME_TAX
            income_tax = Frm1.vspdData.text

            Frm1.vspdData.Col = C_RES_TAX
            res_tax = Frm1.vspdData.text
            
            Frm1.vspdData.Col = C_REAL_PROV_AMT
            Frm1.vspdData.text = prov_tot_amt - sub_tot_amt - income_tax - res_tax

         Case  C_SUB_TOT_AMT,C_INCOME_TAX, C_RES_TAX
            Frm1.vspdData.Col = C_SUB_TOT_AMT
			sub_tot_amt =  Frm1.vspdData.text
			
            Frm1.vspdData.Col = C_PROV_TOT_AMT
            prov_tot_amt = Frm1.vspdData.text
			
            Frm1.vspdData.Col = C_INCOME_TAX
            income_tax = Frm1.vspdData.text
            
            Frm1.vspdData.Col = C_RES_TAX
            res_tax = Frm1.vspdData.text
            
            Frm1.vspdData.Col = C_REAL_PROV_AMT
            Frm1.vspdData.text = prov_tot_amt - sub_tot_amt - income_tax - res_tax
            
         Case  C_EMP_NO
   
            If Frm1.vspdData.value = "" Then
  	            Frm1.vspdData.Col = C_EMP_NM
                Frm1.vspdData.value = ""
            Else
				IntRetCd = CommonQueryRs(" emp_nm "," HAA011T "," emp_no =  " & FilterVar(Frm1.vspdData.value , "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
                If IntRetCd = false then
			        Call DisplayMsgBox("800048","X","X","X")
  	            Frm1.vspdData.Col = C_EMP_NM
                Frm1.vspdData.value = ""			        
                Else
		       	    Frm1.vspdData.Col = C_EMP_NM
		       	    Frm1.vspdData.value = Trim(Replace(lgF0,Chr(11),""))
                End if 
            End if 
    End Select    

   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼버튼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col

	Select Case Col
		Case C_EMP_NO_POP
			Call OpenEmp(1)
	End Select
'	ggoSpread.Source = frm1.vspdData
'	ggoSpread.UpdateRow Row
    
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


Function getRef()

	Dim IntRetCD
	
    ggoSpread.Source = Frm1.vspdData

    If ggoSpread.SSCheckChange = True Or lgBlnFlgChgValue  = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	ggoSpread.ClearSpreadData    															

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	Dim strVal

	Call MakeKeyStream("")
    
    With Frm1
		strVal = BIZ_PGM_REF_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		

	Call RunMyBizASP(MyBizASP, strVal)                                               '☜: Run Biz Logic
	
End Function



'========================================================================================================= 
Sub fnBttnConf()	
	Dim intRow
	Dim chk
	
	With frm1.vspdData	
	
		chk = "1"
		For intRow = 1 To .MaxRows			
	   		.Row = intRow
			.Col = C_CHK
			If .Text = "1" Then
				chk = "0"
				ggoSpread.Source = frm1.vspdData
				ggoSpread.UpdateRow intRow					
	'					Exit For
			End If
		Next		
			
		For intRow = 1 To .MaxRows	
	   		.Row = intRow			
	 
			If 1=1 Then

				.Col = C_CHK
				.Text = chk		
				ggoSpread.Source = frm1.vspdData					
				ggoSpread.UpdateRow intRow				
			End If	
		Next
	End With		
End Sub

' --- 급여계산 -----
Function fncCalcPay()
	Dim IntRetCD
	
    ggoSpread.Source = Frm1.vspdData

    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("189217", "X", "X", "X")
        Exit Function
    End If

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
	
	Call MakeKeyStream("")

	Dim strVal, strDel, iLen, lRow
    strVal = ""

	With frm1.vspdData
	
		iLen = .MaxRows
    
		For lRow = 1 To iLen
    
           .Row = lRow
           .Col = C_CHK
        
           If .Value = "1" Then
				.Col = C_EMP_NO	  : strVal = strVal & FilterVar(.Text, "''", "S") & ","
			End If
		Next	
		
		If strVal = "" Then
			Call LayerShowHide(0)
			Call DisplayMsgBox("900025", parent.VB_INFORMATION,"x","x")					
			Exit Function
		End If
		
		frm1.txtKeyStream.value = lgKeyStream
		frm1.txtMode.value        = parent.UID_M0002
		frm1.txtUpdtUserId.value  = parent.gUsrID
		frm1.txtInsrtUserId.value = parent.gUsrID
		frm1.txtSpread.value		 = Left(strVal, Len(strVal)-1)

	End With
	
	
	Call ExecMyBizASP(frm1, BIZ_PGM_CALC_ID)	
    
End Function

'========================================================================================================
'	Name : OpenEmp()
'========================================================================================================
Function OpenEmp(iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "일용직 사원팝업"			' 팝업 명칭 
	arrParam(1) = "HAA011T"						' TABLE 명칭 

	If iWhere = 0 Then	'TextBox(Condition)
		arrParam(2) = UCase(Trim(frm1.txtEmp_no.value))			' Code Condition
	Else 'spread
		arrParam(2) = frm1.vspdData.Text						' Code Condition
	End If
	
	arrParam(3) = ""					' Name Cindition
	arrParam(4) = ""					' Where Condition%>
	arrParam(5) = "사번"			' 조건필드의 라벨 명칭 
	
    arrField(0) = "emp_no"					' Field명(0)
	arrField(1) = "emp_nm"					' Field명(1)
	    
    arrHeader(0) = "사번"		' Header명(0)
    arrHeader(1) = "이름"		' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 'TextBox(Condition)
			frm1.txtEmp_no.focus
		Else 'spread
			frm1.vspdData.Col = C_EMP_NO
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		With frm1
			If iWhere = 0 Then 'TextBox(Condition)
				.txtEmp_no.value = arrRet(0)
				.txtName.value = arrRet(1)
				.txtEmp_no.focus
			Else 'spread
				.vspdData.Col = C_EMP_NO
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_EMP_NM
				.vspdData.Text = arrRet(1)

				.vspdData.Col = C_EMP_NO
				.vspdData.action =0
			End If
		End With
	End If	
	
End Function

'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtDept_cd.value			            '  Code Condition
   	arrParam(1) = ""
	arrParam(2) = lgUsrIntCd                            ' 자료권한 Condition  

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDept_cd.focus
		Exit Function
	Else
		With frm1
		    .txtDept_cd.value = arrRet(0)
		    .txtDept_nm.value = arrRet(1)
		    .txtDept_cd.focus
		End With
	End If
End Function

'========================================================================================================
'   Event Name : txtEmp_no_Onchange
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    
    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
    Else
        IntRetCd = CommonQueryRs(" emp_nm "," HAA011T "," emp_no =  " & FilterVar(frm1.txtEmp_no.value , "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call DisplayMsgBox("800048","X","X","X")	 
		    frm1.txtName.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement 
            
            txtEmp_no_Onchange = true
            Exit Function      
        Else
			frm1.txtName.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
    
End Function

Function txtDept_cd_OnChange()
    Dim IntRetCd
    Dim strDept_nm
    Dim strInternal_cd
    
    txtDept_cd_OnChange = true

    If RTrim(frm1.txtDept_cd.value) = "" Then
        frm1.txtDept_nm.value = ""
        frm1.txtDept_cd.focus()
    Else
        IntRetCd =  FuncDeptName(frm1.txtDept_cd.value,"",lgUsrIntCd,strDept_nm,strInternal_cd)
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call  DisplayMsgBox("970000", "x","부서코드","x")
            else
                Call  DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
            frm1.txtDept_nm.value = ""
            frm1.txtDept_cd.focus()
            Set gActiveElement = document.ActiveElement
            exit function
        else
            frm1.txtDept_nm.value = strDept_nm
        end if
    End if

End Function
'=======================================
'   Event Name :txtPAY_YYMM_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================

Sub txtPAY_YYMM_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtPAY_YYMM.Action = 7
        frm1.txtPAY_YYMM.focus
    End If
End Sub

Sub txtPAY_YYMM_Keypress(Key) 
    If Key = 13 Then
        Call MainQuery()
    End If
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
    	            <TD HEIGHT=20 WIDTH=10%>
    	                <FIELDSET CLASS="CLSFLD">
			            <TABLE <%=LR_SPACE_TYPE_40%>>
			    	        <TR>
			    	        	<TD CLASS="TD5" NOWRAP>급여년월</TD>
			    	    		<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtPAY_YYMM CLASS=FPDTYYYYMM title=FPDATETIME ALT="급여년월" tag="12X1" id=txtPAY_YYMM></OBJECT>');</SCRIPT></TD>
			    	    		<TD CLASS="TD5" NOWRAP>사번</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp(0)">
			    	    						<INPUT TYPE=TEXT ID="txtName" NAME="txtName" SIZE=20 tag="14X" class=protected readonly=true tabindex="-1"></TD>
			            	</TR>
			    	        <TR>
			    	    		<TD CLASS="TD5" NOWRAP></TD>
			    	    		<TD CLASS="TD6"></TD>
			    	        	<TD CLASS="TD5" NOWRAP>부서</TD>
			    	    		<TD CLASS="TD6"><INPUT NAME="txtDept_Cd" ALT="사번" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDept()">
			    	    			<INPUT TYPE=TEXT ID="txtDept_nm" NAME="txtDept_nm" SIZE=20 tag="14X" class=protected readonly=true tabindex="-1"></TD>
			            	</TR>
			            </TABLE>
			    	    </FIELDSET>
			        </TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP COLSPAN=2>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
	<TR HEIGHT="20">
		<TD WIDTH="100%" >
	  		<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
				<TD WIDTH=10>&nbsp</TD>
				<TD>
					<BUTTON NAME="btnConf" CLASS="CLSSBTN" OnClick="VBScript:Call fncCalcPay()">급여계산</BUTTON>&nbsp;&nbsp;&nbsp;
					<BUTTON NAME="btnConf" CLASS="CLSMBTN" OnClick="VBScript:Call fnBttnConf()">일괄선택/취소</BUTTON>&nbsp;
					<BUTTON NAME="btnConf" CLASS="CLSMBTN" OnClick="vbscript:getRef">사원데이타 가져오기</BUTTON></TD>
				<TD WIDTH=10>&nbsp</TD>
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
<INPUT TYPE=HIDDEN NAME="hMajorCd" tag="24"><INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

