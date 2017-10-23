<%@ LANGUAGE="VBSCRIPT" %>

<!--
======================================================================================================
*  1. Module Name          : 인사/급여관리 
*  2. Function Name        : 근태관리 
*  3. Program ID           : H4002ma1
*  4. Program Name         : 회사근무시간등록 
*  5. Program Desc         : Single Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : 
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
Const CookieSplit = 1233
Const BIZ_PGM_ID = "h4002mb1.asp"												'비지니스 로직 ASP명 
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

Dim C_APPLY_STRT_DT
Dim C_WK                      '  근무코드 
Dim C_WK_CD                   '  근무내부코드 
Dim C_WK_STRT_TYPE            '  일 
Dim C_WK_STRT_TYPE_CD         '  일내부코드 
Dim C_WK_STRT_HHMM            '  근무시작 
Dim C_WK_END_TYPE             '  일 
Dim C_WK_END_TYPE_CD          '  일내부코드 
Dim C_WK_END_HHMM             '  근무종료 
Dim C_BAS_HHMM                '  특근기준 
Dim C_RELA1_STRT_TYPE         '  일 
Dim C_RELA1_STRT_TYPE_CD      '  일내부코드 
Dim C_RELA1_STRT_HHMM         '  휴식시작 
Dim C_RELA1_END_TYPE          '  일 
Dim C_RELA1_END_TYPE_CD       '  일내부코드 
Dim C_RELA1_END_HHMM          '  휴식종료 
Dim C_RELA2_STRT_TYPE         '  일 
Dim C_RELA2_STRT_TYPE_CD      '  일내부코드 
Dim C_RELA2_STRT_HHMM         '  휴식시작 
Dim C_RELA2_END_TYPE          '  일 
Dim C_RELA2_END_TYPE_CD       '  일내부코드 
Dim C_RELA2_END_HHMM          '  휴식종료 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_APPLY_STRT_DT		  = 1	  '  적용시작일 
    C_WK                  = 2     '  근무코드 
    C_WK_CD               = 3     '  근무내부코드 
    C_WK_STRT_TYPE        = 4     '  일 
    C_WK_STRT_TYPE_CD     = 5     '  일내부코드 
    C_WK_STRT_HHMM        = 6     '  근무시작 
    C_WK_END_TYPE         = 7     '  일 
    C_WK_END_TYPE_CD      = 8     '  일내부코드 
    C_WK_END_HHMM         = 9     '  근무종료 
    C_BAS_HHMM            = 10     '  특근기준 
    C_RELA1_STRT_TYPE     = 11    '  일 
    C_RELA1_STRT_TYPE_CD  = 12    '  일내부코드 
    C_RELA1_STRT_HHMM     = 13    '  휴식시작 
    C_RELA1_END_TYPE      = 14    '  일 
    C_RELA1_END_TYPE_CD   = 15    '  일내부코드 
    C_RELA1_END_HHMM      = 16    '  휴식종료 
    C_RELA2_STRT_TYPE     = 17    '  일 
    C_RELA2_STRT_TYPE_CD  = 18    '  일내부코드 
    C_RELA2_STRT_HHMM     = 19    '  휴식시작 
    C_RELA2_END_TYPE      = 20    '  일 
    C_RELA2_END_TYPE_CD   = 21    '  일내부코드 
    C_RELA2_END_HHMM      = 22    '  휴식종료 

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
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	frm1.txtDept_cd.focus                                                           '<---- 포커스위치지정 
    frm1.txtApply_strt_dt.Text = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)
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
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    lgKeyStream = Frm1.txtWk_type.Value & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtDept_cd.Value & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtHol_type.Value & parent.gColSep
    lgKeyStream = lgKeyStream & Frm1.txtApply_strt_dt.Text & parent.gColSep
End Sub        

'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iWKArr
    Dim iWKCDArr
    Dim iDAYArr
    Dim iDAYCDArr
    Dim iNameArr1
    Dim iCodeArr1
    Dim iNameArr2
    Dim iCodeArr2    

    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0048", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iWKArr = lgF0
    iWKCDArr = lgF1
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0111", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iDAYArr = lgF0
    iDAYCDArr = lgF1
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0047", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr1 = lgF0
    iCodeArr1 = lgF1
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0049", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr2 = lgF0
    iCodeArr2 = lgF1
    Call SetCombo2(frm1.txtWk_type,iCodeArr1, iNameArr1,Chr(11))            ''''''''DB에서 불러 condition에서 
    Call SetCombo2(frm1.txtHol_type,iCodeArr2, iNameArr2,Chr(11))
       	
    ggoSpread.SetCombo Replace(iWKArr,Chr(11),vbTab), C_WK                  ''''''''DB에서 불러 gread에서 
    ggoSpread.SetCombo Replace(iWKCDArr,Chr(11),vbTab), C_WK_CD
    ggoSpread.SetCombo Replace(iDAYArr,Chr(11),vbTab), C_WK_STRT_TYPE
    ggoSpread.SetCombo Replace(iDAYCDArr,Chr(11),vbTab), C_WK_STRT_TYPE_CD
    ggoSpread.SetCombo Replace(iDAYArr,Chr(11),vbTab), C_WK_END_TYPE
    ggoSpread.SetCombo Replace(iDAYCDArr,Chr(11),vbTab), C_WK_END_TYPE_CD
    ggoSpread.SetCombo Replace(iDAYArr,Chr(11),vbTab), C_RELA1_STRT_TYPE
    ggoSpread.SetCombo Replace(iDAYCDArr,Chr(11),vbTab), C_RELA1_STRT_TYPE_CD
    ggoSpread.SetCombo Replace(iDAYArr,Chr(11),vbTab), C_RELA1_END_TYPE
    ggoSpread.SetCombo Replace(iDAYCDArr,Chr(11),vbTab), C_RELA1_END_TYPE_CD
    ggoSpread.SetCombo Replace(iDAYArr,Chr(11),vbTab), C_RELA2_STRT_TYPE
    ggoSpread.SetCombo Replace(iDAYCDArr,Chr(11),vbTab), C_RELA2_STRT_TYPE_CD
    ggoSpread.SetCombo Replace(iDAYArr,Chr(11),vbTab), C_RELA2_END_TYPE
    ggoSpread.SetCombo Replace(iDAYCDArr,Chr(11),vbTab), C_RELA2_END_TYPE_CD
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
		
			.Row = intRow
			.Col = C_WK_CD
			intIndex = .value
			.Col = C_WK
			.value = intindex
			
			.Row = intRow
			.Col = C_WK_STRT_TYPE_CD	
			intIndex = .value
			.Col = C_WK_STRT_TYPE
			.value = intindex
			
			.Row = intRow
			.Col = C_WK_END_TYPE_CD
			intIndex = .value
			.Col = C_WK_END_TYPE
			.value = intindex
			
			.Row = intRow
			.Col = C_RELA1_STRT_TYPE_CD
			intIndex = .value
			.Col = C_RELA1_STRT_TYPE
			.value = intindex
			
			.Row = intRow
			.Col = C_RELA2_STRT_TYPE_CD
			intIndex = .value
			.Col = C_RELA2_STRT_TYPE
			.value = intindex
			
			.Row = intRow
			.Col = C_RELA2_END_TYPE_CD
			intIndex = .value
			.Col = C_RELA2_END_TYPE
			.value = intindex
		Next
	End With
End Sub
'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
		.Row = Row
        Select Case Col
            Case C_WK            '부서 
                .Col = Col
                intIndex = .Value
				.Col = C_WK_CD
				.Value = intIndex
			Case C_WK_STRT_TYPE     ' 근무시작 일 type
                .Col = Col
                intIndex = .Value
				.Col = C_WK_STRT_TYPE_CD
				.Value = intIndex
			Case C_WK_END_TYPE     ' 근무종료 일 type
                .Col = Col
                intIndex = .Value
				.Col = C_WK_END_TYPE_CD
				.Value = intIndex
			Case C_RELA1_STRT_TYPE     ' 휴식시작 일 type
                .Col = Col
                intIndex = .Value
				.Col = C_RELA1_STRT_TYPE_CD
				.Value = intIndex
			Case C_RELA1_END_TYPE     ' 휴식종료 일 type
                .Col = Col
                intIndex = .Value
				.Col = C_RELA1_END_TYPE_CD
				.Value = intIndex
			Case C_RELA2_STRT_TYPE     ' 휴식시작 일 type
               .Col = Col
                intIndex = .Value
				.Col = C_RELA2_STRT_TYPE_CD
				.Value = intIndex
			Case C_RELA2_END_TYPE     ' 휴식종료 일 type
                .Col = Col
                intIndex = .Value
				.Col = C_RELA2_END_TYPE_CD
				.Value = intIndex
		End Select
	End With
   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub
'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()  
	With frm1.vspdData	
	  
	    ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
		.ReDraw = false	
	    .MaxCols = C_RELA2_END_HHMM + 1							    			'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols															'공통콘트롤 사용 Hidden Column
	    .ColHidden = True
	    .MaxRows = 0
		ggoSpread.ClearSpreadData  	
	    
		Call GetSpreadColumnPos("A") 	
		
	    ggoSpread.SSSetDate  C_APPLY_STRT_DT,       "적용시작일",12, 2, parent.gDateFormat 	
	    ggoSpread.SSSetCombo C_WK,                  "근무코드"  ,10
	    ggoSpread.SSSetCombo C_WK_CD,               "근무코드"  ,6       'hidden
	    ggoSpread.SSSetCombo C_WK_STRT_TYPE,        "일"        ,6
	    ggoSpread.SSSetCombo C_WK_STRT_TYPE_CD,     "일내부코드",6      'hidden
	    ggoSpread.SSSetTime  C_WK_STRT_HHMM,        "근무시작"  ,10,2,1,1
	    ggoSpread.SSSetCombo C_WK_END_TYPE,         "일"        ,6
	    ggoSpread.SSSetCombo C_WK_END_TYPE_CD,      "일내부코드",6      'hidden
		ggoSpread.SSSetTime  C_WK_END_HHMM,         "근무종료"  ,10,2,1,1
		ggoSpread.SSSetTime  C_BAS_HHMM,            "특근기준"  ,10,2,1,1
		ggoSpread.SSSetCombo C_RELA1_STRT_TYPE,     "일"        ,6
		ggoSpread.SSSetCombo C_RELA1_STRT_TYPE_CD,  "일내부코드",6       'hidden
		ggoSpread.SSSetTime  C_RELA1_STRT_HHMM,     "휴식시작"  ,10,2,1,1
	    ggoSpread.SSSetCombo C_RELA1_END_TYPE,      "일"        ,6
	    ggoSpread.SSSetCombo C_RELA1_END_TYPE_CD,   "일내부코드",6       'hidden
	    ggoSpread.SSSetTime  C_RELA1_END_HHMM,      "휴식종료"  ,10,2,1,1
	    ggoSpread.SSSetCombo C_RELA2_STRT_TYPE,     "일"        ,6
	    ggoSpread.SSSetCombo C_RELA2_STRT_TYPE_CD,  "일내부코드",6       'hidden
	    ggoSpread.SSSetTime  C_RELA2_STRT_HHMM,     "휴식시작"  ,10,2,1,1
	    ggoSpread.SSSetCombo C_RELA2_END_TYPE,      "일"        ,6
	    ggoSpread.SSSetCombo C_RELA2_END_TYPE_CD,   "일내부코드",6       'hidden
	    ggoSpread.SSSetTime  C_RELA2_END_HHMM,      "휴식종료"  ,10,2,1,1

        call ggoSpread.MakePairsColumn(C_WK_STRT_TYPE,C_WK_STRT_HHMM,"1")
	    call ggoSpread.MakePairsColumn(C_WK_END_TYPE,C_WK_END_HHMM,"1")
	    call ggoSpread.MakePairsColumn(C_RELA1_STRT_TYPE,C_RELA1_STRT_HHMM,"1")
	    call ggoSpread.MakePairsColumn(C_RELA1_END_TYPE,C_RELA1_END_HHMM,"1")
	    call ggoSpread.MakePairsColumn(C_RELA2_STRT_TYPE,C_RELA2_STRT_HHMM,"1")
'	    call ggoSpread.MakePairsColumn(C_RELA2_END_TYPE,C_APPLY_STRT_DT,"1")

	    Call ggoSpread.SSSetColHidden(C_WK_CD,C_WK_CD,True)	
	    Call ggoSpread.SSSetColHidden(C_WK_STRT_TYPE_CD,C_WK_STRT_TYPE_CD,True)
	    Call ggoSpread.SSSetColHidden(C_WK_END_TYPE_CD,C_WK_END_TYPE_CD,True)    
	    Call ggoSpread.SSSetColHidden(C_RELA1_STRT_TYPE_CD,C_RELA1_STRT_TYPE_CD,True)    
	    Call ggoSpread.SSSetColHidden(C_RELA1_END_TYPE_CD,C_RELA1_END_TYPE_CD,True)        
	    Call ggoSpread.SSSetColHidden(C_RELA2_STRT_TYPE_CD,C_RELA2_STRT_TYPE_CD,True)    
	    Call ggoSpread.SSSetColHidden(C_RELA2_END_TYPE_CD,C_RELA2_END_TYPE_CD,True)    
	    
		.ReDraw = true

    Call SetSpreadLock

    End With
End Sub
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos,str,iDx
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_APPLY_STRT_DT       = iCurColumnPos(1)           
            C_WK                  = iCurColumnPos(2)     '  근무코드 
            C_WK_CD               = iCurColumnPos(3)     '  근무내부코드 
            C_WK_STRT_TYPE        = iCurColumnPos(4)     '  일 
            C_WK_STRT_TYPE_CD     = iCurColumnPos(5)     '  일내부코드 
            C_WK_STRT_HHMM        = iCurColumnPos(6)     '  근무시작 
            C_WK_END_TYPE         = iCurColumnPos(7)     '  일 
            C_WK_END_TYPE_CD      = iCurColumnPos(8)     '  일내부코드 
            C_WK_END_HHMM         = iCurColumnPos(9)     '  근무종료 
            C_BAS_HHMM            = iCurColumnPos(10)     '  특근기준 
            C_RELA1_STRT_TYPE     = iCurColumnPos(11)    '  일 
            C_RELA1_STRT_TYPE_CD  = iCurColumnPos(12)    '  일내부코드 
            C_RELA1_STRT_HHMM     = iCurColumnPos(13)    '  휴식시작 
            C_RELA1_END_TYPE      = iCurColumnPos(14)    '  일 
            C_RELA1_END_TYPE_CD   = iCurColumnPos(15)    '  일내부코드 
            C_RELA1_END_HHMM      = iCurColumnPos(16)    '  휴식종료 
            C_RELA2_STRT_TYPE     = iCurColumnPos(17)    '  일 
            C_RELA2_STRT_TYPE_CD  = iCurColumnPos(18)    '  일내부코드 
            C_RELA2_STRT_HHMM     = iCurColumnPos(19)    '  휴식시작 
            C_RELA2_END_TYPE      = iCurColumnPos(20)    '  일 
            C_RELA2_END_TYPE_CD   = iCurColumnPos(21)    '  일내부코드 
            C_RELA2_END_HHMM      = iCurColumnPos(22)    '  휴식종료 

       For iDx = 1 To  frm1.vspdData.MaxCols - 1	
			str = str & iCurColumnPos(iDx) & ":"
       Next

    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_APPLY_STRT_DT, -1 , C_APPLY_STRT_DT
    ggoSpread.SpreadLock C_WK, -1 , C_WK_CD
    ggoSpread.SpreadLock C_WK_CD, -1 , C_WK_CD
    ggoSpread.SSSetRequired C_WK_STRT_TYPE, -1 , -1
    ggoSpread.SSSetRequired C_WK_STRT_TYPE_CD, -1 , -1
    ggoSpread.SSSetRequired C_WK_STRT_HHMM, -1 , -1
    ggoSpread.SSSetRequired C_WK_END_TYPE, -1 , -1
    ggoSpread.SSSetRequired C_WK_END_TYPE_CD, -1 , -1
	ggoSpread.SSSetRequired C_WK_END_HHMM, -1 , -1	
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1 	
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
    ggoSpread.SSSetRequired C_WK, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_WK_CD, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_WK_STRT_TYPE, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_WK_STRT_TYPE_CD, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_WK_STRT_HHMM, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_WK_END_TYPE, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_WK_END_TYPE_CD, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired C_WK_END_HHMM, pvStartRow, pvEndRow
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

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field    
    Call FuncGetAuth(gStrRequestMenuID , parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call InitVariables                                                              'Initializes local global variables
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitComboBox
    Call SetDefaultVal    
    Call SetToolbar("1100110100001111")										        '버튼 툴바 제어 
        
    Call CookiePage(0)
   
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


    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
   
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If txtDept_cd_Onchange() Then                                        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if
    Call InitVariables                                                        '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    Call SetSpreadLock	  '=====>v표시 

    Call DisableToolBar(parent.TBC_QUERY)
	If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End if                                              '☜: Query db data
       
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
    
    FncDelete = True                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim CFlag
   	Dim strStrtDt
   	Dim strEndDt
   	Dim strStrtDt1
   	Dim strEndDt1
   	Dim strStrtDt2
   	Dim strEndDt2
   	Dim lRow
   	Dim strStrtDtType
   	Dim strEndDtType
   	Dim strStrtDtType1
   	Dim strEndDtType1
   	Dim strStrtDtType2
   	Dim strEndDtType2

    FncSave = False                                                              '☜: Processing is NG    
    Err.Clear                                                                    '☜: Clear err status    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If    
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    If txtDept_cd_Onchange() Then                                        'enter key 로 조회시 부서코드를 check후 해당사항 없으면 query종료...
        Exit Function
    End if    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If    
	
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_WK_STRT_TYPE_CD
                    strStrtDtType = .vspdData.text
                                    
   	                .vspdData.Col = C_WK_END_TYPE_CD
                    strEndDtType =  .vspdData.text
                                        
   	                .vspdData.Col = C_RELA1_STRT_TYPE_CD
                    strStrtDtType1 = .vspdData.text
                                    
   	                .vspdData.Col = C_RELA1_END_TYPE_CD
                    strEndDtType1 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_STRT_TYPE_CD
                    strStrtDtType2 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_END_TYPE_CD
                    strEndDtType2 =  .vspdData.text
               
                    
                    If strStrtDtType = "-1" Then         '전일 
                    ElseIf  strStrtDtType = "0" Then     '당일 
                        If strEndDtType = "-1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
'''------------------------------코드세분(커서의 에러부분 정확한 이동위함)---------------
'''------------------------------ORIGINAL SOURCE CODE-----------------------------------------------
                        
'                        If strStrtDtType1 = "-1" or strStrtDtType2 = "-1" or  strEndDtType1 = "-1"  or strEndDtType2 = "-1"  Then
'	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
'	                        .vspdData.Row = lRow
 ' 	                        .vspdData.Col = C_WK_END_TYPE
  '                          .vspdData.Action = 0
  '                          Set gActiveElement = document.activeElement
  '                          Exit Function
  '                      End If 
  '''------------------------------------------------------------------------------------------------
                        If strStrtDtType1 = "-1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA1_STRT_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If                         
                        If strStrtDtType2 = "-1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA2_STRT_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If                         
                        If strEndDtType1 = "-1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA1_END_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If                         
                        If strEndDtType2 = "-1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA2_END_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If                         

                    Else                                 '익일 
                        If (strEndDtType = "-1") or (strEndDtType = "0") Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
'''-----------------------------코드세분(커서의 에러부분 정확한 이동위함)---------------  
'''-------------------------------------------------ORIGINAL SOURCE CODE------------------------------------------                                                                                                                                                                                                                     
'                        If strStrtDtType1 = "-1" or strStrtDtType2 = "-1" or strStrtDtType1 = "0" or strStrtDtType2 = "0" or strEndDtType1 = "-1" or strEndDtType2 = "-1" or strEndDtType1 = "0" or strEndDtType2 = "0" Then
'	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
'	                        .vspdData.Row = lRow
 ' 	                        .vspdData.Col = C_RELA1_END_TYPE   '''C_WK_END_TYPE
 '                           .vspdData.Action = 0
 '                           Set gActiveElement = document.activeElement
 '                           Exit Function
 '                       End If 
 '''---------------------------------------------------------------------------------------------------------------
'1 
                        If strStrtDtType1 = "-1"  Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA1_STRT_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
'2                        
                        If strStrtDtType2 = "-1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA2_STRT_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
'3                       
                        If strStrtDtType1 = "0" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA1_STRT_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
'4                        
                        If strStrtDtType2 = "0" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA2_STRT_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
'5                        
                        If strEndDtType1 = "-1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA1_END_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
'6
                        If strEndDtType2 = "-1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA2_END_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
'7                        
                        If strEndDtType1 = "0" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA1_END_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
'8                        
                        If strEndDtType2 = "0" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA2_END_TYPE   '''C_WK_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
                        

                    End if
            End Select
        Next
	End With
	
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_WK_STRT_TYPE_CD
                    strStrtDtType = .vspdData.text
                                    
   	                .vspdData.Col = C_WK_END_TYPE_CD
                    strEndDtType =  .vspdData.text
                                        
   	                .vspdData.Col = C_RELA1_STRT_TYPE_CD
                    strStrtDtType1 = .vspdData.text
                                    
   	                .vspdData.Col = C_RELA1_END_TYPE_CD
                    strEndDtType1 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_STRT_TYPE_CD
                    strStrtDtType2 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_END_TYPE_CD
                    strEndDtType2 =  .vspdData.text
                    
                    If strStrtDtType1 = "-1" Then         '전일 
                    ElseIf  strStrtDtType1 = "0" Then     '당일 
                        If strEndDtType1 = "-1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA1_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
                    Else       '익일 
                        If (strEndDtType1 = "-1") or (strEndDtType1 = "0") Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA1_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 

                    End if
            End Select
        Next
	End With
	
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_WK_STRT_TYPE_CD
                    strStrtDtType = .vspdData.text
                                    
   	                .vspdData.Col = C_WK_END_TYPE_CD
                    strEndDtType =  .vspdData.text
                                        
   	                .vspdData.Col = C_RELA1_STRT_TYPE_CD
                    strStrtDtType1 = .vspdData.text
                                    
   	                .vspdData.Col = C_RELA1_END_TYPE_CD
                    strEndDtType1 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_STRT_TYPE_CD
                    strStrtDtType2 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_END_TYPE_CD
                    strEndDtType2 =  .vspdData.text

                    If strStrtDtType2 = "-1" Then         '전일 
                    ElseIf  strStrtDtType2 = "0" Then     '당일 
                        If strEndDtType2 = "-1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA2_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
                    Else                                 '익일 
                        If (strEndDtType2 = "-1") or (strEndDtType2 = "0") Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA2_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
                    End if
            End Select
        Next
	End With
	
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_WK_STRT_TYPE_CD
                    strStrtDtType = .vspdData.text
                                    
   	                .vspdData.Col = C_WK_END_TYPE_CD
                    strEndDtType =  .vspdData.text
                                        
   	                .vspdData.Col = C_RELA1_STRT_TYPE_CD
                    strStrtDtType1 = .vspdData.text
                                    
   	                .vspdData.Col = C_RELA1_END_TYPE_CD
                    strEndDtType1 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_STRT_TYPE_CD
                    strStrtDtType2 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_END_TYPE_CD
                    strEndDtType2 =  .vspdData.text

                    If strEndDtType = "-1" Then         '전일 
'''-------------------------------ORIGINAL SOURCE CODE-------------------------------------------------------------------                    
'                         If strEndDtType1 = "1" or strEndDtType2 = "1" or strEndDtType1 = "0" or strEndDtType2 = "0" Then
'	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
'	                        .vspdData.Row = lRow
 ' 	                        .vspdData.Col = C_WK_END_TYPE
'                            .vspdData.Action = 0
'                            Set gActiveElement = document.activeElement
'                            Exit Function
'                        End If 
'''----------------------------------------------------------------------------------------------------------------------
                         If strEndDtType1 = "1" or strEndDtType1 = "0" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_REAL1_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
                         If strEndDtType2 = "1" or strEndDtType2 = "0" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_REAL2_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If                         

'''-------------------------------ORIGINAL SOURCE CODE-------------------------------------------------                        
'                   ElseIf  strEndDtType = "0" Then     '당일 
'                        If strEndDtType1 = "1"  or strEndDtType2 = "1"  Then
'	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
'	                        .vspdData.Row = lRow
'  	                        .vspdData.Col = C_WK_END_TYPE
'                            .vspdData.Action = 0
'                            Set gActiveElement = document.activeElement
'                            Exit Function
'                        End If 
'''----------------------------------------------------------------------------------------------------                        
                        If strEndDtType1 = "1" Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_REAL1_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If 
                        If strEndDtType2 = "1"  Then
	                        Call DisplayMsgBox("800445","X","X","X")	'일단위를 확인하십시요.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_REAL2_END_TYPE
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End If                                                 
      
                    Else                                 '익일 
                    End if
            End Select
        Next
	End With
	
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_WK_STRT_TYPE_CD
                    strStrtDtType = .vspdData.text
                                    
   	                .vspdData.Col = C_WK_END_TYPE_CD
                    strEndDtType =  .vspdData.text
                                        
   	                .vspdData.Col = C_RELA1_STRT_TYPE_CD
                    strStrtDtType1 = .vspdData.text
                                    
   	                .vspdData.Col = C_RELA1_END_TYPE_CD
                    strEndDtType1 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_STRT_TYPE_CD
                    strStrtDtType2 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_END_TYPE_CD
                    strEndDtType2 =  .vspdData.text

   	                .vspdData.Col = C_WK_STRT_HHMM
                    strStrtDt =  left(Replace(.vspdData.Text,":",""),6)
                    
                    If strStrtDt = "" Then
	                    Call DisplayMsgBox("970021","X","근무시작시간","X")	  '근무시작시간은 입력필수 항목 입니다.
	                    .vspdData.Row = lRow
  	                    .vspdData.Col = C_WK_STRT_HHMM
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if
                    
   	                .vspdData.Col = C_WK_END_HHMM
                    strEndDt =  left(Replace(.vspdData.Text,":",""),6)
    
                   If strEndDt = "" Then
	                    Call DisplayMsgBox("970021","X","근무종료시간","X")	      '근무종료시간은 입력필수 항목 입니다.
	                    .vspdData.Row = lRow
  	                    .vspdData.Col = C_WK_END_HHMM
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if
                    
                    If (strEndDt = "" or strEndDt = "") AND (strStrtDt<>"" and strStrtDt<>"") Then
	                    Call DisplayMsgBox("800063","X","X","X")	'시작시간 입력시 종료시간을 입력해야 합니다.
	                    .vspdData.Row = lRow
  	                    .vspdData.Col = C_WK_END_HHMM
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if
                        
                    If strStrtDtType = strEndDtType Then
                        If strStrtDt >= strEndDt then
	                        Call DisplayMsgBox("970024","X","근무시작시간","근무종료시간")	'근무시작시간은 근무종료시간 작아야합니다.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_WK_END_HHMM
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End if 
                    End if  
                    
            End Select
        Next
	End With
	
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_WK_STRT_TYPE_CD
                    strStrtDtType = .vspdData.text
                                    
   	                .vspdData.Col = C_WK_END_TYPE_CD
                    strEndDtType =  .vspdData.text
                                        
   	                .vspdData.Col = C_RELA1_STRT_TYPE_CD
                    strStrtDtType1 = .vspdData.text
                                    
   	                .vspdData.Col = C_RELA1_END_TYPE_CD
                    strEndDtType1 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_STRT_TYPE_CD
                    strStrtDtType2 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_END_TYPE_CD
                    strEndDtType2 =  .vspdData.text

   	                .vspdData.Col = C_RELA1_STRT_HHMM
                    strStrtDt1 =  left(Replace(.vspdData.Text,":",""),6)
                    
   	                .vspdData.Col = C_RELA1_END_HHMM
                    strEndDt1 =  left(Replace(.vspdData.Text,":",""),6)

                    If (strEndDt1 = "000000" or strEndDt1 = "") AND (strStrtDt1 <> "000000" and strStrtDt1 <> "")  Then
	                    Call DisplayMsgBox("800063","X","X","X")	'시작시간 입력시 종료시간을 입력해야 합니다.
	                    .vspdData.Row = lRow
  	                    .vspdData.Col = C_RELA1_END_HHMM
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if

                    If strStrtDtType1 = strEndDtType1 Then
                        If strStrtDt1 > strEndDt1 then
	                        Call DisplayMsgBox("970024","X","휴식시작시간","휴식종료시간")	'휴식시작시간은 휴식종료시간보다 작아야합니다.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA1_END_HHMM
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End if 
                    End if 
                    
            End Select
        Next
	End With
	
	
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_WK_STRT_TYPE_CD
                    strStrtDtType = .vspdData.text
                                    
   	                .vspdData.Col = C_WK_END_TYPE_CD
                    strEndDtType =  .vspdData.text
                                        
   	                .vspdData.Col = C_RELA1_STRT_TYPE_CD
                    strStrtDtType1 = .vspdData.text
                                    
   	                .vspdData.Col = C_RELA1_END_TYPE_CD
                    strEndDtType1 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_STRT_TYPE_CD
                    strStrtDtType2 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_END_TYPE_CD
                    strEndDtType2 =  .vspdData.text

   	                .vspdData.Col = C_RELA2_STRT_HHMM
                    strStrtDt2 =  left(Replace(.vspdData.Text,":",""),6)
                    
   	                .vspdData.Col = C_RELA2_END_HHMM
                    strEndDt2 =  left(Replace(.vspdData.Text,":",""),6)
   
                    If (strEndDt2 = "000000" or strEndDt2 = "") AND (strStrtDt2 <> "000000"  and strStrtDt2 <> "")  Then
	                    Call DisplayMsgBox("800063","X","X","X")	'시작시간 입력시 종료시간을 입력해야 합니다.
	                    .vspdData.Row = lRow
  	                    .vspdData.Col = C_RELA2_END_HHMM
                        .vspdData.Action = 0
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if
                    
                    If strStrtDtType2 = strEndDtType2 Then
                        If strStrtDt2 > strEndDt2 then
	                        Call DisplayMsgBox("970024","X","휴식시작시간","휴식종료시간")	'휴식시작시간은 휴식종료시간 작아야합니다.
	                        .vspdData.Row = lRow
  	                        .vspdData.Col = C_RELA2_END_HHMM
                            .vspdData.Action = 0
                            Set gActiveElement = document.activeElement
                            Exit Function
                        End if 
                    End if 
                    
            End Select
        Next
        
	End With
	

	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            Select Case .vspdData.Text
                Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

   	                .vspdData.Col = C_WK_STRT_TYPE_CD
                    strStrtDtType = .vspdData.text
                                    
   	                .vspdData.Col = C_WK_END_TYPE_CD
                    strEndDtType =  .vspdData.text
                                        
   	                .vspdData.Col = C_RELA1_STRT_TYPE_CD
                    strStrtDtType1 = .vspdData.text
                                    
   	                .vspdData.Col = C_RELA1_END_TYPE_CD
                    strEndDtType1 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_STRT_TYPE_CD
                    strStrtDtType2 =  .vspdData.text
                                    
   	                .vspdData.Col = C_RELA2_END_TYPE_CD
                    strEndDtType2 =  .vspdData.text

   	                .vspdData.Col = C_WK_STRT_HHMM
                    strStrtDt =  left(Replace(.vspdData.Text,":",""),6)
                 
   	                .vspdData.Col = C_WK_END_HHMM
                    strEndDt =  left(Replace(.vspdData.Text,":",""),6)
    
   	                .vspdData.Col = C_RELA1_STRT_HHMM
                    strStrtDt1 =  left(Replace(.vspdData.Text,":",""),6)
                    
   	                .vspdData.Col = C_RELA1_END_HHMM
                    strEndDt1 =  left(Replace(.vspdData.Text,":",""),6)
    
   	                .vspdData.Col = C_RELA2_STRT_HHMM
                    strStrtDt2 =  left(Replace(.vspdData.Text,":",""),6)
                    
   	                .vspdData.Col = C_RELA2_END_HHMM
                    strEndDt2 =  left(Replace(.vspdData.Text,":",""),6)


                    If strStrtDtType = strStrtDtType1 Then
                        If (strStrtDt1 = "000000" or strStrtDt1 = "") Then
                        Else
                            If  strStrtDt > strStrtDt1 Then
	                            Call DisplayMsgBox("800444","X","X","X")	'휴식시간은 근무 시작시간과 종료시간 안에 포함되어야 합니다.
	                            .vspdData.Row = lRow
  	                            .vspdData.Col = C_RELA1_END_HHMM
                                .vspdData.Action = 0
                                Set gActiveElement = document.activeElement
                                Exit Function
                            End if
                        End if
                    End if
                    
                    If strStrtDtType = strStrtDtType2 Then
                        If (strStrtDt2 = "000000" or strStrtDt2 = "") Then
                        Else
                            If  strStrtDt > strStrtDt2 Then
	                            Call DisplayMsgBox("800444","X","X","X")	'휴식시간은 근무 시작시간과 종료시간 안에 포함되어야 합니다.
	                            .vspdData.Row = lRow
  	                            .vspdData.Col = C_RELA2_END_HHMM
                                .vspdData.Action = 0
                                Set gActiveElement = document.activeElement
                                Exit Function
                            End if
                        End if
                    End if
                    
                    If strEndDtType = strEndDtType1 Then
                        If (strEndDt1 = "000000" or strEndDt1 = "") Then
                        Else
                            If  strEndDt < strEndDt1 Then
	                            Call DisplayMsgBox("800444","X","X","X")	'휴식시간은 근무 시작시간과 종료시간 안에 포함되어야 합니다.
	                            .vspdData.Row = lRow
  	                            .vspdData.Col = C_RELA1_END_HHMM
                                .vspdData.Action = 0
                                Set gActiveElement = document.activeElement
                                Exit Function
                            End if
                        End if
                    End if
                    
                    If strEndDtType = strEndDtType2 Then
                        If (strEndDt2 = "000000" or strEndDt2 = "") Then
                        Else
                            If  strEndDt < strEndDt2 Then
	                            Call DisplayMsgBox("800444","X","X","X")	'휴식시간은 근무 시작시간과 종료시간 안에 포함되어야 합니다.
	                            .vspdData.Row = lRow
  	                            .vspdData.Col = C_RELA2_END_HHMM
                                .vspdData.Action = 0
                                Set gActiveElement = document.activeElement
                                Exit Function
                            End if
                        End if
                    End if
                    
             End Select
        Next
	End With
	
	Call MakeKeyStream("X")
    
    If DbSave = False Then
       Exit Function
    End If
				                                                    '☜: Save db data    
    FncSave = True                                                              '☜: Processing is OK    
    
End Function
'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()  '====>표시 

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
    Call  initData()
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
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

	If LayerShowHide(1)	= False Then
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
	
    DbSave = False                                                          
    
    If LayerShowHide(1) = False Then
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
               Case ggoSpread.InsertFlag
                                                            strVal = strVal & "C" & parent.gColSep
                                                            : strVal = strVal & lRow & parent.gColSep
                                                            : strVal = strVal & .txtWk_type.value & parent.gColSep
                                                            : strVal = strVal & .txtDept_cd.value & parent.gColSep
                                                            : strVal = strVal & .txtHol_type.value & parent.gColSep
                                                            : strVal = strVal & .txtApply_strt_dt.Text & parent.gColSep
                    .vspdData.Col = C_WK_CD 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_STRT_TYPE_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_STRT_HHMM	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_END_TYPE_CD        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_END_HHMM           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_BAS_HHMM              : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA1_STRT_TYPE_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA1_STRT_HHMM       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA1_END_TYPE_CD     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA1_END_HHMM        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA2_STRT_TYPE_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA2_STRT_HHMM       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA2_END_TYPE_CD     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA2_END_HHMM        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1               
               Case ggoSpread.UpdateFlag                                    '☜: Update
                                                            strVal = strVal & "U" & parent.gColSep
                                                            : strVal = strVal & lRow & parent.gColSep
                                                            : strVal = strVal & .txtWk_type.value & parent.gColSep
                                                            : strVal = strVal & .txtDept_cd.value & parent.gColSep
                                                            : strVal = strVal & .txtHol_type.value & parent.gColSep
                                                            : strVal = strVal & .txtApply_strt_dt.Text & parent.gColSep
                    .vspdData.Col = C_WK_CD 	            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_STRT_TYPE_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_STRT_HHMM	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_END_TYPE_CD        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_WK_END_HHMM           : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_BAS_HHMM              : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA1_STRT_TYPE_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA1_STRT_HHMM       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA1_END_TYPE_CD     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA1_END_HHMM        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA2_STRT_TYPE_CD    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA2_STRT_HHMM       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA2_END_TYPE_CD     : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RELA2_END_HHMM        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_APPLY_STRT_DT           : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete    
                                                        strDel = strDel & "D" & parent.gColSep
                                                        : strDel = strDel & lRow & parent.gColSep
                                                        : strDel = strDel & .txtWk_type.value & parent.gColSep
                                                        : strDel = strDel & .txtDept_cd.value & parent.gColSep
                                                        : strDel = strDel & .txtHol_type.value & parent.gColSep
                    .vspdData.Col = C_WK_CD	            : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_APPLY_STRT_DT       : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
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
    
    DbDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    DbDelete = True                                                        '⊙: Processing is OK
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
	Call SetToolbar("1100111100111111")
	frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
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
' Name : OpenCondAreaPopup()        
' Desc : developer describe this line 
'========================================================================================================
Function OpenCondAreaPopup(Byval iWhere)
	Dim arrRet,strBasDt,strBasDtAdd
	Dim arrParam(2)

	strBasDt = frm1.txtApply_strt_dt.Text	

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = frm1.txtDept_cd.value			            '  Code Condition

	arrParam(1) = strBasDt
	arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDept_cd.focus
		Exit Function
	Else
		Call SubSetCondArea(arrRet,iWhere)
	End If	
	
End Function
	
'======================================================================================================
'	Name : SetCondArea()           
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SubSetCondArea(Byval arrRet, Byval iWhere)
	With Frm1
		Select Case iWhere
		    Case "1"
		        .txtDept_cd.value = arrRet(0)
		        .txtDept_nm.value = arrRet(1)
		        .txtInternal_cd.value = arrRet(2)
		        .txtDept_cd.focus
        End Select
	End With
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub
'========================================================================================================
'   Event Name : vspdData_Change 
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim IntRetCD
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub
'========================================================================================================
'   Event Name : cdCheck          '<==받은값의 유무,크고작음,같음에 따른 True, False 값 설정 
'   Event Desc :
'========================================================================================================
Function cdCheck(inType,strInValue,strComValue)
    Select Case UCase(inType)
        Case "EQU"
            If strInValue = strComValue Then
                cdCheck = True
            Else
                cdCheck = False
            End If
        Case "COM"
            If strInValue >= strComValue Then
                cdCheck = True
            Else
                cdCheck = False
            End If
        Case Else
            cdCheck = False
        End Select        
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================

Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
          gMouseClickStatus = "SPCR"
        End If
End Sub    
'========================================================================================================
'   Event Name : txtDept_cd_OnChange()
'   Event Desc :
'========================================================================================================
Function txtDept_cd_OnChange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    Dim strBasDtAdd
    Dim strYear,strMonth,strDay, strDate

		
	call ExtractDateFrom(frm1.txtApply_strt_dt.Text,frm1.txtApply_strt_dt.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	strDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear,strMonth,strDay)
    If frm1.txtDept_cd.value = "" Then
		frm1.txtDept_nm.value = ""
		frm1.txtinternal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtDept_cd.value , strDate , lgUsrIntCd,Dept_Nm , Internal_cd)
        
        If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if			
		    frm1.txtDept_nm.value = ""
		    frm1.txtInternal_cd.value = ""
            Set gActiveElement = document.ActiveElement 
            
            txtdept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtDept_nm.value = Dept_Nm
		    frm1.txtInternal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtApply_strt_dt_DblClick
'   Event Desc :
'========================================================================================================
Sub txtApply_strt_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")     
        frm1.txtApply_strt_dt.Action = 7
        frm1.txtApply_strt_dt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtApply_strt_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtApply_strt_dt_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
    End If
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
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!-- space Area-->

	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>회사근무시간등록</font></td>
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
			        <TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
			    </TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
			     		 <TD CLASS=TD5 NOWRAP>근무조구분</TD>       
				    	 <TD CLASS=TD6 ><SELECT Name="txtWk_type" ALT="근무조구분" STYLE="WIDTH: 100px" tag="12"></SELECT></TD>
				    	 <TD CLASS=TD5 NOWRAP>부서코드</TD>
			             <TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_cd" ALT="부서코드" TYPE="Text" SiZE=10 MAXLENGTH=10  tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: OpenCondAreaPopup('1')">
			                                  <INPUT NAME="txtDept_nm" ALT="부서코드명" TYPE="Text" SiZE=20 MAXLENGTH=40  tag="14XXXU">
						                      <INPUT NAME="txtInternal_cd" ALT="내부코드" TYPE="HIDDEN" SiZE=7 MAXLENGTH=7  tag="14XXXU"></TD>
				       </TR>
		               <TR>
			             <TD CLASS=TD5 NOWRAP>적용시작일</TD>
			             <TD CLASS=TD6 ><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 Name="txtApply_strt_dt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="적용시작일" tag="12X1" VIEWASTEXT> </OBJECT>');</SCRIPT></TD CLASS=TD6 >
		             	 <TD CLASS=TD5 NOWRAP>휴일구분</TD>
				    	 <TD CLASS=TD6 NOWRAP><SELECT Name="txtHol_type" ALT="휴일구분" STYLE="WIDTH: 100px" tag="12"></SELECT></TD>					 
				    	</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<!-- Condition Area-->
	
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<!-- space Area-->
				
				<TR>
					<TD WIDTH=100% HEIGHT=0 VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT=100% WIDTH=100% >
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				
				<!-- Content Area **** Single **** -->
				
			</TABLE>
		</TD>
	</TR>
	
	<!-- Space Area -->
	
	<!-- Button, Batch, Print, Jump Area -->
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
	</TR>
	<!-- iFrame Area -->
	
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
	<!-- Hidden Area -->

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
	<!-- Cursor Area 작업중...-->
</B
