<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<% Response.Expires = -1%>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : ha103ma1
*  4. Program Name         : 퇴직기초자료 등록 
*  5. Program Desc         : 퇴직기초자료 등록,변경,삭제 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/06/05
*  8. Modified date(Last)  : 2001/06/05
*  9. Modifier (First)     : TGS(CHUN HYUNG WON)
* 10. Modifier (Last)      : 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID      = "ha103mb1.asp"						           '☆: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Dim C_EMP_NO
Dim C_EMP_NO_POP
Dim C_NAME
Dim C_ENTR_DT
Dim C_RETIRE_DT
Dim C_RETIRE_PAY_BAS_DT
Dim C_LONG_DAY
Dim C_ADJUST_DAY
Dim C_HONOR_AMT
Dim C_ETC_AMT
Dim C_PROV_YY_MM_AMT
Dim C_EXACT_YY_MM_AMT
Dim C_RETIRE_ANU_AMT
Dim C_RETIRE_INSUR
Dim C_ETC_SUB1_AMT
Dim C_ETC_SUB2_AMT
Dim C_ETC_SUB3_AMT
Dim C_ETC_SUB4_AMT
Dim C_REMARK

Const C_SHEETMAXROWS    = 22	                                      '☜: Visble row
Const C_SHEETMAXROWS_D  = 100                                          '☜: Fetch count at a time

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
Dim lsInternal_cd

'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
Sub InitSpreadPosVariables()	 
	 C_EMP_NO			= 1                                                        'Column constant for Spread Sheet
	 C_EMP_NO_POP		= 2
	 C_NAME				= 3
	 C_ENTR_DT			= 4
	 C_RETIRE_DT		= 5
	 C_RETIRE_PAY_BAS_DT = 6
	 C_LONG_DAY			= 7	 
	 C_ADJUST_DAY		= 8
	 C_HONOR_AMT		= 9
	 C_ETC_AMT			= 10
	 C_PROV_YY_MM_AMT	= 11
	 C_EXACT_YY_MM_AMT  = 12
	 C_RETIRE_ANU_AMT	= 13
	 C_RETIRE_INSUR		= 14
	 C_ETC_SUB1_AMT		= 15
	 C_ETC_SUB2_AMT		= 16
	 C_ETC_SUB3_AMT		= 17
	 C_ETC_SUB4_AMT		= 18 
	 C_REMARK			= 19
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
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
	lgOldRow = 0
	lsInternal_cd     = ""
		
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
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
Sub MakeKeyStream(pOpt)
   
    lgKeyStream       = Trim(Frm1.txtEmp_no.Value) & parent.gColSep       'You Must append one character(parent.gColSep)
    lgKeyStream       = lgKeyStream & lgUsrIntCd & parent.gColSep
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()	
    With frm1.vspdData

	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_REMARK + 1											<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        .MaxRows = 0

			Call AppendNumberPlace("7","4","0")
			Call GetSpreadColumnPos("A")        

            ggoSpread.SSSetEdit     C_EMP_NO,                       "사번", 12,,,13,2
            ggoSpread.SSSetButton   C_EMP_NO_POP    
            ggoSpread.SSSetEdit     C_NAME,                         "성명", 12,,,,2
            ggoSpread.SSSetDate     C_ENTR_DT,                      "퇴직금계산입사일", 15,2, parent.gDateFormat
            ggoSpread.SSSetDate     C_RETIRE_DT,                    "퇴직금계산퇴사일", 15,2, parent.gDateFormat
            ggoSpread.SSSetDate     C_RETIRE_PAY_BAS_DT,            "퇴직금산정기준일", 15,2, parent.gDateFormat            
			
			ggoSpread.SSSetEdit    C_LONG_DAY,						"근속일수", 15,,,30,2
            ggoSpread.SSSetFloat    C_ADJUST_DAY,                   "근속조정일수",  15,"7",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat    C_HONOR_AMT,                    "명예수당", 15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
            ggoSpread.SSSetFloat    C_ETC_AMT,                      "기타수당", 15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat    C_PROV_YY_MM_AMT,               "지급연월차", 15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat    C_EXACT_YY_MM_AMT,              "정산연월차", 15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat    C_RETIRE_ANU_AMT,               "퇴직전환금", 15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat    C_RETIRE_INSUR,					"퇴직보험금", 15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat    C_ETC_SUB1_AMT,                 "기타공제1", 15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat    C_ETC_SUB2_AMT,                 "기타공제2", 15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat    C_ETC_SUB3_AMT,                 "기타공제3", 15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetFloat    C_ETC_SUB4_AMT,                 "기타공제4", 15,parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
            ggoSpread.SSSetEdit     C_REMARK,                       "비고", 30,,,30,2
            
            Call ggoSpread.MakePairsColumn(C_EMP_NO,  C_EMP_NO_POP)

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
        ggoSpread.SpreadLock          C_EMP_NO		, -1, C_EMP_NO
        ggoSpread.SpreadLock          C_EMP_NO_POP	, -1, C_EMP_NO_POP
        ggoSpread.SpreadLock          C_NAME		, -1, C_NAME
        ggoSpread.SpreadLock          C_ENTR_DT		, -1, C_ENTR_DT
        ggoSpread.SpreadLock          C_RETIRE_DT	, -1, C_RETIRE_DT
		ggoSpread.SpreadLock          C_LONG_DAY	, -1, C_LONG_DAY
        ggoSpread.SpreadLock          C_RETIRE_PAY_BAS_DT , -1, C_RETIRE_PAY_BAS_DT

        ggoSpread.SSSetProtected      .vspdData.MaxCols   , -1, -1		
        
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    .vspdData.ReDraw = False
        ggoSpread.SSSetRequired    C_EMP_NO		, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected   C_NAME		, pvStartRow, pvEndRow
        ggoSpread.SSSetProtected   C_LONG_DAY	, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired    C_ENTR_DT	, pvStartRow, pvEndRow
        ggoSpread.SSSetRequired    C_RETIRE_DT	, pvStartRow, pvEndRow

        ggoSpread.SSSetRequired    C_RETIRE_PAY_BAS_DT	, pvStartRow, pvEndRow        
        
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

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)                
            
            C_EMP_NO			= iCurColumnPos(1)                                                        'Column constant for Spread Sheet
			C_EMP_NO_POP		= iCurColumnPos(2)
			C_NAME				= iCurColumnPos(3)
			C_ENTR_DT			= iCurColumnPos(4)
			C_RETIRE_DT			= iCurColumnPos(5)
			
            C_RETIRE_PAY_BAS_DT = iCurColumnPos(6)
			C_LONG_DAY			= iCurColumnPos(7)
			C_ADJUST_DAY		= iCurColumnPos(8)
			C_HONOR_AMT			= iCurColumnPos(9)
			C_ETC_AMT			= iCurColumnPos(10)
			C_PROV_YY_MM_AMT	= iCurColumnPos(11)
			C_EXACT_YY_MM_AMT   = iCurColumnPos(12)
			C_RETIRE_ANU_AMT	= iCurColumnPos(13)
			C_RETIRE_INSUR		= iCurColumnPos(14)
			C_ETC_SUB1_AMT		= iCurColumnPos(15)
			C_ETC_SUB2_AMT		= iCurColumnPos(16)
			C_ETC_SUB3_AMT		= iCurColumnPos(17)
			C_ETC_SUB4_AMT		= iCurColumnPos(18)            
			C_REMARK			= iCurColumnPos(19)
    End Select    
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

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
	Call SetToolbar("1100110100101111")												'⊙: Set ToolBar
    
    frm1.txtEmp_no.focus
    
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
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
 '   **********************사번체크 
    If  txtEmp_no_Onchange() = false then
        Exit Function
    End If
'   **********************
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
'    Call SetDefaultVal
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    
    Call MakeKeyStream("X")


    If DbQuery = False Then

		Exit Function
	End If															'☜: Query db data
	
    FncQuery = True																'☜: Processing is OK
   

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgbox("900015", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call ggoOper.ClearField(Document, "1")                                       '☜: Clear Condition Field
    Call ggoOper.ClearField(Document, "2")                                       '☜: Clear Contents  Field
    Call ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
	Call SetToolbar("1110111100111111")							                 '⊙: Set ToolBar
'    Call SetDefaultVal
    Call InitVariables                                                           '⊙: Initializes local global variables
    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd
    
    FncDelete = False                                                             '☜: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                            'Check if there is retrived data
        Call DisplayMsgbox("900002","X","X","X")                                  '☜: Please do Display first. 
        Exit Function
    End If
    
    IntRetCD = DisplayMsgbox("900003", parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete? 
	If IntRetCD = vbNo Then											        
		Exit Function	
	End If
    
    
    If DbDelete = False Then
		Exit Function
	End If											                  '☜: Delete db data
    
    FncDelete=  True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim strEntr_dt
    Dim strRetire_dt
    Dim strEntr_dt_HT
    Dim strRetire_dt_HT
    Dim iRow,IntRetCd2
    Dim flag
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgbox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	
	ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If

    With Frm1.vspdData
        For iRow = 1 To  .MaxRows
            .Row = iRow
            .Col = 0

           Select Case .Text
               Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			        flag = .Text
   	                .Col = C_ENTR_DT   : strEntr_dt  = .Text
   	                .Col = C_RETIRE_DT : strRetire_dt = .Text
   	                .Row = iRow : .Col = C_ENTR_DT    : strEntr_dt_HT  = .Text
   	                .Row = iRow : .Col = C_RETIRE_DT  : strRetire_dt_HT = .Text
					
					.Col = C_EMP_NO
					.Row = iRow
					IntRetCd2 = CommonQueryRs(" emp_no " ," hGA040t ","emp_no='" & .Text& "' and entr_dt='" & strEntr_dt_HT & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)					
					If ((IntRetCd2 =true) and (flag = ggoSpread.InsertFlag)) then
						Call DisplayMsgbox("800446","X","X","X")
						Exit Function
					end if
                    
'                    If UNICDate("'"& strRetire_dt &"'") <= UNICDate("'"& strEntr_dt &"'") then  : 에러 
                    If CompareDateByFormat(strEntr_dt, strRetire_dt,strEntr_dt_HT,strRetire_dt_HT, "970023",parent.gDateFormat,parent.gComDateType,False) = false then
                       
	                    Call DisplayMsgbox("800192","X","X","X")	'입사일과 퇴사일을 확인 하십시요.
	                    .Row = iRow
  	                    .Col = C_RETIRE_DT 
  	                    .Text = ""
                    
'  	                    .Col = C_ENTR_DT 
'  	                    .Text = ""
                        Set gActiveElement = document.activeElement
                        Exit Function
                    End if 
          End Select
        Next
    End With
    
    Call MakeKeyStream("X")
    If DbSave = False Then
		Exit Function
	End If			                                                    '☜: Save db data
    
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
        .Text = ""
        .Col  = C_EMP_NO
        .Text = ""
        .Col  = C_ENTR_DT
        .Text = ""
        .Col  = C_RETIRE_DT
        .Text = ""
        .Col  = C_RETIRE_PAY_BAS_DT
        .Text = ""        
    End With
	
    Set gActiveElement = document.ActiveElement   

End Function


'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo  
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
	Dim imRow

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
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgbox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgbox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables														 '⊙: Initializes local global variables

	IF LayerShowHide(1) = False Then
		Exit Function
	End If	

    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "P"	                         '☆: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 
	
    FncPrev = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call DisplayMsgbox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgbox("900017", parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
'    Call SetDefaultVal
    Call InitVariables														 '⊙: Initializes local global variables

	IF LayerShowHide(1) = False Then
		Exit Function
	End If	


    Call MakeKeyStream("X")

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data
    strVal = strVal     & "&txtPrevNext="        & "N"	                         '☆: Direction
    
	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 
	
    FncNext = True                                                               '☜: Processing is OK
	
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

	IF LayerShowHide(1) = False Then
		Exit Function
	End If	

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex             '☜: Next key tag
    strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)        '☜: Max fetched data at a time
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
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
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	IF LayerShowHide(1) = False Then
		Exit Function
	End If	

	With frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

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
                    .vspdData.Col = C_EMP_NO            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ENTR_DT   	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RETIRE_DT  	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RETIRE_PAY_BAS_DT : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    
                    .vspdData.Col = C_ADJUST_DAY  	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HONOR_AMT  	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ETC_AMT  	        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PROV_YY_MM_AMT  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EXACT_YY_MM_AMT  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RETIRE_ANU_AMT  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RETIRE_INSUR  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ETC_SUB1_AMT  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ETC_SUB2_AMT  	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ETC_SUB3_AMT  	: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_ETC_SUB4_AMT      : strVal = strVal & Trim(.vspdData.text) & parent.gColSep 
                    .vspdData.Col = C_REMARK			: strVal = strVal & Trim(.vspdData.text) & parent.gRowSep
                    
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                          strVal = strVal & "U" & parent.gColSep
                                                          strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO            : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ENTR_DT   	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RETIRE_DT  	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RETIRE_PAY_BAS_DT : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    
                    .vspdData.Col = C_ADJUST_DAY  	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_HONOR_AMT  	    : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_ETC_AMT  	        : strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_PROV_YY_MM_AMT  	: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_EXACT_YY_MM_AMT  	: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_RETIRE_ANU_AMT  	: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_RETIRE_INSUR  	: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_ETC_SUB1_AMT  	: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_ETC_SUB2_AMT  	: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_ETC_SUB3_AMT  	: strVal = strVal & Trim(.vspdData.text) & parent.gColSep
                    .vspdData.Col = C_ETC_SUB4_AMT      : strVal = strVal & Trim(.vspdData.text) & parent.gColSep 
                    .vspdData.Col = C_REMARK			: strVal = strVal & Trim(.vspdData.text) & parent.gRowSep
                     
                    lGrpCnt = lGrpCnt + 1

               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                          strDel = strDel & "D" & parent.gColSep
                                                          strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO            : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ENTR_DT   	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_RETIRE_DT  	    : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	IF LayerShowHide(1) = False Then
		Exit Function
	End If	
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete
		
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
	
	
	
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	
	Call SetToolbar("1100111100111111")												'⊙: Set ToolBar
    frm1.txtName.focus 

    Call InitData()
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()	
End Function

'======================================================================================================
'	Name : OpenEmp()
'	Description : Employee PopUp
'======================================================================================================
Function OpenEmp(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_no.value			' Code Condition
		If frm1.txtEmp_no.value="" Then
    	    arrParam(1) = ""'frm1.txtName.value		' Name Cindition
    	Else
    	    arrParam(1) = ""                		' Name Cindition
        End If    	
	Else 'spread
	    frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
	    frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	End If
	arrParam(2) = lgUsrIntCd        			' Internal_cd
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent ,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Employee Popup에서 Return되는 값 setting
'======================================================================================================
Function SetEmp(Byval arrRet, Byval iWhere)
    Dim strField
    Dim strWhere
    Dim IntRetCd
	With frm1	
		If iWhere = 0 Then 'TextBox(Condition)
			.txtEmp_no.value = arrRet(0)
			.txtName.value = arrRet(1)
		Else 'spread
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
            '----------퇴직금계산입사일/퇴사일 Default 설정 
    	    .vspdData.Col = C_EMP_NO                    
            strField = " IsNull(a.group_entr_dt,a.entr_dt), a.retire_dt, b.entr_dt, b.retire_dt, DateAdd(Day, 1, b.retire_dt) "
            strWhere = " a.emp_no  =  " & FilterVar(.vspdData.Text, "''", "S") & ""
            strWhere = strWhere & " And a.emp_no *= b.emp_no "
            strWhere = strWhere & " And b.retire_dt = (Select MAX(retire_dt) From hga040t where emp_no =  " & FilterVar(.vspdData.Text, "''", "S") & ")"
            IntRetCd = CommonQueryRs(strField," haa010t a, hga040t b ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		    If IsNull(Trim(Replace(lgF3,Chr(11),""))) Or Trim(Replace(lgF3,Chr(11),""))="" Then
    	            .vspdData.Col = C_ENTR_DT
                    .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF0,Chr(11),"")))                    
    	            .vspdData.Col = C_RETIRE_DT
                    .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF1,Chr(11),"")))                    
   	                .vspdData.Col = C_RETIRE_PAY_BAS_DT
                    .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF1,Chr(11),"")))
		    Else			
		   
		        If Replace(lgF1,Chr(11),"") <> "" and Replace(lgF3,Chr(11),"") <> "" then
		    	    If (Trim(Replace(lgF1,Chr(11),""))<>"") And (CDate(Trim(Replace(lgF1,Chr(11),""))) >= CDate(Trim(Replace(lgF3,Chr(11),"")))) then
    	                .vspdData.Col = C_ENTR_DT
                        .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF4,Chr(11),"")))                        
    	                .vspdData.Col = C_RETIRE_DT
                        .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF1,Chr(11),"")))

    	                .vspdData.Col = C_RETIRE_PAY_BAS_DT
                        .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF1,Chr(11),"")))
		    	    End If
		    	Else
    	            .vspdData.Col = C_ENTR_DT
                    .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF4,Chr(11),"")))                    
                End If		                	
		    End if
		End if    
	End With
End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col-1
	Select Case Col
	    Case C_EMP_NO_POP
                    Call OpenEmp(1)
    End Select    
End Sub

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
    Dim IntRetCD 
    Dim strField
    Dim strWhere
    Dim strName

    Dim strRetire_dt
        
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal
    
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    With frm1
        Select Case Col
             Case  C_EMP_NO
                    If Trim(.vspdData.Text) = "" Then
               	        .vspdData.Text = ""
    	                .vspdData.Col = C_NAME
                        .vspdData.Text = ""
               	    Else
	                        IntRetCd = FuncGetEmpInf2(Trim(.vspdData.Text),lgUsrIntCd,strName,strDept_nm,_
	                                    strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                        If  IntRetCd < 0 then
	                            If  IntRetCd = -1 then
    	                    		Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
                                Else
                                    Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
                                End if
               	                    .vspdData.Text = ""
    	                            .vspdData.Col = C_NAME
                                    .vspdData.Text = ""
    	                            .vspdData.Col = C_EMP_NO
                                    .vspdData.Action = 0 ' go to 
                                    Set gActiveElement = document.ActiveElement
                                    Exit Sub
                            Else
    	                            .vspdData.Col = C_NAME
                                    .vspdData.Text=strName
                                '----------퇴직금계산입사일/퇴사일 Default 설정 
    	                        .vspdData.Col = C_EMP_NO
                                strField = " IsNull(a.group_entr_dt,a.entr_dt), a.retire_dt, b.entr_dt, b.retire_dt, DateAdd(Day, 1, b.retire_dt) "
                                strWhere = " a.emp_no  =  " & FilterVar(.vspdData.Text, "''", "S") & ""
                                strWhere = strWhere & " And a.emp_no *= b.emp_no "
                                strWhere = strWhere & " And b.retire_dt = (Select MAX(retire_dt) From hga040t where emp_no =  " & FilterVar(.vspdData.Text, "''", "S") & ")"
                           
                                IntRetCd = CommonQueryRs(strField," haa010t a, hga040t b ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		                        If IsNull(Trim(Replace(lgF3,Chr(11),""))) Or Trim(Replace(lgF3,Chr(11),""))="" Then
    	                                .vspdData.Col = C_ENTR_DT
                                        .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF0,Chr(11),"")))                                        
    	                                .vspdData.Col = C_RETIRE_DT
                                        .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF1,Chr(11),"")))                                        
    									.vspdData.Col = C_RETIRE_PAY_BAS_DT
										.vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF1,Chr(11),"")))
		                        Else			
		                            if Replace(lgF1,Chr(11),"") <> "" and Replace(lgF3,Chr(11),"") <> "" then
		                        	    If (Trim(Replace(lgF1,Chr(11),""))<>"") And (CDate(Trim(Replace(lgF1,Chr(11),""))) > CDate(Trim(Replace(lgF3,Chr(11),"")))) then
    	                                    .vspdData.Col = C_ENTR_DT
                                            .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF4,Chr(11),"")))                                            
    	                                    .vspdData.Col = C_RETIRE_DT
                                            .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF1,Chr(11),"")))
    										.vspdData.Col = C_RETIRE_PAY_BAS_DT
											.vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF1,Chr(11),"")))
		                        	    End If

		                        	Else
    	                                .vspdData.Col = C_ENTR_DT
                                        .vspdData.Text=UNIDateClientFormat(Trim(Replace(lgF4,Chr(11),"")))                                        
                                    End if		                	
		                        End if
                            End If
                    End If
             Case  C_RETIRE_DT
	                        IntRetCd = FuncGetEmpInf2(Trim(.vspdData.Text),lgUsrIntCd,strName,strDept_nm,_
	                                    strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)

							strRetire_dt = Trim(.vspdData.Text)
							                                            
    	                    .vspdData.Col = C_RETIRE_PAY_BAS_DT
                            .vspdData.Text= strRetire_dt
        End Select
    End With
             
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("1101111111")
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
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub
'-----------------------------------------
Sub vspdData_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SPC" Then
          gMouseClickStatus = "SPCR"
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
    
'    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>    	           
    	If lgStrPrevKeyIndex <> "" Then                         
      		Call DisableToolBar(parent.TBC_QUERY)
      		If DBQuery = False Then
      			Call RestoreToolBar ()
      			Exit Sub
      		End If
    	End If
    End if
End Sub

Sub cboYesNo_OnChange()
    lgBlnFlgChgValue = True
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
		frm1.txtEmp_no.value = ""
        frm1.txtName.value = ""
		txtEmp_no_Onchange = true
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			frm1.txtEmp_no.value = ""
            Call ggoOper.ClearField(Document, "2")
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = false
        Else
            frm1.txtName.value = strName

            txtEmp_no_Onchange = true
        End if 
    End if
    
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00 %> ></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>퇴직기초자료등록</font></td>
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
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
			    	    		<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    		<TD CLASS="TD6" NOWRAP><INPUT NAME="txtEmp_no"  SIZE=13 MAXLENGTH=13 ALT="사번" TYPE="Text"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp(0)">
			    	    		                       <INPUT NAME="txtName"  SIZE=20  MAXLENGTH=30 ALT="성명" TYPE="Text"   tag="14XXXU"></TD>
                                <TD CLASS="TDT" NOWRAP></TD>
			    	    		<TD CLASS="TD6" NOWRAP></TD>	    	    		                       
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
					    <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
