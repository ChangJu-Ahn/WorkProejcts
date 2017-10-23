<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 교육사항조회 
*  3. Program ID           : H3011ma1
*  4. Program Name         : H3011ma1
*  5. Program Desc         : 근무이력관리/교육사항조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/01/03
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : CHCHO
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
Const BIZ_PGM_ID = "h3011mb1.asp"                                      'Biz Logic ASP 
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

Dim C_DEPT_CD
Dim C_DEPT_NM
Dim C_NAME 
Dim C_EMP_NO 
Dim C_EDU_CD 
Dim C_EDU_START_DT 
Dim C_EDU_END_DT
Dim C_EDU_OFFICE
Dim C_EDU_NAT 
Dim C_EDU_CONT 
Dim C_EDU_TYPE 
Dim C_EDU_SCORE 
Dim C_END_DT 
Dim C_EDU_FEE 
Dim C_FEE_TYPE 
Dim C_REPAY_AMT
Dim C_REPORT_TYPE
Dim C_ADD_POINT

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================
sub InitSpreadPosVariables()
	C_DEPT_CD		= 1
	C_DEPT_NM		= 2
	C_NAME			= 3
	C_EMP_NO		= 4
	C_EDU_CD 		= 5
	C_EDU_START_DT 	= 6
	C_EDU_END_DT	= 7
	C_EDU_OFFICE 	= 8
	C_EDU_NAT		= 9
	C_EDU_CONT		= 10
	C_EDU_TYPE 		= 11
	C_EDU_SCORE 	= 12
	C_END_DT 		= 13
	C_EDU_FEE		= 14
	C_FEE_TYPE 		= 15
	C_REPAY_AMT 	= 16
	C_REPORT_TYPE 	= 17
	C_ADD_POINT		= 18
end sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
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
	
	Call  ExtractDateFrom("<%=GetsvrDate%>", parent.gServerDateFormat ,  parent.gServerDateType ,strYear,strMonth,strDay)
	
	frm1.txtEdu_start_dt.Focus	
	frm1.txtEdu_start_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtEdu_start_dt.Month = strMonth 
	frm1.txtEdu_start_dt.Day = strDay
	frm1.txtEdu_end_dt.Year = strYear 		 '년월일 default value setting
	frm1.txtEdu_end_dt.Month = strMonth 
	frm1.txtEdu_end_dt.Day = strDay
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

		strTemp =  ReadCookie(CookieSplit)
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
    lgKeyStream = Trim(Frm1.txtEmp_no.Value) & parent.gColSep              
    lgKeyStream = lgKeyStream & Trim(Frm1.txtName.Value) & parent.gColSep                             'You Must append one character( parent.gColSep)
    lgKeyStream = lgKeyStream & Trim(Frm1.txtEdu_cd.Value) & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtEdu_office.Value) & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtEdu_start_dt.Text) & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtEdu_end_dt.text) & parent.gColSep
    lgKeyStream = lgKeyStream & Trim(Frm1.txtdept_cd.value) & parent.gColSep    
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
	Call initSpreadPosVariables()  
		
	With frm1.vspdData
	
         ggoSpread.Source = Frm1.vspdData
   		ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
	   .ReDraw = false

       .MaxCols = C_ADD_POINT + 1                                                      ' ☜:☜: Add 1 to Maxcols
	   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
       .ColHidden = True                                                            ' ☜:☜:
    
       .MaxRows = 0	
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData        
		Call GetSpreadColumnPos("A")       
        Call  AppendNumberPlace("6","7","2")

         ggoSpread.SSSetEdit     C_DEPT_CD,       "부서코드",    10,,,10
         ggoSpread.SSSetEdit     C_DEPT_NM,       "부서명",      13,,,13
         ggoSpread.SSSetEdit     C_NAME,		  "이름",        14,,,14         
         ggoSpread.SSSetEdit     C_EMP_NO,        "사번",        10,,,10
         ggoSpread.SSSetEdit     C_EDU_CD,        "교육코드",    10,,,20
         ggoSpread.SSSetDate     C_EDU_START_DT,  "교육시작일",  10,2,  parent.gDateFormat
         ggoSpread.SSSetDate     C_EDU_END_DT,    "교육종료일",  10,2,  parent.gDateFormat
         ggoSpread.SSSetEdit     C_EDU_OFFICE,    "교육기관",    10,,,30
         ggoSpread.SSSetEdit     C_EDU_NAT,       "교육국가",    10,,,20
         ggoSpread.SSSetEdit     C_EDU_CONT,      "교육명",      16,,,40
         ggoSpread.SSSetEdit     C_EDU_TYPE,      "구분",        06,,,10
         ggoSpread.SSSetFloat    C_EDU_SCORE,     "점수",        10,  parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
         ggoSpread.SSSetDate     C_END_DT,        "의무교육기간",10,2,  parent.gDateFormat
         ggoSpread.SSSetFloat    C_EDU_FEE,       "교육비",      10,  parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
         ggoSpread.SSSetEdit     C_FEE_TYPE,      "정산",        06,,,06
         ggoSpread.SSSetFloat    C_REPAY_AMT,     "고용보험환급",10,  parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
         ggoSpread.SSSetEdit     C_REPORT_TYPE,   "레포터",      06,,,06
         ggoSpread.SSSetFloat    C_ADD_POINT,     "고과",        10,  parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec

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
            C_DEPT_CD		= iCurColumnPos(1)
            C_DEPT_NM		= iCurColumnPos(2)
			C_NAME			= iCurColumnPos(3)
			C_EMP_NO		= iCurColumnPos(4)
			C_EDU_CD		= iCurColumnPos(5)
			C_EDU_START_DT	= iCurColumnPos(6)
			C_EDU_END_DT	= iCurColumnPos(7)
			C_EDU_OFFICE	= iCurColumnPos(8)
			C_EDU_NAT		= iCurColumnPos(9)
			C_EDU_CONT		= iCurColumnPos(10)
			C_EDU_TYPE		= iCurColumnPos(11)
			C_EDU_SCORE		= iCurColumnPos(12)
			C_END_DT		= iCurColumnPos(13)
			C_EDU_FEE		= iCurColumnPos(14)
			C_FEE_TYPE		= iCurColumnPos(15)
			C_REPAY_AMT		= iCurColumnPos(16)
			C_REPORT_TYPE	= iCurColumnPos(17)
			C_ADD_POINT		= iCurColumnPos(18)
    End Select    
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
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
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
		
    Call  ggoOper.FormatField(Document, "1", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal
    Call SetToolbar("1100000000001111")										        '버튼 툴바 제어 
    frm1.txtemp_no.Focus

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
    If  ValidDateCheck(frm1.txtEdu_start_dt, frm1.txtEdu_end_dt)=False Then
        Exit Function
    End If
     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    If  txtDept_cd_Onchange()  then
        Exit Function
    End If

    if  txtEmp_no_Onchange() then
       Exit Function
    End If		
    if  txtEdu_cd_Onchange() then
       Exit Function
    End If
    if  txtEdu_office_Onchange() then
       Exit Function
    End If    
    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")

	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
       
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

    Dim strEdu_start_dt
    Dim strEdu_end_dt
    Dim lRow

    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
     ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
	With Frm1
        For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            Select Case .vspdData.Text
               Case  ggoSpread.InsertFlag,  ggoSpread.UpdateFlag
                    .vspdData.Col = C_EDU_START_DT
                    strEdu_start_dt = .vspdData.text
                    .vspdData.Col = C_EDU_END_DT
                    strEdu_end_dt = .vspdData.text
                    IF strEdu_start_dt > strEdu_end_dt THEN
                        Call  DisplayMsgBox("800049","x","x","x")
                        .vspdData.Col = C_EDU_START_DT
                        .vspdData.Action = 0 ' go to 
                        Exit Function
                    END IF	
            End Select
        Next
    End With

	Call  DisableToolBar( parent.TBC_SAVE)
	If DbSAVE = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
    
    FncSave = True                                                              '☜: Processing is OK
    
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    Dim IntRetCD
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
    	lDelRows =  ggoSpread.DeleteRow
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
    Call parent.FncExport( parent.C_MULTI)                                         '☜: 화면 유형 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================================
Function FncFind() 
    Call parent.FncFind( parent.C_MULTI, False)                                    '☜:화면 유형, Tab 유무 
End Function

'========================================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================================
Function FncExit()
    Dim IntRetCD
	
	FncExit = False
	
     ggoSpread.Source = frm1.vspdData	
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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
 
               Case  ggoSpread.InsertFlag                                      '☜: Update
                                                   strVal = strVal & "C" & parent.gColSep
                                                   strVal = strVal & lRow & parent.gColSep
                                                   strVal = strVal & .txtEmp_no.value & parent.gColSep
                   .vspdData.Col = C_EDU_CD  	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_START_DT: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_END_DT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_OFFICE  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_NAT  	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_CONT  	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_TYPE  	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_SCORE 	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_END_DT   	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_FEE 	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_FEE_TYPE 	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_REPAY_AMT 	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_REPORT_TYPE : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_ADD_POINT   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.UpdateFlag                                      '☜: Update
                                                   strVal = strVal & "U" & parent.gColSep
                                                   strVal = strVal & lRow & parent.gColSep
                                                   strVal = strVal & .txtEmp_no.value & parent.gColSep
                   .vspdData.Col = C_EDU_CD  	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_START_DT: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_END_DT	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_OFFICE  : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_NAT  	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_CONT  	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_TYPE  	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_SCORE 	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_END_DT   	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_FEE 	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_FEE_TYPE 	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_REPAY_AMT 	 : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_REPORT_TYPE : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_ADD_POINT   : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case  ggoSpread.DeleteFlag                                      '☜: Delete

                                                   strDel = strDel & "D" & parent.gColSep
                                                   strDel = strDel & lRow & parent.gColSep
                                                   strDel = strDel & .txtEmp_no.value & parent.gColSep
                   .vspdData.Col = C_EDU_CD  	 : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                   .vspdData.Col = C_EDU_START_DT: strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	
       .txtMode.value        =  parent.UID_M0002
       .txtUpdtUserId.value  =  parent.gUsrID
       .txtInsrtUserId.value =  parent.gUsrID
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
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    
	Call  DisableToolBar( parent.TBC_DELETE)
	If DbDELETE = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
    
    FncDelete = True                                                        '⊙: Processing is OK


End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
	
    lgIntFlgMode =  parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
    Call SetToolbar("1100000000001111")										        '버튼 툴바 제어 
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
	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If																'☜: Query db data
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value
	    arrParam(1) = ""
    Else
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""			' Name Cindition
	End If
    arrParam(2) = lgUsrIntCd
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no.focus		
		Exit Function
	Else
		Call SetEmpName(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no.value = arrRet(0)
		.txtName.value = arrRet(1)
		ggoSpread.Source = Frm1.vspdData    
		ggoSpread.ClearSpreadData 		
		Set gActiveElement = document.ActiveElement
		.txtEmp_no.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'===========================================================================
' Function Name : OpenCode
' Function Desc : OpenCode Reference Popup
'===========================================================================
Function OpenCode(strCode, iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1
	    	arrParam(1) = "B_MINOR"				            	' TABLE 명칭 
	    	arrParam(2) = Trim(strCode)	                        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "MAJOR_CD=" & FilterVar("H0033", "''", "S") & ""		    		' Where Condition
	    	arrParam(5) = "교육코드"		   				    ' TextBox 명칭 
	
	    	arrField(0) = "MINOR_CD"							' Field명(0)
	    	arrField(1) = "MINOR_NM"    						' Field명(1)%>
    
	    	arrHeader(0) = "교육코드"			        		' Header명(0)%>
	    	arrHeader(1) = "교육명"	        					' Header명(1)%>

	    Case 2
	    	arrParam(1) = "B_MINOR"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(strCode)		                	' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "MAJOR_CD=" & FilterVar("H0037", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "교육기관"  						    ' TextBox 명칭 
	
	    	arrField(0) = "MINOR_CD"							' Field명(0)
	    	arrField(1) = "MINOR_NM"    						' Field명(1)
    
	    	arrHeader(0) = "교육기관코드"		        		' Header명(0)
	    	arrHeader(1) = "교육기관명"	       					' Header명(1)
	End Select

    arrParam(3) = ""	
	arrParam(0) = arrParam(5)								    ' 팝업 명칭 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Select Case iWhere
		    Case 1
		        frm1.txtEdu_cd.focus
		    Case 2
				frm1.txtEdu_office.focus
		End Select
	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
	End If	
	
End Function

'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetCode()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCode(arrRet, iWhere)

	With frm1
		Select Case iWhere
		    Case 1
		        .txtEdu_cd.value = arrRet(0)
		        .txtEdu_cd_nm.value = arrRet(1)
		        .txtEdu_cd.focus
		    Case 2
		        .txtEdu_office.value = arrRet(0)
		        .txtEdu_office_nm.value = arrRet(1)
				.txtEdu_office.focus
		End Select

		lgBlnFlgChgValue = True

	End With
End Function

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
    Dim intIndex
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col
             
   	If Frm1.vspdData.CellType =  parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) <  UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼을 클릭할 경우 발생 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub
'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000111111")       

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

    If Button = 2 And  gMouseClickStatus = "SPC" Then
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


'========================================================================================================
' Name : txtEdu_start_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtEdu_start_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")         
        frm1.txtEdu_start_dt.Action = 7 
		frm1.txtEdu_start_dt.focus
    End If
End Sub

'========================================================================================================
' Name : txtEdu_end_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtEdu_end_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")         
        frm1.txtEdu_end_dt.Action = 7 
        frm1.txtEdu_end_dt.focus
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
		frm1.txtEmp_no.value = ""
    Else
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
			ggoSpread.Source = Frm1.vspdData    
			ggoSpread.ClearSpreadData             
            call InitVariables()
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true
        Else
            frm1.txtName.value = strName
        End if 
    End if
    
End Function

'========================================================================================================
'   Event Name : txtEdu_cd_change
'   Event Desc :
'========================================================================================================
Function txtEdu_cd_Onchange()
    Dim IntRetCd


    If frm1.txtEdu_cd.value = "" Then
		frm1.txtEdu_cd_nm.value = ""
    Else
        IntRetCd =  CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0033", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtEdu_cd.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call  DisplayMsgBox("971001","X","교육코드","X")
			 frm1.txtEdu_cd_nm.value = ""
             frm1.txtEdu_cd.focus
            Set gActiveElement = document.ActiveElement
            txtEdu_cd_Onchange = true 
            
            Exit Function          
        Else
			frm1.txtEdu_cd_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End Function

'========================================================================================================
'   Event Name : txtEdu_office_change
'   Event Desc :
'========================================================================================================
Function txtEdu_office_Onchange()
    Dim IntRetCd


    If frm1.txtEdu_office.value = "" Then
		frm1.txtEdu_office_nm.value = ""
    Else
        IntRetCd =  CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD=" & FilterVar("H0037", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtEdu_office.value , "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
        If IntRetCd = false then
			Call  DisplayMsgBox("971001","X","교육기관","X")
			 frm1.txtEdu_office_nm.value = ""
             frm1.txtEdu_office.focus
            Set gActiveElement = document.ActiveElement
            txtEdu_office_Onchange = true 
            Exit Function          
        Else
			frm1.txtEdu_office_nm.value = Trim(Replace(lgF0,Chr(11),""))
        End if 
    End if  
End Function
'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtDept_cd.value			<%' 조건부에서 누른 경우 Code Condition%>
	Else 'spread
		arrParam(0) = frm1.vspdData.Text			<%' Grid에서 누른 경우 Code Condition%>
	End If
	arrParam(1) = ""								<%' Name Cindition%>
    arrParam(2) = lgUsrIntCd

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 'TextBox(Condition)
			frm1.txtDept_cd.focus
		Else 'spread
			frm1.vspdData.Col = C_Dept
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		With frm1
			If iWhere = 0 Then 'TextBox(Condition)
				.txtDept_cd.value = arrRet(0)
				.txtDept_Nm.value = arrRet(1)
				.txtDept_cd.focus
			Else 'spread
				.vspdData.Col = C_DeptNm
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_Dept
				.vspdData.Text = arrRet(0)
				.vspdData.action =0
			End If
		End With
	End If	
			
End Function
 
Function txtDept_cd_OnChange()

    Dim IntRetCd
    Dim strDept_nm, lsInternal_cd

    frm1.txtDept_nm.value = ""

    if  frm1.txtDept_cd.value <> "" then
        IntRetCd = FuncDeptName(frm1.txtDept_cd.value,"",lgUsrIntCd,strDept_nm,lsInternal_cd)
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call DisplayMsgBox("800012", "x","x","x")   ' 등록되지 않은 부서코드입니다.
            else
                Call DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
             lsInternal_cd = ""
            frm1.txtDept_cd.focus
            Set gActiveElement = document.ActiveElement
            txtDept_cd_OnChange = true
        else
            frm1.txtDept_nm.value = strDept_nm
        end if
    end if
			
End Function
'==========================================================================================
'   Event Name : txtEdu_start_dt_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtEdu_start_dt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================================================================
'   Event Name : txtEdu_end_dt_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtEdu_end_dt_KeyDown(KeyCode, Shift)
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>교육사항조회</font></td>
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
    	            <TD HEIGHT=20 WIDTH=90%>
    	                <FIELDSET CLASS="CLSFLD">
			                <TABLE <%=LR_SPACE_TYPE_40%>>
			            	    <TR>
									<TD CLASS=TD5 NOWRAP>부서</TD>
									<TD CLASS="TD6" NOWRAP>
									    <INPUT NAME="txtDept_cd" ALT="부서" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDept(0)">&nbsp;<INPUT NAME="txtDept_nm" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="14">
									</TD>
							    	<TD CLASS=TD5 NOWRAP>교육코드</TD>
							    	<TD CLASS="TD6" NOWRAP>
                                        <INPUT ID=txtEdu_cd NAME="txtEdu_cd" ALT="교육코드" TYPE="Text" SiZE=10 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode frm1.txtEdu_cd.value,1">
                                        <INPUT ID=txtEdu_cd_nm NAME="txtEdu_cd_nm" ALT="교육코드명" TYPE="Text" SiZE=20 tag=14XXXU>
							    	</TD>
			            	    </TR>
			    	            <TR>
			    	    	    	<TD CLASS="TD5" NOWRAP>사원</TD>
			    	    	    	<TD CLASS="TD6"><INPUT ID=txtEmp_no NAME="txtEmp_no" ALT="사원" TYPE="Text" MAXLENGTH=13 SiZE=13 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmpName 1"></TD>
							    	<TD CLASS=TD5 NOWRAP>교육기관</TD>
							    	<TD CLASS="TD6" NOWRAP>
    							    	<INPUT ID=txtEdu_office NAME="txtEdu_office" ALT="교육기관" TYPE="Text" SiZE=10 tag=11XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode frm1.txtEdu_office.value,2">
    							    	<INPUT ID=txtEdu_office_nm NAME="txtEdu_office_nm" ALT="교육기관명" TYPE="Text" SiZE=20 tag=14XXXU>
							    	</TD>
			    	    	    	
			            	    </TR>			            	    
			            	    <TR>
			    	            	<TD CLASS="TD5" NOWRAP>성명</TD>
			    	    	    	<TD CLASS="TD6"><INPUT ID=txtName NAME="txtName" ALT="성명" TYPE="Text" MAXLENGTH=30 SiZE=20 tag=14XXXU></TD>
			            	    
			    	    	        <TD CLASS="TD5" NOWRAP>교육기간</TD>
							    	<TD CLASS="TD6" NOWRAP>
							    		<script language =javascript src='./js/h3011ma1_fpDateTime1_txtEdu_start_dt.js'></script>
	                                    &nbsp;~&nbsp;
	                                    <script language =javascript src='./js/h3011ma1_fpDateTime2_txtEdu_end_dt.js'></script>
							    	</TD>
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
									<script language =javascript src='./js/h3011ma1_vaSpread1_vspdData.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hPayCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>



