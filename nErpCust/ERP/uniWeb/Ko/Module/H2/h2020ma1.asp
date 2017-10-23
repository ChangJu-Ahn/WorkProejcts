<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          	: Human Resources
*  2. Function Name        	: Multi Sample
*  3. Program ID           	: h2020ma1_lko311
*  4. Program Name         	: h2020ma1_lko311
*  5. Program Desc         	: 여권/비자등록 
*  6. Comproxy List        	:
*  7. Modified date(First) 	: 2005/04/18
*  8. Modified date(Last)  	: 2005/04/18
*  9. Modifier (First)     	: Lee SiNa
* 10. Modifier (Last)      	: Lee SiNa
* 11. Comment              	: 
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
Const BIZ_PGM_ID = "h2020mb1.asp"                                      '비지니스 로직 ASP명 
Const C_SHEETMAXROWS    = 21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd

Dim C_TYPE
Dim C_TYPE_NM
Dim C_EMPNO
Dim C_EmpPopup 
Dim C_EMPNM
Dim C_DEPT_NM
Dim C_ROLL_PSTN 
Dim C_ENG_NM
Dim C_RES_NO
Dim C_PASS_NO
Dim C_NAT_CD
Dim C_NATPOPUP
Dim C_NAT_NM
Dim C_ISSUE_DT
Dim C_EXPIRE_DT
Dim C_REMARK

Dim IsOpenPop          

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================

Sub initSpreadPosVariables()  
	C_TYPE		= 1
	C_TYPE_NM	= 2
	C_EMPNO		= 3
	C_EmpPopup	= 4
	C_EMPNM		= 5
	C_DEPT_NM	= 6
	C_ROLL_PSTN	= 7
	C_ENG_NM	= 8
	C_RES_NO	= 9
	C_PASS_NO	= 10
	C_NAT_CD	= 11
	C_NATPOPUP	= 12
	C_NAT_NM	= 13
	C_ISSUE_DT	= 14
	C_EXPIRE_DT = 15
	C_REMARK	= 16

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

	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	
End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

    Call CommonQueryRs("MINOR_CD, MINOR_NM ","B_MINOR","MAJOR_CD = " & FilterVar("H0005", "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.txtPay_cd, lgF0, lgF1, Chr(11))
    
    iCodeArr = "0" & chr(11) & "1" & chr(11)  
    iNameArr = "재직" & chr(11) & "퇴직" & chr(11)
    Call SetCombo2(frm1.txtRetire_cd, iCodeArr, iNameArr, Chr(11))    

    iCodeArr = "0" & chr(11) & "1" & chr(11)  
    iNameArr = "여권" & chr(11) & "비자" & chr(11)
    Call SetCombo2(frm1.txtType, iCodeArr, iNameArr, Chr(11)) 

	ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_TYPE
	ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_TYPE_NM 

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
    lgKeyStream   = Frm1.txtType.Value & parent.gColSep 
    lgKeyStream   = lgKeyStream & Frm1.txtRetire_cd.Value & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtPay_cd.Value & parent.gColSep        
    lgKeyStream   = lgKeyStream & Frm1.txtFr_internal_cd.Value & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtTo_internal_cd.Value & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtFrRoll_pstn.Value & parent.gColSep
    lgKeyStream   = lgKeyStream & Frm1.txtToRoll_pstn.Value & parent.gColSep     
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim strFlag
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows			
			.Row = intRow
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
		    Case C_TYPE_NM
		        .Col = Col
		        intIndex = .Value 
				.Col = C_TYPE
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
		ggoSpread.Spreadinit "V20050426",,parent.gAllowDragDropSpread    
	    .ReDraw = false
        .MaxCols = C_REMARK + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        .MaxRows = 0
		Call GetSpreadColumnPos("A")  

        ggoSpread.SSSetCombo    C_TYPE			, "여권/비자구분코드",   10, 0
        ggoSpread.SSSetCombo    C_TYPE_NM		, "여권/비자구분",   12, 0
        ggoSpread.SSSetEdit		C_EMPNO         , "사번", 10,,, 13, 2
        ggoSpread.SSSetButton	C_EmpPopup
        ggoSpread.SSSetEdit		C_EMPNM         , "성명", 10
        
        ggoSpread.SSSetEdit		C_DEPT_NM		, "부서", 18
        ggoSpread.SSSetEdit		C_ROLL_PSTN		, "직급", 8
        ggoSpread.SSSetEdit		C_ENG_NM		, "영문명", 16
        ggoSpread.SSSetEdit		C_RES_NO		, "주민번호", 16
        ggoSpread.SSSetEdit		C_PASS_NO		, "여권번호", 13,,,13

        ggoSpread.SSSetEdit		C_NAT_CD         , "국가코드", 8,,, 8, 2
        ggoSpread.SSSetButton	C_NATPOPUP
        ggoSpread.SSSetEdit		C_NAT_NM         , "국가", 13,,,13

        ggoSpread.SSSetDate     C_ISSUE_DT,       "발급일",   9,2, parent.gDateFormat
        ggoSpread.SSSetDate     C_EXPIRE_DT,       "만료일",   9,2, parent.gDateFormat        
        ggoSpread.SSSetEdit		C_REMARK         , "비고", 24

        Call ggoSpread.MakePairsColumn(C_EMPNO,C_EmpPopup)
        Call ggoSpread.MakePairsColumn(C_NAT_CD,C_NATPOPUP)

        call ggoSpread.SSSetColHidden(C_TYPE,C_TYPE,True)
        		        
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
            
			C_TYPE			= iCurColumnPos(1)
			C_TYPE_NM		= iCurColumnPos(2)
			C_EMPNO			= iCurColumnPos(3)
			C_EmpPopup		= iCurColumnPos(4)
			C_EMPNM			= iCurColumnPos(5)
			C_DEPT_NM		= iCurColumnPos(6)
			C_ROLL_PSTN		= iCurColumnPos(7)	
			C_ENG_NM		= iCurColumnPos(8) 
			C_RES_NO		= iCurColumnPos(9)
			C_PASS_NO		= iCurColumnPos(10)
			C_NAT_CD		= iCurColumnPos(11)
			C_NATPOPUP		= iCurColumnPos(12)
			C_NAT_NM		= iCurColumnPos(13)
			C_ISSUE_DT		= iCurColumnPos(14)
			C_EXPIRE_DT		= iCurColumnPos(15)
			C_REMARK		= iCurColumnPos(16)
		
    End Select    
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    With frm1
        .vspdData.ReDraw = False
        ggoSpread.SpreadLock    C_TYPE, -1, C_TYPE        
        ggoSpread.SpreadLock    C_TYPE_NM, -1, C_TYPE_NM
        ggoSpread.SpreadLock    C_EMPNO, -1, C_EMPNO
        ggoSpread.SpreadLock    C_EmpPopup, -1, C_EmpPopup        
        ggoSpread.SpreadLock    C_EMPNM, -1, C_EMPNM
        ggoSpread.SpreadLock    C_DEPT_NM, -1, C_DEPT_NM
        ggoSpread.SpreadLock    C_ENG_NM, -1, C_ENG_NM        
        ggoSpread.SpreadLock    C_RES_NO, -1, C_RES_NO
        ggoSpread.SpreadLock    C_ROLL_PSTN, -1, C_ROLL_PSTN        
        ggoSpread.SpreadLock    C_NAT_CD, -1, C_NAT_CD        
        ggoSpread.SpreadLock    C_NATPOPUP, -1, C_NATPOPUP
        ggoSpread.SpreadLock    C_NAT_NM, -1, C_NAT_NM

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
         ggoSpread.SSSetRequired		C_TYPE_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired		C_EMPNO,   pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_EMPNM,   pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_DEPT_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_ROLL_PSTN, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_ENG_NM, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_RES_NO, pvStartRow, pvEndRow
         ggoSpread.SSSetProtected		C_NAT_NM, pvStartRow, pvEndRow
                        
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
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet                                                            'Setup the Spread sheet
	Call InitComboBox	    
    Call InitVariables                                                              'Initializes local global variables
    Call FuncGetAuth(gStrRequestMenuID , parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
    Call SetToolbar("1100000000001111")										        '버튼 툴바 제어 
       
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
    Dim strFrDept, strToDept
    Dim Fr_dept_cd , To_dept_cd, rFrDept ,rToDept
    
     
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    If txtFr_dept_cd_Onchange() Then
        Exit Function
    End if    
    If txtTo_dept_cd_Onchange() Then
        Exit Function
    End if    

    Fr_dept_cd = frm1.txtFr_dept_cd.value
    To_dept_cd = frm1.txtTo_dept_cd.value   
    
    If fr_dept_cd = "" then    
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtFr_internal_cd.value = rFrDept				
		frm1.txtFr_dept_nm.value = ""
	End If	
	
	If to_dept_cd = "" then
        IntRetCd = FuncGetTermDept(lgUsrIntCd ,"", rFrDept ,rToDept)
		frm1.txtTo_internal_cd.value = rToDept
		frm1.txtTo_dept_nm.value = ""
	End If  
  
    Fr_dept_cd = frm1.txtFr_internal_cd.value
    To_dept_cd = frm1.txtTo_internal_cd.value
    
    If (Fr_dept_cd<> "") AND (To_dept_cd<>"") Then       
        If Fr_dept_cd > To_dept_cd then
	        Call DisplayMsgBox("800153","X","X","X")	'시작부서코드는 종료부서코드보다 작아야합니다.
            frm1.txtFr_internal_cd.value = ""
            frm1.txtTo_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.activeElement
            Exit Function
        End IF 
    END IF 
    
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    
    Call DisableToolBar(parent.TBC_QUERY)
	IF DBQUERY =  False Then
		Call RestoreToolBar()
		Exit Function
	End IF
       
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
    Dim strReturn_value, strSQL
    Dim HFlag,MFlag,Rowcnt
    Dim strVdate
    Dim strWhere
    Dim strDay_time
    
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
    HFlag = False      '올바른 값입력 
    MFlag = False    

    For Rowcnt = 1 To frm1.vspdData.MaxRows
        frm1.vspdData.Row = Rowcnt
        frm1.vspdData.Col = 0

        Select Case frm1.vspdData.Text
           
            Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag

	    End Select
    Next
  
    FncSave = True                                            
    
	Call DisableToolBar(parent.TBC_SAVE)
	If DbSave = False Then                                    '☜: Save db data     Processing is OK
		Call RestoreToolBar()
        Exit Function
    End If
    
End Function
'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
	Dim lRow
	
    FncCopy = False           
    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
   
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow, .ActiveRow
'            .Col = C_EMPNO
 '           .Text = ""
  '          .Col = C_EMPNM
   '         .Text = ""
    '        .Col = C_DEPT_NM
     '       .Text = ""
      '      .Col = C_ENG_NM
       '     .Text = ""
        '    .Col = C_RES_NO
         '   .Text = ""
          '  .Col = C_ROLL_PSTN
           ' .Text = ""
 
			Frm1.vspdData.Col = C_TYPE

			 If trim(Frm1.vspdData.text) = "0" Then
					frm1.vspdData.ReDraw = False 
					
					ggoSpread.SSSetRequired		C_PASS_NO, frm1.VspdData.ActiveRow, frm1.VspdData.ActiveRow
					ggoSpread.SSSetProtected	C_NAT_CD,  frm1.VspdData.ActiveRow, frm1.VspdData.ActiveRow					
					frm1.vspdData.ReDraw = True 				            
			 ElseIf trim(Frm1.vspdData.text) = "1" Then
					frm1.vspdData.ReDraw = False             
					ggoSpread.SSSetProtected	C_PASS_NO, frm1.VspdData.ActiveRow, frm1.VspdData.ActiveRow
					ggoSpread.SSSetRequired		C_NAT_CD,  frm1.VspdData.ActiveRow, frm1.VspdData.ActiveRow				
					frm1.vspdData.ReDraw = True 				
			 End If     

	        .ReDraw = True
 
            .Col = C_EMPNO
		    .Focus
		    .Action = 0 ' go to 
		 End If
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
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey                 '☜: Next key tag
    End With
		
    If lgIntFlgMode = parent.OPMD_UMODE Then
    Else
    End If

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
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
	
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
        
               Case ggoSpread.InsertFlag                                      '☜: Insert
                                                  strVal = strVal & "C" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
                                                  
                    .vspdData.Col = C_TYPE		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EMPNO		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_NAT_CD		    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ISSUE_DT		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EXPIRE_DT		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PASS_NO		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REMARK		        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
                    
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                  strVal = strVal & "U" & parent.gColSep
                                                  strVal = strVal & lRow & parent.gColSep
												  
                    .vspdData.Col = C_TYPE		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EMPNO		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_NAT_CD		    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_ISSUE_DT		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_EXPIRE_DT		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_PASS_NO		        : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REMARK		        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                  strDel = strDel & "D" & parent.gColSep
                                                  strDel = strDel & lRow & parent.gColSep
                                                  
                    .vspdData.Col = C_TYPE		        : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep                                                 
                    .vspdData.Col = C_EMPNO		        : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_NAT_CD		    : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
 
       .txtMode.value        = parent.UID_M0002
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
	Dim Rowcnt												     
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("110011110011111")									
	frm1.vspdData.focus	
	
    For Rowcnt = 1 To frm1.vspdData.MaxRows
        frm1.vspdData.Row = Rowcnt
        frm1.vspdData.Col = C_TYPE

        Select Case frm1.vspdData.Text
           
            Case 0
				frm1.vspdData.ReDraw = False 
				ggoSpread.SSSetRequired		C_PASS_NO, Rowcnt, Rowcnt
				frm1.vspdData.ReDraw = True 
								            
            Case 1
				frm1.vspdData.ReDraw = False             
				ggoSpread.SSSetProtected	C_PASS_NO, Rowcnt, Rowcnt
				frm1.vspdData.ReDraw = True 		
	    End Select
    Next
    	
End Function

'========================================================================================================
' Function Name : DbQueryFail
' Function Desc : Called by MB Area when query operation is unsuccessful
'========================================================================================================
Function DbQueryFail()	
 
	Call SetToolbar("110011110011111")									
	
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
'	Name : OpenEmp()
'	Description : Employee PopUp
'========================================================================================================
Function OpenEmp(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	frm1.vspdData.row =  iWhere
	If iWhere = 0 Then 'TextBox(Condition)
'		arrParam(0) = UCase(Trim(frm1.txtEmpNo.value))			<%' Code Condition%>
'		arrParam(1) = ""'frm1.txtEmpNm.value		    ' Name Cindition
	Else 'spread
		frm1.vspdData.col = C_EMPNO
		arrParam(0) = frm1.vspdData.Text			<%' Code Condition%>
		frm1.vspdData.col = C_EMPNM
		arrParam(1) = frm1.vspdData.Text			<%' Code Condition%>		
	End If

	arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 'TextBox(Condition)
'			frm1.txtEmpNo.focus
		Else 'spread
			frm1.vspdData.Col = C_EMPNO
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		With frm1
			If iWhere = 0 Then
'				.txtEmp_no.value = arrRet(0)
'				.txtName.value = arrRet(1)
'				.txtEmp_no.focus
			Else
				.vspdData.Col = C_EMPNO
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_EMPNM
				.vspdData.Text = arrRet(1)
				.vspdData.Col = C_EMPNO	
				Call CommonQueryRs(" NAME,DEPT_NM,dbo.ufn_GetCodeName("& FilterVar("H0002", "''", "S") & ",ROLL_PSTN) ROLL_PSTN_NM,ENG_NAME, RES_NO, NAT_CD ,dbo.ufn_h_GetCodeName("& FilterVar("b_country", "''", "S") & ",NAT_CD,'') NAT_NM", " HAA010T "," EMP_NO = "& FilterVar(trim(Frm1.vspdData.text), "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

				Frm1.vspdData.Col = C_DEPT_NM
				Frm1.vspdData.text =  Replace(lgF1, Chr(11), "") 

				Frm1.vspdData.Col = C_ROLL_PSTN
				Frm1.vspdData.text =  Replace(lgF2, Chr(11), "") 

				Frm1.vspdData.Col = C_ENG_NM
				Frm1.vspdData.text =  Replace(lgF3, Chr(11), "") 						
			
				Frm1.vspdData.Col = C_RES_NO
				Frm1.vspdData.text =  Replace(lgF4, Chr(11), "") 	

				Frm1.vspdData.Col = C_TYPE
 			
				If trim(Frm1.vspdData.text) = "0" Then
					Frm1.vspdData.Col = C_NAT_CD
					Frm1.vspdData.text =  Replace(lgF5, Chr(11), "") 

					Frm1.vspdData.Col = C_NAT_NM
					Frm1.vspdData.text =  Replace(lgF6, Chr(11), "") 								
				End If
						
				.vspdData.Col = C_EMPNO
				.vspdData.action =0
							
			End If
		End With
	End If	
			
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

	If iWhere = 0 Then
		arrParam(0) = frm1.txtFr_dept_cd.value			            '  Code Condition
	ElseIf iWhere = 1 Then
		arrParam(0) = frm1.txtTo_dept_cd.value			            ' Code Condition
	End If
	
	arrParam(1) = ""
	arrParam(2) = lgUsrIntcd
	
	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
		     Case "0"
               frm1.txtFr_dept_cd.focus
             Case "1"  
               frm1.txtTo_dept_cd.focus
             Case Else
        End Select	
		Exit Function
	Else
		With frm1
			Select Case iWhere
			     Case "0"
		           .txtFr_dept_cd.value = arrRet(0)
		           .txtFr_dept_nm.value = arrRet(1)
		           .txtFr_internal_cd.value = arrRet(2)
		           .txtFr_dept_cd.focus
		         Case "1"  
		           .txtTo_dept_cd.value = arrRet(0)
		           .txtTo_dept_nm.value = arrRet(1) 
		           .txtTo_internal_cd.value = arrRet(2) 
		           .txtTo_dept_cd.focus
		         Case Else
		    End Select
		End With
	End If	
			
End Function
  

'========================================================================================================
'   Event Name : txtFr_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtFr_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If frm1.txtFr_dept_cd.value = "" Then
		frm1.txtFr_dept_nm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFr_dept_cd.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFr_dept_nm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFr_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtFr_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtFr_dept_nm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtTo_dept_cd_Onchange
'   Event Desc :
'========================================================================================================
Function txtTo_dept_cd_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
    
    If frm1.txtTo_dept_cd.value = "" Then
		frm1.txtTo_dept_nm.value = ""
		frm1.txtTo_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtTo_dept_cd.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtTo_dept_nm.value = ""
		    frm1.txtTo_internal_cd.value = ""
            frm1.txtTo_dept_cd.focus
            Set gActiveElement = document.ActiveElement 
            txtTo_dept_cd_Onchange = true
            Exit Function      
        Else
			frm1.txtTo_dept_nm.value = Dept_Nm
		    frm1.txtTo_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtFrRoll_pstn_OnChange
'   Event Desc :
'========================================================================================================
Function txtFrRoll_pstn_OnChange()
    txtFrRoll_pstn_OnChange = true
    
    If  frm1.txtFrRoll_pstn.value = "" Then
        frm1.txtFrRoll_pstn_nm.value = ""
        frm1.txtFrRoll_pstn.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0002", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtFrRoll_pstn.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtFrRoll_pstn_nm.value = ""
            Call  DisplayMsgBox("970000", "x",frm1.txtFrRoll_pstn.alt,"x")
	        frm1.txtFrRoll_pstn.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtFrRoll_pstn_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

'========================================================================================================
'   Event Name : txtToRoll_pstn_OnChange
'   Event Desc :
'========================================================================================================
Function txtToRoll_pstn_OnChange()
    txtToRoll_pstn_OnChange = true
    
    If  frm1.txtToRoll_pstn.value = "" Then
        frm1.txtToRoll_pstn_nm.value = ""
        frm1.txtToRoll_pstn.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0002", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtToRoll_pstn.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtToRoll_pstn_nm.value = ""
            Call  DisplayMsgBox("970000", "x",frm1.txtToRoll_pstn.alt,"x")
	        frm1.txtToRoll_pstn.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtToRoll_pstn_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function
'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	ggoSpread.Source = frm1.vspdData
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
    
    If Row > 0 Then
		Select Case Col
			Case C_EmpPopup
				Call OpenEmp(Row)
			Case C_NATPOPUP
				Call OpenCode(3,Row)        
		End Select    
	End If
            
End Sub
'========================================================================================================
'   Event Name : vspdData_Change 
'   Event Desc :
'========================================================================================================
Function vspdData_Change(ByVal Col , ByVal Row )

   Dim iDx
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strValidDt

    Dim strAllColVal
	Dim arrRet
	Dim strFg
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_TYPE_NM
			Frm1.vspdData.Col = C_TYPE

            If trim(Frm1.vspdData.text) = "0" Then
				frm1.vspdData.ReDraw = False 
				
				Frm1.vspdData.Col = C_EMPNM
				Frm1.vspdData.text =  Replace(lgF0, Chr(11), "") 
				
				ggoSpread.SSSetRequired		C_PASS_NO, Row, Row
				ggoSpread.SSSetProtected	C_NAT_CD,  Row, Row					
				frm1.vspdData.ReDraw = True 				            
            ElseIf trim(Frm1.vspdData.text) = "1" Then
				
				frm1.vspdData.ReDraw = False 
				Frm1.vspdData.Col = C_PASS_NO
				Frm1.vspdData.text = ""            
				ggoSpread.SSSetProtected	C_PASS_NO, Row, Row
				ggoSpread.SSSetRequired		C_NAT_CD,  Row, Row				
				frm1.vspdData.ReDraw = True 				
            End If         
     
         Case  C_EMPNO

			Call CommonQueryRs(" NAME,DEPT_NM,dbo.ufn_GetCodeName("& FilterVar("H0002", "''", "S") & ",ROLL_PSTN) ROLL_PSTN_NM,ENG_NAME, RES_NO, NAT_CD ,dbo.ufn_h_GetCodeName("& FilterVar("b_country", "''", "S") & ",NAT_CD,'') NAT_NM", " HAA010T "," EMP_NO = "& FilterVar(trim(Frm1.vspdData.text), "''", "S"), lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			Frm1.vspdData.Col = C_EMPNM
			Frm1.vspdData.text =  Replace(lgF0, Chr(11), "") 

			Frm1.vspdData.Col = C_DEPT_NM
			Frm1.vspdData.text =  Replace(lgF1, Chr(11), "") 
			
			Frm1.vspdData.Col = C_ROLL_PSTN
			Frm1.vspdData.text =  Replace(lgF2, Chr(11), "") 

			Frm1.vspdData.Col = C_ENG_NM
			Frm1.vspdData.text =  Replace(lgF3, Chr(11), "") 
			
			Frm1.vspdData.Col = C_RES_NO
			Frm1.vspdData.text =  Replace(lgF4, Chr(11), "") 

			Frm1.vspdData.Col = C_TYPE
 			
			If trim(Frm1.vspdData.text) = "0" Then
				Frm1.vspdData.Col = C_NAT_CD
				Frm1.vspdData.text =  Replace(lgF5, Chr(11), "") 

				Frm1.vspdData.Col = C_NAT_NM
				Frm1.vspdData.text =  Replace(lgF6, Chr(11), "") 								
			End If
			
		 Case  C_NAT_CD
			If  frm1.vspdData.text <> "" Then
			    if  CommonQueryRs(" country_nm "," B_COUNTRY "," country_cd =  " & FilterVar(frm1.vspdData.text , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
			        Frm1.vspdData.Col = C_NAT_NM
			        frm1.vspdData.text = ""
			        Call  DisplayMsgBox("970000", "x","국가코드","x")
			        Set gActiveElement = document.ActiveElement
			        exit function
			    Else
			        Frm1.vspdData.Col = C_NAT_NM
			        frm1.vspdData.text = Replace(lgF0, Chr(11), "")	    
			    End If
			else 
			        Frm1.vspdData.Col = C_NAT_NM
			        frm1.vspdData.text = ""
			End If
	 
    End Select  
   	If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Function

'========================================================================================================
'   Event Name : txtFrDept_Onchange
'   Event Desc :
'========================================================================================================
Function txtFrDept_Onchange()
    Dim IntRetCd,Dept_Nm,Internal_cd
   
    If frm1.txtFrDept.value = "" Then
		frm1.txtFrDeptNm.value = ""
		frm1.txtFr_internal_cd.value = ""
    Else
        IntRetCd = FuncDeptName(frm1.txtFrDept.value , "" , lgUsrIntCd,Dept_Nm , Internal_cd)
    
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800098","X","X","X")	'부서정보코드에 등록되지 않은 코드입니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
		    frm1.txtFrDeptNm.value = ""
		    frm1.txtFr_internal_cd.value = ""
            frm1.txtFrDept.focus
            Set gActiveElement = document.ActiveElement 
            txtFrDept_Onchange = true
            Exit Function      
        Else
			frm1.txtFrDeptNm.value = Dept_Nm
		    frm1.txtFr_internal_cd.value = Internal_cd
        End if 
    End if  
    
End Function

'========================================================================================================
'   Event Name : txtEmpNo_OnChange
'   Event Desc :
'========================================================================================================
Function txtEmpNo_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd

    If  frm1.txtEmpNo.value = "" Then
		frm1.txtEmpNm.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(frm1.txtEmpNo.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtEmpNm.value = ""
            Frm1.txtEmpNo.focus 
            Set gActiveElement = document.ActiveElement
			txtEmpNo_Onchange = true
        Else
			frm1.txtEmpNm.value = strName
        End if 
    End if  
End Function

'===========================================================================
' Function Name : OpenCode
' Function Desc : OpenCode Reference Popup
'===========================================================================
Function OpenCode(Byval caseCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If trim(iWhere) <>"" Then
		Frm1.vspdData.row = iWhere
	End If
	
	Select Case caseCode
	    Case 1  ' 시작직급 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtFrRoll_pstn.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0002", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직급"    						    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직급코드"			        		' Header명(0)
	    	arrHeader(1) = "직급명"	        					' Header명(1)	
	    	
	    Case 2  ' 종료직급 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtToRoll_pstn.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0002", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직급"    						    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직급코드"			        		' Header명(0)
	    	arrHeader(1) = "직급명"	        					' Header명(1)
        Case 3  ' 국가코드 
            arrParam(1) = "B_COUNTRY"		    			    ' TABLE 명칭 
			Frm1.vspdData.Col = C_NAT_CD   
            arrParam(2) = Trim(frm1.vspdData.text)            ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = ""                				    ' Where Condition
            arrParam(5) = "국가코드"	    				    ' TextBox 명칭 
	
            arrField(0) = "country_cd"	    				    ' Field명(0)
            arrField(1) = "country_nm"                          ' Field명(1)
    
            arrHeader(0) = "국가코드"                           ' Header명(0)
            arrHeader(1) = "국가명"                             ' Header명(1)
	End Select

    'arrParam(3) = ""	
	arrParam(0) = arrParam(5)								    ' 팝업 명칭 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	If arrRet(0) = "" Then
		Select Case caseCode
			Case 1
				frm1.txtFrRoll_pstn.focus		
			Case 2
				frm1.txtToRoll_pstn.focus
			Case 3
				Frm1.vspdData.Col = C_NAT_CD  
				Frm1.vspdData.action = 0			
		End Select
		Exit Function
	Else
		Select Case caseCode
			Case 1
				frm1.txtFrRoll_pstn.value = arrRet(0)
				frm1.txtFrRoll_pstn_nm.value = arrRet(1)  
				frm1.txtFrRoll_pstn.focus		
			Case 2
				frm1.txtToRoll_pstn.value = arrRet(0)
				frm1.txtToRoll_pstn_nm.value = arrRet(1)  
				frm1.txtToRoll_pstn.focus
			Case 3
				Frm1.vspdData.Col = C_NAT_CD  
				Frm1.vspdData.text =  arrRet(0)		
				Frm1.vspdData.Col = C_NAT_NM  
				Frm1.vspdData.text =  arrRet(1)					
		End Select
	End If	
	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>여권/비자등록</font></td>
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
								<TD CLASS=TD5 NOWRAP>여권/비자구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtType" ALT="여권/비자구분" CLASS ="cbonormal" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
								<TD CLASS=TD5 NOWRAP>재직/퇴직구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtRetire_cd" ALT="재직/퇴직구분" CLASS ="cbonormal" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
			            	</TR>						
			            	<TR>
								<TD CLASS=TD5 NOWRAP>급여구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT Name="txtPay_cd" ALT="급여구분" CLASS ="cbonormal" tag="11"><OPTION Value=""></OPTION></SELECT></TD>
							    <TD CLASS=TD5 NOWRAP>부서</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFr_dept_cd" ALT="부서코드" TYPE="Text" SiZE="10" MAXLENGTH="10"  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(0)">
			                                         <INPUT NAME="txtFr_dept_nm" ALT="부서코드명" TYPE="Text" SiZE="20" MAXLENGTH="40"  tag="14XXXU">&nbsp;~&nbsp;
		                                             <INPUT NAME="txtFr_internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7" tag="14XXXU">
			            	</TR>
			            	<TR>	
								<TD CLASS=TD5 NOWRAP>직급</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtFrRoll_pstn" ALT="시작직급" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFr_Roll_pstn" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode 1,''">&nbsp;<INPUT NAME="txtFrRoll_pstn_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24">&nbsp;~&nbsp;
													 <INPUT NAME="txtToRoll_pstn" ALT="종료직급" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTo_Roll_pstn" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCode 2,''">&nbsp;<INPUT NAME="txtToRoll_pstn_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
			            			            	
		                        <TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtto_dept_cd" MAXLENGTH="10" SIZE="10"  ALT ="Order ID" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept(1)">
							                         <INPUT NAME="txtto_dept_nm" MAXLENGTH="40" SIZE="20"  ALT ="Order ID" tag="14XXXU">
    			                                     <INPUT NAME="txtTo_internal_cd" ALT="내부부서코드" TYPE="HIDDEN" SiZE="7" MAXLENGTH="7"  tag="14XXXU"></TD>
			            	</TR>			            	

						</TABLE>
						</FIELDSET>
					</TD>
				</TR>	
				<TR>
				    <TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT=100% WIDTH=100% >
									<script language =javascript src='./js/h2020ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

