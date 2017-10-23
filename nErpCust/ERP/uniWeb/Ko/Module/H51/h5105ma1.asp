<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h5105ma1
*  4. Program Name         : 보수월액 등록 
*  5. Program Desc         : 보수월액 조회,등록,변경,삭제 
*  6. Comproxy List        :
*  7. Modified date(First) : 2007.01
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : lee wolsan
* 10. Modifier (Last)      : 
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS    = 21	                                      '☜: Visble row
Const BIZ_PGM_ID      = "h5105mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID1     = "h5105mb2.asp"						           '☆: Biz Logic ASP Name

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

Dim C_EMPNO 
dim C_EMPNO_pop    
Dim C_EMPNONM	  
Dim C_DEPTCD	   
Dim C_DEPTCDNM  
Dim C_ACQDATE1  
Dim C_ACQDATE2  
dim C_YY
Dim C_INCOMETOT 
Dim C_WORKCNT   
Dim C_INCOMEAVR 
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
Sub InitSpreadPosVariables()	 
		
		
		C_EMPNO      = 1 
		C_EMPNO_pop= 2
		C_EMPNONM	 = 3 
		C_DEPTCD	 = 4 
		C_DEPTCDNM   = 5 
		C_ACQDATE1   = 6 
		C_ACQDATE2   = 7
		C_YY         = 8 
		C_INCOMETOT  = 9 
		C_WORKCNT    = 10
		C_INCOMEAVR  = 11

End Sub
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
	lgOldRow = 0

	gblnWinEvent        = False
	lgBlnFlawChgFlg     = False
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
	frm1.txtyyyymm.year=strYear
	frm1.txtyyyymm.Month=strMonth
	
	frm1.txtyyyy.year=strYear-1
	'frm1.txtapply_yymm_dt.Month=strMonth
	'frm1.txtapply_yymm_dt.day=strDay
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
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iDx
    
     Call  SetCombo(frm1.txtInsur_type,"1","건강보험")
     Call  SetCombo(frm1.txtInsur_type,"2","국민연금")


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
 
	    ggoSpread.Source = frm1.vspdData	
        ggoSpread.Spreadinit "V20021129",,parent.gAllowDragDropSpread    

	    .ReDraw = false
        .MaxCols = C_INCOMEAVR + 1												<%'☜: 최대 Columns의 항상 1개 증가시킴 %>
	    .Col = .MaxCols															<%'공통콘트롤 사용 Hidden Column%>
        .ColHidden = True
        
        .MaxRows = 0
        ggoSpread.ClearSpreadData
	
       ' Call  AppendNumberPlace("7","2","2")
        
	    Call  GetSpreadColumnPos("A")
         
         ggoSpread.SSSetEdit     C_EMPNO,         "사번", 10
         ggoSpread.SSSetButton   C_EMPNO_pop 
         ggoSpread.SSSetEdit     C_EMPNONM,			"성명" ,12
         ggoSpread.SSSetEdit     C_DEPTCD,		"부서" ,10
         ggoSpread.SSSetEdit     C_DEPTCDNM,   "부서" ,12
         ggoSpread.SSSetDate     C_ACQDATE1,       "자격취득일"     ,  11, 2, parent.gDateFormat
         ggoSpread.SSSetDate     C_ACQDATE2,	"자격상실일"           ,  11, 2, parent.gDateFormat
         ggoSpread.SSSetEdit     C_YY,			"정산년도" ,12,2
         ggoSpread.SSSetFloat    C_INCOMETOT,    "연보수총액"           ,15,"7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"z"
         ggoSpread.SSSetFloat    C_WORKCNT,    "근무월수"  ,15,"6", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
         ggoSpread.SSSetFloat    C_INCOMEAVR,    "보수총액"           ,15,"7", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"

         Call ggoSpread.SSSetColHidden(C_DEPTCD,  C_DEPTCD, True)

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
    
    CALL SetSpreadColor(-1,-1)
    
         ggoSpread.SpreadLock    C_EMPNONM		, -1, C_ACQDATE2
      
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
    
         ggoSpread.SSSetRequired	C_EMPNO			, pvStartRow, pvEndRow
         'ggoSpread.SSSetRequired	C_yy			, pvStartRow, pvEndRow
         ggoSpread.SSSetRequired	C_INCOMEAVR			, pvStartRow, pvEndRow
         ggoSpread.SpreadLock  C_EMPNONM, pvStartRow, C_ACQDATE2 ,pvEndRow
         ggoSpread.SpreadLock  C_yy, pvStartRow, C_yy ,pvEndRow
         
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

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)     
                       
            C_EMPNO      = iCurColumnPos(1 )   
			C_EMPNO_pop= iCurColumnPos(2 )
			C_EMPNONM	 = iCurColumnPos(3 )   
			C_DEPTCD	 = iCurColumnPos(4 )   
			C_DEPTCDNM   = iCurColumnPos(5 )   
			C_ACQDATE1   = iCurColumnPos(6 )   
			C_ACQDATE2   = iCurColumnPos(7 )   
			C_YY         = iCurColumnPos(8 )   
			C_INCOMETOT  = iCurColumnPos(9 )   
			C_WORKCNT    = iCurColumnPos(10 )  
			C_INCOMEAVR  = iCurColumnPos(11 )  
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() event
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format

'    Call  AppendNumberPlace("6","2","2")
		
    Call  ggoOper.FormatField(Document, "1", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call  ggoOper.FormatField(Document, "3", ggStrIntegeralPart,  ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    

	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
    Call  ggoOper.FormatDate(frm1.txtyyyymm,  parent.gDateFormat, 2) 
    Call  ggoOper.FormatDate(frm1.txtyyyy,  parent.gDateFormat, 3) 
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    Call InitData()
    
    Call SetDefaultVal
	Call SetToolbar("1100110100101111")												'⊙: Set ToolBar
    
    Call InitComboBox
	
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
    
    FncQuery = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing
     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    ggoSpread.ClearSpreadData

    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    
   
    
	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
       
    FncQuery = True																'☜: Processing is OK

End Function
'========================================================================================================
' Name : FncQuery1
' Desc : 자동입력 버튼으로 실행되는 쿼리 
'========================================================================================================
Function FncQuery1()
    Dim IntRetCD 
    
    FncQuery1 = False                                                            '☜: Processing is NG
    
    Err.Clear                                                                   '☜: Protect system from crashing

     ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
	ggoSpread.ClearSpreadData 									'⊙: Clear Contents  Field

    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
       Exit Function
    End If
    
    
    
	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery1 = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
       
    FncQuery1 = True																'☜: Processing is OK

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
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
	Call SetToolbar("1110111100111111")							                 '⊙: Set ToolBar
    Call SetDefaultVal
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
    
'    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                            'Check if there is retrived data
 '       Call  DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first. 
  '      Exit Function
  '  End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete? 
	If IntRetCD = vbNo Then											        
		Exit Function	
	End If
    
    Call  DisableToolBar( parent.TBC_DELETE)
	If DbDelete = False Then
		Call  RestoreToolBar()
        Exit Function
    End If
    
    FncDelete=  True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim iRow
    Dim intGrade
    Dim intStd_strt_amt, intStd_strt_amt2
    Dim intStd_end_amt, intStd_end_amt2
    Dim intStd_amt
    Dim intInsur_rate
    Dim txtAmt
	dim strWhere

    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
     ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '⊙: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then
       Exit Function
    End If
	 ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '⊙: Check contents area
       Exit Function
    End If
    
   
    
   
     if chkQty()<>true then	Exit Function
	Call  DisableToolBar( parent.TBC_SAVE)
	If DbSave = False Then
		Call  RestoreToolBar()
        Exit Function
    End If

    FncSave = True                                                              '☜: Processing is OK
    
End Function



'==========================================================================
'chkQty
'==========================================================================
Function chkQty()
	Dim i,j

	
	chkQty=false

	
		for i=1 to frm1.vspdData.maxRows
			if GetSpreadvalue(frm1.vspdData,C_INCOMEAVR,i,"X","X")="0" then
				Call displaymsgbox("140320","X", "보수총액","X")
				frm1.vspdData.Row=i
				frm1.vspdData.col=C_INCOMEAVR
				frm1.vspdData.action=0
				exit Function
			end if
	    next 
	    
	    
	
	chkQty=true
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
           
           .Row  = .ActiveRow
           .Col  = C_EMPNO
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
	Dim iRow

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

        For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
		    .vspdData.Row = iRow
		    .vspdData.Col = C_yy
		    .vspdData.value=frm1.txtyyyy.year
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
    	lDelRows =  ggoSpread.DeleteRow
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
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     

    Call InitVariables														 '⊙: Initializes local global variables

	if LayerShowHide(1) = false then
	   exit Function
	end if

    

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
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
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData     

    Call InitVariables														 '⊙: Initializes local global variables

	if LayerShowHide(1) = false then
	exit Function
	end if

    

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
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
	Call Parent.FncExport( parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_SINGLE, True)
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
    If lgBlnFlgChgValue = True Or  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
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
    dim syyyymm
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

	if LayerShowHide(1) = false then
		exit Function
	end if

	if len(frm1.txtyyyymm.Month) = 1 then 
     syyyymm = frm1.txtyyyymm.year & "0" & frm1.txtyyyymm.Month
    else
     syyyymm = frm1.txtyyyymm.year & frm1.txtyyyymm.Month
    end if
    
    if  CommonQueryRs(" top 1 1 "," HDB230T "," PAY_YYMM = '"&syyyymm&"' and INSUR_TYPE = " & FilterVar(frm1.txtInsur_type.value, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  then
    
    else
    
        if   CommonQueryRs(" isnull(max(PAY_YYMM),'') "," HDB230T "," INSUR_TYPE = " & FilterVar(frm1.txtInsur_type.value, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)  then
        	
        	if lgF0<>"" then
        	syyyymm = replace(lgF0,chr(11),"")
			frm1.txtyyyymm.year = left(syyyymm,4) 
			frm1.txtyyyymm.Month = Right(syyyymm,2) 
			end if
        end if
    end if


    
    strVal = BIZ_PGM_ID & "?txtMode="         & parent.UID_M0001                   
    strVal = strVal     & "&txtInsur_type="   & frm1.txtInsur_type.value
    strVal = strVal     & "&txtyyyymm="       & syyyymm
    strVal = strVal     & "&txtEmp_No="       & frm1.txtEmp_No.value
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey            
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows         
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
End Function

'========================================================================================================
' Name : DbQuery1
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery1()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery1 = False                                                              '☜: Processing is NG

	if LayerShowHide(1) = false then
	exit Function
	end if

    strVal = BIZ_PGM_ID1 & "?txtMode="            & parent.UID_M0001                     '☜: Query
    strVal = strVal      & "&txtKeyStream="       & lgKeyStream                   '☜: Query Key
    strVal = strVal      & "&lgStrPrevKey=" & lgStrPrevKey             '☜: Next key tag
    strVal = strVal      & "&txtMaxRows="         & Frm1.vspdData.MaxRows         '☜: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery1 = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
    
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	dim ColSep,RowSep
	dim syyyymm
    Err.Clear                                                                    '☜: Clear err status
	
	ColSep = parent.gColSep               
	RowSep = parent.gRowSep   
		
	DbSave = False														         '☜: Processing is NG
		
	if LayerShowHide(1) = false then
	exit Function
	end if
		
	With frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With
   
    if len(frm1.txtyyyymm.Month) = 1 then 
     syyyymm = frm1.txtyyyymm.year & "0" & frm1.txtyyyymm.Month
    else
     syyyymm = frm1.txtyyyymm.year & frm1.txtyyyymm.Month
    end if
    strVal = ""
    lGrpCnt = 1

	With Frm1
       
        For lRow = 1 To .vspdData.MaxRows
		.vspdData.Row = lRow
		.vspdData.Col = 0
		
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag
					strVal = strVal & "C" & parent.gColSep & lRow & parent.gColSep		'☜: C=Create
			    Case ggoSpread.UpdateFlag
					strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep		'☜: U=Update
				Case ggoSpread.DeleteFlag	
					strVal = strVal & "D" & parent.gColSep & lRow & parent.gColSep		'☜: d=delete	
			End Select			
			Select Case .vspdData.Text
			    Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag,ggoSpread.DeleteFlag  	'☜: 신규, 수정 
			    
					strVal = strVal & syyyymm  & ColSep
					strVal = strVal & Trim(GetSpreadText(.vspdData,C_EMPNO,lRow,"X","X")) & ColSep
					strVal = strVal & frm1.txtInsur_type.value  & ColSep
					strVal = strVal & frm1.txtyyyy.year  & ColSep
					strVal = strVal & Trim(GetSpreadVALUE(.vspdData,C_INCOMETOT,lRow,"X","X")) & ColSep
					strVal = strVal & Trim(GetSpreadVALUE(.vspdData,C_WORKCNT,lRow,"X","X")) & ColSep
			        strVal = strVal & Trim(GetSpreadVALUE(.vspdData,C_INCOMEAVR,lRow,"X","X")) & ColSep
			        strVal = strVal & UNIConvDateToYYYYMMdd(Trim(GetSpreadtext(.vspdData,C_ACQDATE1,lRow,"X","X")),parent.gDateFormat,"-") & ColSep
			        strVal = strVal & UNIConvDateToYYYYMMdd(Trim(GetSpreadtext(.vspdData,C_ACQDATE2,lRow,"X","X")),parent.gDateFormat,"-") & ColSep
			        strVal = strVal & parent.gChangeOrgId  & RowSep 
		     
			End Select
				
		Next

	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strVal

	End With

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
		
	if LayerShowHide(1) = false then
	exit Function
	end if
		
	strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003                                '☜: Delete
	strVal = strVal& "&txtInsuretype=" &frm1.txtInsur_type.value 
	strVal = strVal& "&txtyyyymm=" &frm1.txtyyyymm.year & frm1.txtyyyymm.month
	
	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
	call DbQueryOk1()
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
	Call SetToolbar("1100111100111111")											 '⊙: Set ToolBar
    Call  ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus    
End Function
'========================================================================================================
' Function Name : DbQueryOk1
' Function Desc : Called by MB Area when query operation is successful자동입력 버튼으로 실행되는 쿼리시 
'========================================================================================================
Function DbQueryOk1()

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode

	Call SetToolbar("1100111100111111")												'⊙: Set ToolBar
    Call  ggoOper.LockField(Document, "Q")
    ggoSpread.ClearSpreadData "T"
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
    ggoSpread.ClearSpreadData
    
    Call MainQuery
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables
	Call MainNew	
End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

   With frm1.vspdData 

		If Row > 0 And Col = C_empno_Pop Then
		    .Row = Row
		    .Col = C_empno

		    Call openEmptName(.Text)
		
		End If
    End With
    
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
    dim sEmpno
    
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    With frm1
        Select Case Col
        
             Case  C_EMPNO
    	              
    	            sEmpno = Trim(.vspdData.Text) 
    	            
    	            call setMedAcqDt(sEmpno,Row)
    	           
                    If sEmpno = "" Then
                    
                       .vspdData.Col = C_EMPNOnm
                       .vspdData.Text = ""
                       .vspdData.Col = C_deptcdnm
                       .vspdData.Text=""
 
               	    Else
               	   
	                        IntRetCd = FuncGetEmpInf2(sEmpno,lgUsrIntCd,strName,strDept_nm,_
	                                    strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                        If  IntRetCd < 0 then
	                            If  IntRetCd = -1 then
    	                    		Call DisplayMsgbox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
                                Else
                                    Call DisplayMsgbox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
                                End if
               	                    .vspdData.Text = ""
    	                            .vspdData.Col = C_EMPNOnm
                                    .vspdData.Text = ""
    	                            .vspdData.Col = C_EMPNO
    	                             .vspdData.Text=""
    	                             .vspdData.Col = C_deptcdnm
    	                             .vspdData.Text=""
                                   
                                    .vspdData.Action = 0 ' go to 
                                    Set gActiveElement = document.ActiveElement
                                    Exit Sub
                            Else
    	                            .vspdData.Col = C_EMPNOnm
                                    .vspdData.Text=strName
                                    .vspdData.Col = C_deptcdnm
                                    .vspdData.Text=strDept_nm
                           
                               
                            End If
                           
                            
                    End If
             
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


'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)

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

       If Button = 2 And  gMouseClickStatus = "SPC" Then
           gMouseClickStatus = "SPCR"
        End If
End Sub    



'========================================================================================
' Function Name : vspdData_TopLeftChange
' Function Desc : 
'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    
    if  lgIntFlgMode <> parent.OPMD_UMODE then exit sub
     
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) _
	And Not( lgStrPrevKey = "") Then
		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
		If DBQuery = False Then 
		   Call RestoreToolBar()
		   Exit Sub 
		End If 
    End if
    
End Sub

'==========================================================================================
'   Event Name : btnCb_control_OnClick()
'   Event Desc : 보수총액 가져오기 
'==========================================================================================

Sub btnCb_control_OnClick(ByVal iWhere)

	dim IntRetCD
	dim iRetArr
	dim arrVal1,arrVal2
	dim ii
	
	dim strSelect,strFrom,strWhere
	dim pay_yyyymm
	dim strYear
	dim baseDt
	
	
	 If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			'Call BtnDisabled(0)
			Exit sub
		End If
    End If
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
		Call BtnDisabled(0)
		Exit sub
    End If

   pay_yyyymm=frm1.txtyyyymm.year & frm1.txtyyyymm.Month
   strYear =frm1.txtyyyy.year
   baseDt = frm1.txtapply_yymm_dt.year &"-"& frm1.txtapply_yymm_dt.Month &"-"& frm1.txtapply_yymm_dt.day 
   Call CommonQueryRs(" COUNT(*) "," HDB230T "," PAY_YYMM = '"&pay_yyyymm&"' and INSUR_TYPE = " & FilterVar(frm1.txtInsur_type.value, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	If Trim(Replace(lgF0,Chr(11),"")) > 0 then
        intRetCD = DisplayMsgBox("800502",parent.VB_YES_NO,"X","X")	'이미생성된 자료가 있습니다.
        if intRetCD = vbNO then
			Call BtnDisabled(0)
			Exit Sub                                    
		else
			call FncDelete
		end if
    End if
    
        
    If Not chkField(Document, "2") Then									         '☜: This function check required field
		Call BtnDisabled(0)
		Exit sub
    End If
    If Not chkField(Document, "3") Then									         '☜: This function check required field
		Call BtnDisabled(0)
		Exit sub
    End If
    
	
		
	if frm1.txtyyyy.year> frm1.txtyyyymm.year  then
		call DisplayMsgBox("971012","X", frm1.txtyyyy.alt,"X")
		frm1.txtyyyy.focus()
		exit sub
	end if	
			  
	call CommonQueryRs2by2("emp_no,name", "haa010t", "", iRetArr)


		strSelect="a.emp_no,'' btn,   a.name,  a.dept_cd, a.dept_nm, a.med_acq_dt, a.med_loss_dt,'"&frm1.txtyyyy.year&"', isnull(yamt,0),work_month,isnull(yamt/work_month,0) "
		strFrom=" ( select a.emp_no,'' btn,   a.name,  a.dept_cd, d.dept_nm, b.med_acq_dt,  b.med_loss_dt,yamt" & chr(13)
		strFrom=strFrom & " ,CASE WHEN CONVERT(VARCHAR(8), ISNULL(b.med_acq_dt, b.entr_dt), 112) <'"&strYear&"'+ '0101' "& chr(13)
		strFrom=strFrom & " 			  		 THEN DATEDIFF(month, CONVERT(DATETIME, '"&strYear&"' + '0101'), CONVERT(DATETIME ,'"&strYear&"' + '1231')) + 1 "& chr(13)
		strFrom=strFrom & " WHEN CONVERT(VARCHAR(8), ISNULL(b.med_acq_dt, b.entr_dt), 112) >= '"&strYear&"'  + '0101' "& chr(13)
		strFrom=strFrom & " THEN DATEDIFF(month, ISNULL(b.med_acq_dt, b.entr_dt), CONVERT(DATETIME, '"&strYear&"' + '1231')) + 1 "& chr(13)
		strFrom=strFrom & " END  work_month "& chr(13)



		strFrom		= strFrom & " from hdf020t b join haa010t a on a.emp_no = b.emp_no "& chr(13)
		strFrom		= strFrom & " left join (	select a.emp_no,a.income_tot_amt+ a. non_tax5 - "& chr(13)
		strFrom		= strFrom & " sum(   isnull(b.a_pay_tot_amt,0) + isnull(b.a_bonus_tot_amt,0)  + isnull(b.a_after_bonus_amt,0)) yamt "& chr(13)
		strFrom		= strFrom & " from hfa050t a left join hfa040t b on a.year_yy = b.year_yy and a.emp_no = b.emp_no "& chr(13)
		strFrom		= strFrom & " where a.year_yy='"&strYear&"' "& chr(13)
		strFrom		= strFrom & " group by a.emp_no,income_tot_amt,non_tax5 ) c on b.emp_no = c.emp_no "& chr(13)
		strFrom		= strFrom & " left join b_acct_dept d on a.dept_cd  = d.dept_cd and d.org_change_id='"&parent.gChangeOrgId &"' "& chr(13)
		strFrom		= strFrom & " where isnull(b.med_acq_dt, b.entr_dt) <='"&baseDt&"' "& chr(13)
		strFrom		= strFrom & " and (isnull(convert(varchar(8), isnull(b.med_loss_dt, b.retire_dt),112),'29991231' ) >='"&baseDt&"'  )   "& chr(13)
		strFrom		= strFrom & " and CONVERT(CHAR(4), ISNULL(b.med_acq_dt, b.entr_dt), 112) <='"&strYear&"' ) a "& chr(13)
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , iRetArr) = False Then
			If lgIntFlgMode <> parent.OPMD_UMODE Then
				IntRetCD = DisplayMsgBox("124600","X","X","X") 
				exit sub 
			End If			
		Else 
		
			arrVal1 = Split(iRetArr,  Chr(12))			
            
			ggoSpread.Source     = frm1.vspdData	
			 ggoSpread.ClearSpreadData								
			For ii = 0 to  Ubound(arrVal1,1) - 1

				ggoSpread.SSShowData  arrVal1(ii) & chr(12)
				frm1.vspddata.col=0
				frm1.vspddata.text = ggoSpread.InsertFlag
				

			Next	
			
		End If
				
			
	
	exit sub
	
	
	    
	    
End Sub

'--------------------------------------------------------------------------------------------------
'	Name : openEmptName()                                                         <==== 성명/사번 팝업 
'	Description : Employee PopUp
'------------------------------------------------------------------------------------------------
Function openEmptName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = "0" Then                              'TextBox(Condition)
		arrParam(0) = frm1.txtEmp_No.value			' Code Condition
		arrParam(1) = ""'frm1.txtName.value		    ' Name Cindition
	Else                                            'spread
		arrParam(0) = frm1.vspdData.Text			' Code Condition
	End If
	arrParam(2) = lgUsrIntcd
	
	
	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		      "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = "0" Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_EmpNo
			frm1.vspdData.action =0	
		End If	
		Exit Function
	Else
		Call SetEmp(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetEmp()  ------------------------------------------------
'	Name : SetEmp()
'	Description : Employee Popup에서 Return되는 값 setting
'------------------------------------------------------------------------------------------------------
Function SetEmp(Byval arrRet, Byval iWhere)
	
	dim intRetCd
		
	With frm1
		If iWhere = "0" Then
			.txtName.value = arrRet(1)
	    	.txtEmp_no.value = arrRet(0)
			.txtEmp_no.focus
		Else
		
		
		call setMedAcqDt(arrRet(0),.vspdData.activeRow)
		
			.vspdData.Col = C_EmpNonm
			.vspdData.Text = arrRet(1)
			'.vspdData.Col = C_DEPTCD
			'.vspdData.Text = arrRet(3)
			.vspdData.Col = C_DEPTCDnm
			.vspdData.Text = arrRet(2)
			.vspdData.Col = C_EmpNo
			.vspdData.Text = arrRet(0)
			
			.vspdData.action =0	
			
			
		End If
	End With
End Function

Sub setMedAcqDt(sVal,Row)

	With frm1
	
		if sVal="" then
				
				.vspdData.Row = Row
				.vspdData.Col = C_ACQDATE1
				.vspdData.Text = ""
				.vspdData.Col = C_ACQDATE2
				.vspdData.Text = ""
				exit sub
		end if
	
		if  CommonQueryRs(" MED_ACQ_DT,MED_LOSS_DT ", "HDF020T ", "EMP_NO="& filtervar(sVal,"''","S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) then
					
				.vspdData.Row = Row
				.vspdData.Col = C_ACQDATE1
				.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
				.vspdData.Col = C_ACQDATE2
				.vspdData.Text = Trim(Replace(lgF1,Chr(11),""))
				
		    
		else
				.vspdData.Row = Row
				.vspdData.Col = C_ACQDATE1
				.vspdData.Text = ""
				.vspdData.Col = C_ACQDATE2
				.vspdData.Text = ""
				
		end if
    End With	
end Sub

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
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	                
	    If  IntRetCd < 0 then
	        If  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            Else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            End if
			frm1.txtName.value = ""
            Frm1.txtEmp_no.focus 
            Set gActiveElement = document.ActiveElement
			txtEmp_no_Onchange = true
            Exit function
        Else
			frm1.txtName.value = strName
        End if 
    End if  
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" dir=ltr>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>보수월액등록</font></td>
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
								<TD CLASS=TD5 NOWRAP>보험구분</TD>
								<TD CLASS=TD6 NOWRAP><SELECT NAME="txtInsur_type" ALT="보험구분" CLASS=cboNormal TAG="12"></OPTION></SELECT>
								<TD CLASS=TD5 NOWRAP>급여년월</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtyyyymm NAME="txtyyyymm" CLASS=FPDTYYYYMM title=FPDATETIME ALT="급여년월" tag="12X1" VIEWASTEXT> </OBJECT>');</SCRIPT></TD>                 
							</TR>
							<TR>								
								<TD CLASS=TD5 NOWRAP>사원</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_No" MAXLENGTH="13" SIZE=10 ALT ="사번" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnname" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: openEmptName(0)">
								                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE=20 ALT ="성명" tag="14XXXU"></TD>																
    			                <TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP>
								        </TD>         
							</TR> 
					
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%>></TD>
				</TR>
				
				<TR>
    	            <TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
                           <TABLE <%=LR_SPACE_TYPE_40%>>

							<TR>
								<TD CLASS=TD5 NOWRAP>정산년도</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtyyyy NAME="txtyyyy" CLASS=FPDTYYYYMM title=FPDATETIME ALT="정산년도" tag="22X1" VIEWASTEXT> </OBJECT>');</SCRIPT></TD>                 
								<TD CLASS=TD5 NOWRAP>기준일자</TD>
								<TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtapply_yymm_dt NAME="txtapply_yymm_dt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="기준일자" tag="32X1" VIEWASTEXT> </OBJECT>');</SCRIPT></TD>                 
							</TR>
							
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
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
      <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>    
	<TR HEIGHT="20">
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
				    <TD WIDTH=10>&nbsp;</TD>
				    <TD><BUTTON NAME="btnCb_control" CLASS="CLSMBTN">보수총액가져오기</BUTTON></TD>
				    <TD WIDTH=* Align=RIGHT></TD>
				    <TD WIDTH=10>&nbsp;</TD>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
