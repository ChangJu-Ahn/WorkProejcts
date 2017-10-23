<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 
*  3. Program ID           : h4011ma1
*  4. Program Name         : h4011ma1
*  5. Program Desc         : 근태관리/일일인원현황조회 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/
*  8. Modified date(Last)  : 2003/06/11
*  9. Modifier (First)     : mok young bin
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
Const BIZ_PGM_ID = "h4011mb1.asp"                                      'Biz Logic ASP 

Const C_SHEETMAXROWS =   21	                                      '한 화면에 보여지는 최대갯수*1.5%>

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop          
Dim lsInternal_cd

Dim C_DEPT_CD 
Dim C_DEPT_NM 
Dim C_DEPT_CNT
Dim C_RETIRE_CNT
Dim C_ENTR_CNT  
Dim C_DILIG_CD  
Dim C_DILIG_NM  
Dim C_CNT       

Dim C_DEPT_CD2  
Dim C_DEPT_NM2  
Dim C_DEPT_CNT2 
Dim C_RETIRE_CNT2 
Dim C_ENTR_CNT2   
Dim C_DILIG_CD2   
Dim C_DILIG_NM2   
Dim C_CNT2        

'========================================================================================================
' Name : initSpreadPosVariables(spd)	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables(spd) 
	if spd="A" or spd="ALL" then
		C_DEPT_CD         =    1
		C_DEPT_NM         =    2
		C_DEPT_CNT        =    3
		C_RETIRE_CNT      =    4
		C_ENTR_CNT        =    5
		C_DILIG_CD        =    6
		C_DILIG_NM        =    7
		C_CNT             =    8
	end if
	if spd="B" or spd="ALL" then
		C_DEPT_CD2         =    1
		C_DEPT_NM2         =    2
		C_DEPT_CNT2        =    3
		C_RETIRE_CNT2      =    4
		C_ENTR_CNT2        =    5
		C_DILIG_CD2        =    6
		C_DILIG_NM2        =    7
		C_CNT2             =    8

	end if
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
   	frm1.txtBas_dt.focus
   	frm1.txtBas_dt.Year=strYear
   	frm1.txtBas_dt.Month=strMonth
   	frm1.txtBas_dt.Day=strDay
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
Sub MakeKeyStream(pRow)
    lgKeyStream       = Frm1.txtBas_dt.Text & parent.gColSep                                          'You Must append one character(parent.gColSep)
    lgKeyStream = lgKeyStream & lgUsrIntCd & parent.gColSep
End Sub        

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	Dim dblSum
	
	if frm1.vspdData.MaxRows <=0 then
		exit sub
	end if
	
	With frm1.vspdData

        ggoSpread.Source = frm1.vspdData2
        intIndex = ggoSpread.InsertRow
        
        frm1.vspdData2.Col = 0
        frm1.vspdData2.Text = "합계"
        
        frm1.vspdData2.Col = C_RETIRE_CNT2
        frm1.vspdData2.value = FncSumSheet(frm1.vspdData,C_RETIRE_CNT, 1, .MaxRows, false, -1, -1, "V")
        
        frm1.vspdData2.Col = C_ENTR_CNT2
        frm1.vspdData2.value = FncSumSheet(frm1.vspdData,C_ENTR_CNT, 1, .MaxRows, false, -1, -1, "V")
        
        frm1.vspdData2.Col = C_CNT2
        frm1.vspdData2.value = FncSumSheet(frm1.vspdData,C_CNT, 1, .MaxRows, false, -1, -1, "V")        
	End With	
    With frm1
        ggoSpread.Source = frm1.vspdData2
        .vspdData2.ReDraw = False
		ggoSpread.SpreadLock      C_DEPT_CD2, -1, C_DEPT_CD2
		ggoSpread.SpreadLock      C_DEPT_NM2 , -1, C_DEPT_NM2
		ggoSpread.SpreadLock      C_DEPT_CNT2 , -1, C_DEPT_CNT2
		ggoSpread.SpreadLock      C_RETIRE_CNT2 , -1, C_RETIRE_CNT2	  	  	  	  	  
		ggoSpread.SpreadLock      C_ENTR_CNT2 , -1, C_ENTR_CNT2
		ggoSpread.SpreadLock      C_DILIG_CD2 , -1, C_DILIG_CD2
		ggoSpread.SpreadLock      C_DILIG_NM2 , -1, C_DILIG_NM2
		ggoSpread.SpreadLock      C_CNT2 , -1, C_CNT2	
        .vspdData2.ReDraw = True
    End With
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(strSPD)
	Call initSpreadPosVariables(strSPD)	
	if (strSPD = "A" or strSPD = "ALL") then	
		With frm1.vspdData

		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    

		   .ReDraw = false
		   .MaxCols = C_CNT + 1                                                      ' ☜:☜: Add 1 to Maxcols
		   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
		   .ColHidden = True                                                            ' ☜:☜:

		   .MaxRows = 0
    
			Call GetSpreadColumnPos("A")
		    Call AppendNumberPlace("6","5","0")
		   
		    ggoSpread.SSSetEdit     C_DEPT_CD,        "부서"       , 10,,, 10,2
		    ggoSpread.SSSetEdit     C_DEPT_NM,         "부서"      , 30,,, 40,2
		    ggoSpread.SSSetFloat    C_DEPT_CNT,       "부서인원수" , 16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat    C_RETIRE_CNT,     "퇴사자수"   , 16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat    C_ENTR_CNT,       "입사자수"   , 16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetEdit     C_DILIG_CD,       "근태"       , 16,,, 2,2
		    ggoSpread.SSSetEdit     C_DILIG_NM,       "근태명"     , 23,,, 20,2
		    ggoSpread.SSSetFloat    C_CNT,            "인원수"     , 14,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		    Call ggoSpread.SSSetColHidden(C_DEPT_CD,C_DEPT_CD,True)	
		    Call ggoSpread.SSSetColHidden(C_DILIG_CD,C_DILIG_CD,True)

		   .ReDraw = true
		End With
    End if
    
   	if (strSPD = "B" or strSPD = "ALL") then
		With frm1.vspdData2
	
		    ggoSpread.Source = frm1.vspdData2
		    ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
		   .ReDraw = false

		   .MaxCols = C_CNT2 + 1                                                      ' ☜:☜: Add 1 to Maxcols
		   .Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
		   .ColHidden = True
		  
		   .MaxRows = 0

			Call GetSpreadColumnPos("B")  

		    Call AppendNumberPlace("6","5","0")

		    ggoSpread.SSSetEdit     C_DEPT_CD2,        ""               , 10,,, 7,2
		    ggoSpread.SSSetEdit     C_DEPT_NM2,        ""               , 30,,, 20,2
		    ggoSpread.SSSetEdit     C_DEPT_CNT2,       ""               , 16,,, 20,2
		    ggoSpread.SSSetFloat    C_RETIRE_CNT2,     "퇴사자수"   , 16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetFloat    C_ENTR_CNT2,       "입사자수"   , 16,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
		    ggoSpread.SSSetEdit     C_DILIG_CD2,       ""               , 16,,, 2,2
		    ggoSpread.SSSetEdit     C_DILIG_NM2,       ""               , 23,,, 20,2
		    ggoSpread.SSSetFloat    C_CNT2,            "인원수"     , 14 ,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec

		   Call ggoSpread.SSSetColHidden(C_DEPT_CD2,C_DEPT_CD2,True)
		   Call ggoSpread.SSSetColHidden(C_DILIG_CD2,C_DILIG_CD2,True)
		   
		   .ReDraw = true
		   Call SetSpreadLock 
    
		End With
    End if		
 End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
      ggoSpread.Source = frm1.vspdData2
      ggoSpread.SpreadLockWithOddEvenRowColor()

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
      ggoSpread.SSSetProtected   C_DEPT_CD , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_DEPT_NM , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_DEPT_CNT , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_RETIRE_CNT , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_ENTR_CNT , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_DILIG_CD , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_DILIG_NM , pvStartRow, pvEndRow
      ggoSpread.SSSetProtected   C_CNT , pvStartRow, pvEndRow
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
			C_DEPT_CD         =    iCurColumnPos(1)
			C_DEPT_NM         =    iCurColumnPos(2)
			C_DEPT_CNT        =    iCurColumnPos(3)
			C_RETIRE_CNT      =    iCurColumnPos(4)
			C_ENTR_CNT        =    iCurColumnPos(5)
			C_DILIG_CD        =    iCurColumnPos(6)
			C_DILIG_NM        =    iCurColumnPos(7)
			C_CNT             =    iCurColumnPos(8)
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_DEPT_CD2         =    iCurColumnPos(1)
			C_DEPT_NM2         =    iCurColumnPos(2)
			C_DEPT_CNT2        =    iCurColumnPos(3)
			C_RETIRE_CNT2      =    iCurColumnPos(4)
			C_ENTR_CNT2        =    iCurColumnPos(5)
			C_DILIG_CD2        =    iCurColumnPos(6)
			C_DILIG_NM2        =    iCurColumnPos(7)
			C_CNT2             =    iCurColumnPos(8)
            
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

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '⊙: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
           
	Call InitSpreadSheet("ALL")                                                        'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
    
    Call FuncGetAuth(gStrRequestMenuID, parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
    Call SetToolbar("1100000000001111")										        '버튼 툴바 제어 
    
	Call CookiePage (0)
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
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("X")
    Call DisableToolBar(parent.TBC_QUERY)
	If DBQuery=False Then
	   Call RestoreToolBar()
	  	   Exit Function
	End If
                                                                '☜: Query db data
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
    
    Call MakeKeyStream("X")
    Call DisableToolBar(parent.TBC_SAVE)
	If DBSave=False Then
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
Function FncInsertRow() 
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

    imRow = AskSpdSheetAddRowCount()
    If imRow = "" Then
        Exit Function
    End If
	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow ,imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        
       .vspdData.ReDraw = True
    End With
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

    ggoSpread.Source = frm1.vspdData2 
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
    
    If isEmpty(TypeName(gActiveSpdSheet)) Then
		Exit Sub
	Elseif	UCase(gActiveSpdSheet.id) = "VASPREAD" Then
		ggoSpread.Source = frm1.vspdData2 
		Call ggoSpread.SaveSpreadColumnInf()
	End if

End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : m
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("A")      
    ggoSpread.Source = frm1.vspdData
	Call ggoSpread.ReOrderingSpreadData()

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet("B")      
    ggoSpread.Source = frm1.vspdData2
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

	If LayerShowHide(1)=False Then
		Exit Function
	End If
	
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
    
	If LayerShowHide(1)=False Then
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
 
               Case ggoSpread.InsertFlag                                      '☜: Update추가 
                                                    strVal = strVal & "C" & parent.gColSep 'array(0)
                                                    strVal = strVal & lRow & parent.gColSep
                                                    strVal = strVal & Trim(frm1.txtBas_dt.Text) & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CD	    : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_STRT_DT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_END_DT	: strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REMARK        : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1 

               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                    strVal = strVal & "U" & parent.gColSep
                                                    strVal = strVal & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CD	       : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_STRT_DT	   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_END_DT	   : strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_REMARK           : strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
                    lGrpCnt = lGrpCnt + 1
               
               Case ggoSpread.DeleteFlag                                      '☜: Delete

                                                    strDel = strDel & "D" & parent.gColSep
                                                    strDel = strDel & lRow & parent.gColSep
                    .vspdData.Col = C_EMP_NO	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_CD	    : strDel = strDel & Trim(.vspdData.Text) & parent.gColSep
                    .vspdData.Col = C_DILIG_STRT_DT : strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep	'삭제시 key만								
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
    
    FncDelete = False                                                      '⊙: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '☆:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO,"X","X")		            '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    Call DbDelete															'☜: Delete db data
    
    FncDelete = True                                                        '⊙: Processing is OK
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
    lgIntFlgMode = parent.OPMD_UMODE    
    Call ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
	Call SetToolbar("1100000000011111")									
	Frm1.vspdData.focus
End Function

'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    
    Call InitVariables															'⊙: Initializes local global variables
    Call DisableToolBar(parent.TBC_QUERY)
	If DBQuery=False Then
	   Call RestoreToolBar()
	   Exit Function
	End If

End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
End Function
'========================================================================================================
'	Name : OpenCode()
'	Description : Major PopUp
'========================================================================================================

Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iWhere
	    Case C_DEPT_CNT
	        arrParam(0) = "부서인원수 팝업"			             ' 팝업 명칭 
	        arrParam(1) = " haa010t a (nolock), b_major b (nolock), b_minor c (nolock) "			       	     ' TABLE 명칭 
	        arrParam(2) = ""		                                 ' Code Condition
	        arrParam(3) = ""							             ' Name Cindition
	        frm1.vspddata.col = C_DEPT_CD
	        frm1.vspddata.row = Row
	        arrParam(4) = "a.dept_cd =  " & FilterVar(frm1.vspddata.value, "''", "S") & ""
	        arrParam(4) = arrParam(4) & " and (a.retire_dt >=  " & FilterVar(UNIConvDate(frm1.txtBas_dt.Text), "''", "S") & " or a.retire_dt is null) "
	        arrParam(4) = arrParam(4) & " and a.entr_dt <=  " & FilterVar(UNIConvDate(frm1.txtBas_dt.Text), "''", "S") & ""
	        arrParam(4) = arrParam(4) & " and b.major_cd = " & FilterVar("H0002", "''", "S") & " and b.major_cd = c.major_cd and c.minor_cd = a.roll_pstn "
	        arrParam(5) = ""			                            ' TextBox 명칭 
	        
            arrField(0) = " a.dept_nm "					        ' Field명(0)
            arrField(1) = " a.name "				            ' Field명(1)
            arrField(2) = " c.minor_nm "				        ' Field명(2)
    
            arrHeader(0) = "부서명"				                ' Header명(0)
            arrHeader(1) = "성명"			                    ' Header명(1)
            arrHeader(2) = "직위"			                    ' Header명(2)
            
     '   select a.name, a.dept_nm, c.minor_nm 
     '   from haa010t a (nolock), b_major b (nolock), b_minor c (nolock)
     '   where (a.retire_dt >= '20010510' or a.retire_dt is null)
     '   and a.entr_dt <= '20010510'
     '   and a.dept_cd = '1001'
     '   and b.major_cd = 'H0002'
     '   and b.major_cd = c.major_cd 
     '   and c.minor_cd = a.roll_pstn                     : 신광호대리가 준 spl문 

	    Case C_RETIRE_CNT
	        arrParam(0) = "퇴사자 팝업"			                ' 팝업 명칭 
	        arrParam(1) = " haa010t a (nolock), b_major b (nolock), b_minor c (nolock) "				 		' TABLE 명칭 
	        arrParam(2) = ""		                                ' Code Condition
	        arrParam(3) = ""							            ' Name Cindition
	        frm1.vspddata.col = C_DEPT_CD
	        frm1.vspddata.row = Row
	        arrParam(4) = "a.dept_cd =  " & FilterVar(frm1.vspddata.value, "''", "S") & ""
	        arrParam(4) = arrParam(4) & " and a.retire_dt =  " & FilterVar(UNIConvDate(frm1.txtBas_dt.Text), "''", "S") & " "
	        arrParam(4) = arrParam(4) & " and b.major_cd = " & FilterVar("H0002", "''", "S") & " and b.major_cd = c.major_cd  and c.minor_cd = a.roll_pstn "
	        arrParam(5) = ""			                              ' TextBox 명칭 
	        
            arrField(0) = " a.dept_nm "					          ' Field명(0)
            arrField(1) = " a.name "				              ' Field명(1)
            arrField(2) = " c.minor_nm "				          ' Field명(2)
     
            arrHeader(0) = "부서명"				                  ' Header명(0)
            arrHeader(1) = "성명"			                      ' Header명(1)
            arrHeader(2) = "직위"			                      ' Header명(2)
            
     '   select a.name, a.dept_nm, c.minor_nm 
     '   from haa010t a (nolock), b_major b (nolock), b_minor c (nolock)
     '   where a.retire_dt = '20010510'
     '   and a.dept_cd = '1001'
     '   and b.major_cd = 'H0002'
     '   and b.major_cd = c.major_cd 
     '   and c.minor_cd = a.roll_pstn              : 신광호대리가 준 spl문 

	    Case C_ENTR_CNT
	        arrParam(0) = "입사자 팝업"			                  ' 팝업 명칭 
	        arrParam(1) = " haa010t a (nolock), b_major b (nolock), b_minor c (nolock) "				 	      ' TABLE 명칭 
	        arrParam(2) = ""		                                  ' Code Condition
	        arrParam(3) = ""							              ' Name Cindition
	        frm1.vspddata.col = C_DEPT_CD
	        frm1.vspddata.row = Row
	        arrParam(4) = "a.dept_cd =  " & FilterVar(frm1.vspddata.value, "''", "S") & ""
	        arrParam(4) = arrParam(4) & " and a.entr_dt =  " & FilterVar(UNIConvDate(frm1.txtBas_dt.Text), "''", "S") & " "
	        arrParam(4) = arrParam(4) & " and b.major_cd = " & FilterVar("H0002", "''", "S") & " and b.major_cd = c.major_cd  and c.minor_cd = a.roll_pstn "
	        arrParam(5) = ""			                                 ' TextBox 명칭 
	        
            arrField(0) = " a.dept_nm "					             ' Field명(0)
            arrField(1) = " a.name "				                 ' Field명(1)
            arrField(2) = " c.minor_nm "				             ' Field명(2)
    
            arrHeader(0) = "부서명"				                     ' Header명(0)
            arrHeader(1) = "성명"			                         ' Header명(1)
            arrHeader(2) = "직위"			                         ' Header명(2)
            
    '    select a.name, a.dept_nm, c.minor_nm 
    '    from haa010t a (nolock), b_major b (nolock), b_minor c (nolock)
    '    where a.entr_dt = '20010510'
    '    and a.dept_cd = '1001'
    '    and b.major_cd = 'H0002'
    '    and b.major_cd = c.major_cd 
    '    and c.minor_cd = a.roll_pstn                     : 신광호대리가  spl문 
            

	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
	End If	

End Function
'========================================================================================================
'	Name : OpenCode()
'	Description : Major PopUp
'========================================================================================================

Function OpenCodeDilig(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iWhere
	    Case C_DEPT_CNT
	        arrParam(0) = "부서인원수 팝업"			             ' 팝업 명칭 
	        arrParam(1) = " haa010t a (nolock), hca060t b(nolock), b_major c (nolock), b_minor d (nolock) "			       	     ' TABLE 명칭 
	        arrParam(2) = ""		                                 ' Code Condition
	        arrParam(3) = ""							             ' Name Cindition
	        frm1.vspddata.col = C_DEPT_CD
	        frm1.vspddata.row = Row
	        arrParam(4) =               " a.dept_cd =  " & FilterVar(frm1.vspddata.value, "''", "S") & ""
	        arrParam(4) = arrParam(4) & " and b.dilig_dt =  " & FilterVar(UNIConvDate(frm1.txtBas_dt.Text), "''", "S") & ""
	        frm1.vspddata.col = C_DILIG_CD
	        arrParam(4) = arrParam(4) & " and b.dilig_cd =  " & FilterVar(frm1.vspddata.value, "''", "S") & ""
	        arrParam(4) = arrParam(4) & " and a.emp_no = b.emp_no and c.major_cd = " & FilterVar("H0002", "''", "S") & " and c.major_cd = d.major_cd and d.minor_cd = a.roll_pstn " 
	        arrParam(5) = ""			                            ' TextBox 명칭 

'    select a.name, a.dept_nm, d.minor_nm
'    from haa010t a (nolock), hca060t b(nolock), b_major c (nolock), b_minor d (nolock)
'    where
'    a.emp_no = b.emp_no 
'    and a.dept_cd = '1001'
'    and b.dilig_dt = '20000220'
'    and b.dilig_cd = '33'
'    and c.major_cd = 'H0002'
'    and c.major_cd = d.major_cd 
'    and d.minor_cd = a.roll_pstn                    : 신광호대리가 준 sql문 
           
            arrField(0) = " a.dept_nm "					        ' Field명(0)
            arrField(1) = " a.name "				            ' Field명(1)
            arrField(2) = " d.minor_nm "				        ' Field명(2)
    
            arrHeader(0) = "부서명"				                ' Header명(0)
            arrHeader(1) = "성명"			                    ' Header명(1)
            arrHeader(2) = "직위"			                    ' Header명(2)

	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
	End If	

End Function

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo 변경 이벤트 
'=======================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	
	With frm1.vspdData
	End With

   	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc :
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
	Select Case Col
    End Select    
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx
       
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


'=======================================
'   Event Name : txtIntchng_yymm_dt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtBas_dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtBas_dt.Action = 7
        frm1.txtBas_dt.focus
    End If
End Sub

Sub txtBas_dt_Keypress(Key) 
    If Key = 13 Then
        Call MainQuery
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	Call SetPopupMenuItemInf("0000101111") 	

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
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
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
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , True )    
    Call GetSpreadColumnPos("B")    
End Sub
'========================================================================================================
'   Event Name : vspdData2_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , True )    
    Call GetSpreadColumnPos("B")
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")    
End Sub


'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================

Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim intDeptCnt
    Dim intRetireCnt
    Dim intEntrCnt
    Dim intCnt
    Dim strDiligCd
    
   	frm1.vspdData.Row = Row
    frm1.vspdData.Col = Col
    If Row <= 0 then 
    		Exit Sub
    Else    		
	    Select Case Col
	        Case C_DEPT_CNT
	            intDeptCnt = frm1.vspdData.value
	            If intDeptCnt <>  "" Then                      '인원수가 나오는 cell만 pop up을 띄운다.
	                frm1.vspdData.Col = C_DILIG_CD
	                strDiligCd = frm1.vspdData.value
	                If strDiligCd = "" then
                        Call OpenCode("", C_DEPT_CNT, Row)
                    Else
                        Call OpenCodeDilig("", C_DEPT_CNT, Row)
                    End if
                End if
	        Case C_RETIRE_CNT
	            intRetireCnt = frm1.vspdData.value
	            If intRetireCnt <>  "" Then
                    Call OpenCode("", C_RETIRE_CNT, Row)
                End if
	        Case C_ENTR_CNT
	            intEntrCnt = frm1.vspdData.value
	            If intEntrCnt <>  "" Then
                    Call OpenCode("", C_ENTR_CNT, Row)
                End if
        End Select
    End if
       
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if  
	  
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData2.MaxRows = 0 Then
        Exit Sub
    End If
	
End Sub
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal Col1 , ByVal Col2 )
    frm1.vspdData.Col = Col1
    frm1.vspdData2.ColWidth(Col1) = frm1.vspdData.ColWidth(Col1)
End Sub
'========================================================================================================
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal Col1 , ByVal Col2 )
    frm1.vspdData2.Col = Col1
    frm1.vspdData.ColWidth(Col1) = frm1.vspdData2.ColWidth(Col1)
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
        frm1.vspdData2.LeftCol=NewLeft   	
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
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        frm1.vspdData.LeftCol=NewLeft
	   Exit Sub        
    End If
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
End Sub

'=======================================================================================================
'   Event Name : txtBas_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtBas_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
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

	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>일일인원현황조회</font></td>
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
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
			     		 <TD CLASS=TD5 NOWRAP>기준일</TD>       
				    	 <TD CLASS=TD6 NOWRAP><script language =javascript src='./js/h4011ma1_txtBas_dt_txtBas_dt.js'></script></TD>
			     		 <TD CLASS=TDT NOWRAP></TD>       
				    	 <TD CLASS=TD6 NOWRAP></TD>
			           
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
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h4011ma1_vaSpread_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=64 VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/h4011ma1_vaSpread1_vspdData2.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				
			</TABLE>
		</TD>
	</TR>
	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
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
</HTML>
