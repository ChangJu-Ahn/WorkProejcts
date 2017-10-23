
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Cost
*  2. Function Name        : 원가기본정보등록 
*  3. Program ID           : C1010MA1
*  4. Program Name         : 원가기본정보등록 
*  5. Program Desc         : 원가기본정보등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/05/31
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Lee Tae Soo
* 10. Modifier (Last)      :
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID = "C1010MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Global 변수/상수 선언 
'========================================================================================================

Dim C_Opt                                  '☆: Spread Sheet의 Column별 상수 
Dim C_OptPop 													
Dim C_OptNm 
Dim C_OptVal 	
Dim C_OptValPop 												
Dim C_OptValNm 


'========================================================================================================
Sub initSpreadPosVariables()  
	C_Opt                = 1                         '☆: Spread Sheet의 Column별 상수 
	C_OptPop 			 = 2										
	C_OptNm				 = 3	
	C_OptVal 		 	 = 4
	C_OptValPop 		 = 5										
	C_OptValNm			 = 6

End Sub


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          


'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgStrPrevKeyIndex = ""                                      '⊙: initializes Previous Key Index
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
	
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub
	

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   Select Case pOpt
       Case "MQ"
                  lgKeyStream = Frm1.CboCostType.Value  & Parent.gColSep & frm1.txtPlantCd.Value & Parent.gColSep     'You Must append one character(Parent.gColSep)
       Case "MN"
                  lgKeyStream = Frm1.hCboCostType.Value & Parent.gColSep & frm1.hPlantCd.Value & Parent.gColSep      'You Must append one character(Parent.gColSep)
   End Select                 
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

	
'========================================================================================================
Sub InitComboBox()
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                         " MAJOR_CD = " & FilterVar("C0001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    Call SetCombo2(frm1.CboCostType ,lgF0  ,lgF1  ,Chr(11))
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub InitData()
'	Dim intRow
'	Dim intIndex 

'	With frm1.vspdData
'		For intRow = 1 To .MaxRows			
'			.Row = intRow
'			.Col = C_StudyOnOffCd  :  intIndex = .Value             ' .Value means that it is index of cell,not value in combo cell type
'			.Col = C_StudyOnOffNm  :  .Value = intindex					
'		Next	
'	End With
End Sub


'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	With frm1.vspdData
	
       .MaxCols   = C_OptValNm+1                                                  ' ☜:☜: Add 1 to Maxcols
	   .Col       = .MaxCols        : .ColHidden = True

        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    

		Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
		
	   .ReDraw = false
	
       'Call AppendNumberPlace("6","4","0")

		Call GetSpreadColumnPos("A")
		
		'ColumnPosition Header
		ggoSpread.SSSetEdit		C_Opt,		"선택사항",	22,,,	10,2
		ggoSpread.SSSetButton	C_OptPop    
		ggoSpread.SSSetEdit		C_OptNm,	"선택사항명",35    
		ggoSpread.SSSetEdit		C_OptVal,	"선택값",	22,,,	10,2
		ggoSpread.SSSetButton	C_OptValPop
		ggoSpread.SSSetEdit		C_OptValNm,	"선택값명",	35

		call ggoSpread.MakePairsColumn(C_Opt,C_OptPop)
		call ggoSpread.MakePairsColumn(C_OptVal,C_OptValPop)

	   .ReDraw = true
	
       Call SetSpreadLock 
    
    End With
    
End Sub


'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False 
	ggoSpread.SpreadLock	C_Opt		, -1, C_Opt
	ggoSpread.SpreadLock	C_OptPop	, -1, C_OptPop
	ggoSpread.SpreadLock	C_OptNm		, -1, C_OptNm    
	ggoSpread.SSSetRequired C_OptVal	, -1, C_OptVal
    ggoSpread.SpreadLock	C_OptValNm	, -1, C_OptValNm
    ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired		C_Opt		,pvStartRow	,pvEndRow
    ggoSpread.SSSetProtected		C_OptNm		,pvStartRow	,pvEndRow
    ggoSpread.SSSetRequired  	C_OptVal	,pvStartRow	,pvEndRow
    ggoSpread.SSSetProtected		C_OptValNm	,pvStartRow	,pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Name : SubSetErrPos
' Desc : This method set focus to position of error
'      : This method is called in MB area
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


'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Opt				= iCurColumnPos(1)
            C_OptPop			= iCurColumnPos(2)
            C_OptNm				= iCurColumnPos(3)
            C_OptVal			= iCurColumnPos(4)
            C_OptValPop			= iCurColumnPos(5)
            C_OptValNm			= iCurColumnPos(6)
    End Select    
End Sub


'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================


'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
      
    'Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
    
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal

	frm1.cboCostType.focus
	Call SetToolbar("110011010000111")	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
	Call CookiePage (0)                                                              '☜: Check Cookie
	Set gActiveElement = document.activeElement			
End Sub
	

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub


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
    
    Call ggoOper.ClearField(Document, "2")
	    										 '☜: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '⊙: Initializes local global variables
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	If DbQuery("MQ") = False Then                                                      '☜: Query db data
       Exit Function
    End If
	
	If Err.number = 0 Then	
       FncQuery = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
	

'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	If Err.number = 0 Then	
       FncNew = True                                                              '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function
	


'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then                                      '☜:match pointer
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
       Exit Function
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	with frm1
		if .cboCostType.value = "C2100" then
			.txtPlantCd.value = "ACTL"
		end if
	end with

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

     If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'In Multi, You need not to implement this area
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
     If Err.number = 0 Then	
       FncDelete = True                                                           '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function FncCopy()
	Dim IntRetCD
	
	On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                              '☜: Processing is NG

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
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	' Clear key field
	'---------------------------------------------------------------------------------------------------- 

	With Frm1
        .vspdData.Col  = C_Opt
        .vspdData.Row  = .vspdData.ActiveRow
        .vspdData.Text = ""
        
        .vspdData.Col  = C_OptNm
        .vspdData.Text = ""

	End With

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function FncCancel() 
    On Error Resume Next
    
    Dim iDx
	
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
 
    
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
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

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow  .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
        .vspdData.ReDraw = True
    End With
	'------ Developer Coding part (Start ) --------------------------------------------------------------
	Call SetToolbar("110011110001111")                                            '☆: Developer must customize 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    
    Set gActiveElement = document.ActiveElement  
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
     
End Function


'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call SetToolbar("110011110001111")                                            '☆: Developer must customize 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	If Err.number = 0 Then	
       FncDeleteRow = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call parent.FncPrint()                                                       '☜: Protect system from crashing
    If Err.number = 0 Then	 
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
                                                         '☜: Processing is OK
End Function


'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then	 
       FncPrev = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'
'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then	 
       FncNext = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncExport(Parent.C_MULTI)

   If Err.number = 0 Then	 
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                                              '☜: Processing is OK
End Function


'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncFind(Parent.C_MULTI, True)

    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                                               '☜: Processing is OK
End Function


'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub




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
'	Call InitData()
End Sub



'========================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")		             '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================


'========================================================================================================
Function DbQuery(pDirect)

	Dim strVal
	
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
    
    DbQuery = False                                                              '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
	Call MakeKeyStream(pDirect)
	
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
   
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="		 & lgStrPrevKey				'☜: Next key tag
'       strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)    '☜: Max fetched data at a time
    End With
		
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    If Err.number = 0 Then	 
       DbQuery = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function DbSave()
		
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
    Dim iColSep 
    Dim iRowSep   

    On Error Resume Next
    DbSave = False                                                               '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Call DisableToolBar(Parent.TBC_SAVE)                                                '☜: Disable Save Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
		
    Frm1.txtMode.value        = Parent.UID_M0002                                        '☜: Delete
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	
    ggoSpread.Source = frm1.vspdData

    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
    iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	
	
	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
 
               Case ggoSpread.InsertFlag                                      '☜: Update
                                                     strVal = strVal & "C"                       & iColSep
                                                     strVal = strVal & lRow                      & iColSep
                    .vspdData.Col = C_Opt          : strVal = strVal & Trim(.vspdData.Text)      & iColSep
                    .vspdData.Col = C_OptVal       : strVal = strVal & Trim(.vspdData.Text)      & iRowSep  
                    
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                     strVal = strVal & "U"                       & iColSep
                                                     strVal = strVal & lRow                      & iColSep
                    .vspdData.Col = C_Opt          : strVal = strVal & Trim(.vspdData.Text)      & iColSep
                    .vspdData.Col = C_OptVal       : strVal = strVal & Trim(.vspdData.Text)      & iRowSep   

                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                     strDel = strDel & "D"                       & iColSep
                                                     strDel = strDel & lRow                      & iColSep
                    .vspdData.Col = C_Opt          : strDel = strDel & Trim(.vspdData.Text)      & iColSep
                    .vspdData.Col = C_OptVal       : strDel = strDel & Trim(.vspdData.Text)      & iRowSep   

                    lGrpCnt = lGrpCnt + 1
           End Select
       Next

	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    If Err.number = 0 Then	 
       DbSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'In Multi, You need not to implement this area

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
   If Err.number = 0 Then	 
       DbDelete = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement                                                       
End Function


'========================================================================================================
Sub DbQueryOk()
	
    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call SetToolbar("110011110001111")                                            '☆: Developer must customize
	Frm1.vspdData.Focus
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitData()
	Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
End Sub
	

'========================================================================================================
Sub DbSaveOk()
    Call InitVariables															     '⊙: Initializes local global variables
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
    Call ggoOper.ClearField(Document, "2")										     '⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If DbQuery("MQ") = False Then
       Call RestoreToolBar()
       Exit Sub
    End if
    Set gActiveElement = document.ActiveElement   
End Sub
	

'========================================================================================================
Sub DbDeleteOk()
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""											
	arrParam(4) = ""												
	arrParam(5) = "공장"							
	
    arrField(0) = "PLANT_CD"						
    arrField(1) = "PLANT_NM"						
    
    arrHeader(0) = "공장코드"				
    arrHeader(1) = "공장명"			
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If
		
End Function


'========================================================================================================
Function OpenPopUp(Byval strCode, Byval strCode1, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
		Case 0
			arrParam(0) = "선택사항팝업"  	
			arrParam(1) = "B_Minor"					
			arrParam(2) = strCode
			arrParam(3) = ""								
			arrParam(4) = "Major_Cd= " & FilterVar(strCode1, "''", "S") & " "				
			arrParam(5) = "선택사항"    				

			arrField(0) = "MINOR_CD"				
			arrField(1) = "MINOR_NM"					
    
			arrHeader(0) = "선택사항"	  
			arrHeader(1) = "선택사항명"		
		Case 1
			arrParam(0) = "선택값팝업"  	
			arrParam(1) = "B_Minor"	 				
			arrParam(2) = strCode
			arrParam(3) = ""									
			arrParam(4) = "Major_Cd= " & FilterVar(strCode1, "''", "S") & " "			
			arrParam(5) = "선택값"	

			arrField(0) = "MINOR_CD"
			arrField(1) = "MINOR_NM"
    
			arrHeader(0) = "선택값"	
			arrHeader(1) = "선택값명"	

	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				.vspdData.Col = C_Opt
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_OptNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
				
				.vspdData.Col = C_Opt
				.vspdData.Action = 0			
				
			Case 1
				.vspdData.Col = C_OptVal
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_OptValNm
				.vspdData.Text = arrRet(1)
				Call vspddata_Change(.vspddata.col, .vspddata.row)
			
				.vspdData.Col = C_OptVal
				.vspdData.Action = 0			
				
		End Select

	End With
	
End Function

Function SetPlant(byval arrRet)
	frm1.txtPlantCd.focus
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)		
	
End Function

Function cboCostType_onChange()      
with frm1
	if .cboCostType.value = "C1100" then         
		TitlePlant.innerHTML = "공장"
		Call ElementVisible(.txtPlantCd, 1)
		Call ElementVisible(.btnPlantCd, 1)
		Call ElementVisible(.txtPlantNm, 1)
		.txtPlantCd.value = ""
		.txtPlantNm.value = ""
	else if .cboCostType.value = "C2100" then
		TitlePlant.innerHTML = ""
		Call ElementVisible(.txtPlantCd, 0)
		Call ElementVisible(.btnPlantCd, 0)
		Call ElementVisible(.txtPlantNm, 0)
		.txtPlantCd.value = "ACTL"
		end if
	end if
end with
End Function


'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim StrCode1
	Dim intRetCD
	
	With frm1
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_OptPop
				.vspdData.Col = C_Opt
				.vspdData.Row = Row
				
				If Trim(frm1.cboCostType.Value) = "" Then
					intRetCD =  DisplayMsgBox("230117","x","x","x")
					frm1.cboCostType.focus
					Exit Sub
				End If
					
				Call OpenPopup(.vspdData.Text,frm1.cboCostType.value, 0)

			Case C_OptValPop        
				.vspdData.Col = C_Opt
				.vspdData.Row = Row
				  
				If .vspdData.Text = "" Then
					intRetCD = DisplayMsgBox("230118","x","x","x")
					Exit Sub
				End If
				
				StrCode1 = Trim(.vspdData.Text)
				
				.vspdData.Col = C_OptVal
				Call OpenPopup(.vspdData.Text, StrCode1, 1)
		End Select
           	Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
    End With
    
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

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
	
End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)

    Dim iDx
       
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
             
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	
	
	gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col				'Sort in scending
            lgSortKey = 2
        Else
            ggoSpread.SSSort	Col	,lgSortKey  'Sort in descending
            lgSortKey = 1
        End If
    Else
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'	frm1.vspdData.Col = C_MajorCd
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
    
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
    'If Col <= C_SNm Or NewCol <= C_SNm Then
     '   Cancel = True
      '  Exit Sub
   ' End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
           If DbQuery("MN") = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" bgColor=White text=Black>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>원가기본정보등록</font></td>
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
									<TD CLASS="TD5">원가종류</TD>
									<TD CLASS="TD6"><SELECT ID="cboCostType" NAME="cboCostType" ALT="원가종류" tag="12X" STYLE="WIDTH:82px:"></SELECT></TD>
									<TD CLASS="TD5" ID="Titleplant">공장</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPlant()" ID="btnPlantCd">
										 <INPUT TYPE=TEXT ID="txtPlantNm" NAME="txtPlantNm" SIZE=25 tag="14X">
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
								<TD HEIGHT="100%" NOWRAP>
								<script language =javascript src='./js/c1010ma1_I381981438_vspdData.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hCboCostType" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

