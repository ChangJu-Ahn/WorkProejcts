
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 원가요소등록 
'*  3. Program ID           : c1110ma1.asp
'*  4. Program Name         : 원가요소등록 
'*  5. Program Desc         : 원가요소등록 
'*  6. Modified date(First) : 2002/06/03
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Tae Soo
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->
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
Const BIZ_PGM_ID = "C1110MB1.asp"                                      'Biz Logic ASP 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

Dim C_CE_CD		
Dim C_CE_NM		
Dim C_DI_FLAGCD	
Dim C_DI_FLAG		
Dim C_CE_TYPECD	
Dim C_CE_TYPE		
Dim C_MAJOR_FLAG	
Dim C_GainCd		
Dim C_GainNm		

'Const C_SHEETMAXROWS_D  = 100                                          '☆: Server에서 한번에 fetch할 최대 데이타 건수 

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
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	C_CE_CD			= 1
	C_CE_NM			= 2
	C_DI_FLAGCD		= 3
	C_DI_FLAG		= 4
	C_CE_TYPECD		= 5
	C_CE_TYPE		= 6
	C_MAJOR_FLAG	= 7
	C_GainCd		= 8
	C_GainNm		= 9
End Sub


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
       
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

	

'========================================================================================================
Sub InitComboBox()

   
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
    ggoSpread.SetCombo "D" & vbtab & "I" , C_DI_FLAGCD
    ggoSpread.SetCombo "직접" & vbtab & "간접", C_DI_FLAG
    ggoSpread.SetCombo "L" & vbtab & "M" & vbtab & "E", C_CE_TYPECD
    ggoSpread.SetCombo "노무비" & vbtab & "재료비" & vbtab & "경비", C_CE_TYPE
   
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", _
                        " MAJOR_CD = " & FilterVar("G1006", "''", "S") & "  and MINOR_CD in (" & FilterVar("C01", "''", "S") & " ," & FilterVar("C02", "''", "S") & " ," & FilterVar("C03", "''", "S") & " ," & FilterVar("C11", "''", "S") & " ," & FilterVar("C12", "''", "S") & " ) ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
   
  	ggoSpread.SetCombo Replace(lgF0,Chr(11),vbTab), C_GainCd
  	ggoSpread.SetCombo Replace(lgF1,Chr(11),vbTab), C_GainNm
	


	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 
	
	With frm1.vspdData
		For intRow = 1 To .MaxRows
			
			.Row = intRow
		
			.Col = C_DI_FLAGCD
			intIndex = .Value
			.Col = C_DI_FLAG
			.Value = intIndex
			
			.Col = C_CE_TYPECD
			intIndex = .Value
			.Col = C_CE_TYPE
			.Value = intIndex
			
			.Col = C_GainCd
			intIndex = .Value
			.Col = C_GainNm
			.Value = intIndex				
		Next	
	End With
End Sub


'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With frm1.vspdData
	    
	.MaxCols = C_GainNm + 1						
 	
    .Col = .MaxCols							
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021103",,parent.gAllowDragDropSpread    

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 
	
	.ReDraw = false

	Call GetSpreadColumnPos("A")

	' ColumnPosition Header
    ggoSpread.SSSetEdit		C_CE_CD			,"원가요소코드"		,21,,,6,2
	ggoSpread.SSSetEdit	    C_CE_NM			,"원가요소명"		,32,,,30,2
	ggoSpread.SSSetCombo		C_DI_FLAGCD		,"직접/간접 구분"	,16	,0	
	ggoSpread.SSSetCombo		C_DI_FLAG		,"직접/간접 구분"	,16	,0 
	ggoSpread.SSSetCombo		C_CE_TYPECD		,"원가요소TYPE"		,16	,0		
	ggoSpread.SSSetCombo		C_CE_TYPE		,"원가요소TYPE"		,16	,0
	ggoSpread.SSSetCheck		C_MAJOR_FLAG	,"주원가요소여부"	,16	, ,"",true 
	ggoSpread.SSSetCombo		C_GainCd		,"경영손익요소"		,16	,0		
	ggoSpread.SSSetCombo		C_GainNm		,"경영손익요소"		,16	,0

 
    Call ggoSpread.SSSetColHidden(C_DI_FLAGCD,C_DI_FLAGCD,True)
    Call ggoSpread.SSSetColHidden(C_CE_TYPECD,C_CE_TYPECD,True)
    Call ggoSpread.SSSetColHidden(C_GainCd,C_GainCd,True)

	.ReDraw = true
	
'	ggoSpread.SSSetSplit(C_CE_NM)

    Call SetSpreadLock 
    
    End With
    
End Sub


'======================================================================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_CE_CD, -1, C_CE_CD	
	ggoSpread.SSSetRequired C_CE_NM, -1, C_CE_NM	
	ggoSpread.SSSetRequired C_DI_FLAG, -1, C_DI_FLAG	
	ggoSpread.SSSetRequired C_CE_TYPE, -1, C_CE_TYPE
	ggoSpread.SSSetRequired C_GainNm, -1, C_GainNm
	ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1

    .vspdData.ReDraw = True

    End With
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired C_CE_CD		,pvStartRow	,pvEndRow
	ggoSpread.SSSetRequired C_CE_NM		,pvStartRow	,pvEndRow
	ggoSpread.SSSetRequired C_DI_FLAG	,pvStartRow	,pvEndRow	
	ggoSpread.SSSetRequired C_CE_TYPE	,pvStartRow	,pvEndRow	
	ggoSpread.SSSetRequired C_GainNm		,pvStartRow	,pvEndRow	

    .vspdData.ReDraw = True
    
    End With
End Sub


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
			C_CE_CD				= iCurColumnPos(1)
			C_CE_NM				= iCurColumnPos(2)
			C_DI_FLAGCD			= iCurColumnPos(3)    
			C_DI_FLAG			= iCurColumnPos(4)
			C_CE_TYPECD			= iCurColumnPos(5)
			C_CE_TYPE			= iCurColumnPos(6)
			C_MAJOR_FLAG	    = iCurColumnPos(7)
			C_GainCd			= iCurColumnPos(8)
			C_GainNm			= iCurColumnPos(9)
    End Select    
End Sub


'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                        '☜: Clear err status
    
	Call LoadInfTB19029                                                              '☜: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
      
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    'Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
    
    Call InitSpreadSheet                                                             'Setup the Spread sheet

	Call InitVariables
    Call SetDefaultVal

	
    Call SetToolbar("110011010010111")	
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Call InitComboBox
	Call CookiePage (0)                                                              '☜: Check Cookie
			
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
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

    Call InitVariables                                                           '⊙: Initializes local global variables
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
	If DbQuery() = False Then                                                      '☜: Query db data
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
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

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
        .vspdData.Col  = C_CE_CD
        .vspdData.Row  = .vspdData.ActiveRow
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
 	call initdata
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
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function


'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Set gActiveElement = document.ActiveElement   	
    FncNext = True                                                               '☜: Processing is OK
End Function


'========================================================================================================
Function FncExcel() 
    FncExcel = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
Function FncFind() 
    FncFind = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               '☜: Processing is OK
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
	Call InitData()
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
Function DbQuery()

	Dim strVal
	
    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
    
    DbQuery = False                                                              '☜: Processing is NG
	
    Call DisableToolBar(Parent.TBC_QUERY)                                               '☜: Disable Query Button Of ToolBar
    Call LayerShowHide(1)                                                        '☜: Show Processing Message
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001						         
        'strVal = strVal     & "&txtKeyStream="       & lgKeyStream               '☜: Query Key
        strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
        strVal = strVal     & "&lgStrPrevKey="		 &	lgStrPrevKey         '☜: Next key tag
'       strVal = strVal     & "&lgMaxCount="         & CStr(C_SHEETMAXROWS_D)    '☜: Max fetched data at a time
    End With
		
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function


'========================================================================================================
Function DbSave()
		
    Dim lRow        
    Dim lGrpCnt 
    Dim iColSep 
    Dim iRowSep   
	Dim strVal, strDel

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
                    .vspdData.Col = C_CE_CD        : strVal = strVal & Trim(.vspdData.Text)      & iColSep
                    .vspdData.Col = C_CE_NM        : strVal = strVal & Trim(.vspdData.Text)      & iColSep
                    .vspdData.Col = C_DI_FLAGCD    : strVal = strVal & Trim(.vspdData.Text)      & iColSep   
                    .vspdData.Col = C_CE_TYPECD    : strVal = strVal & Trim(.vspdData.Text)      & iColSep  
                    .vspdData.Col = C_MAJOR_FLAG   
                          IF .vspdData.Text = "1" Then
								strVal = strVal & "Y" & iColSep
						  ELSE
								strVal = strVal & "N" & iColSep
						  END IF
                    .vspdData.Col = C_GainCd       : strVal = strVal & Trim(.vspdData.Text)      & iRowSep   
                    lGrpCnt = lGrpCnt + 1
               Case ggoSpread.UpdateFlag                                      '☜: Update
                                                     strVal = strVal & "U"                       & iColSep
                                                     strVal = strVal & lRow                      & iColSep
                    .vspdData.Col = C_CE_CD        : strVal = strVal & Trim(.vspdData.Text)      & iColSep
                    .vspdData.Col = C_CE_NM        : strVal = strVal & Trim(.vspdData.Text)      & iColSep
                    .vspdData.Col = C_DI_FLAGCD    : strVal = strVal & Trim(.vspdData.Text)      & iColSep   
                    .vspdData.Col = C_CE_TYPECD    : strVal = strVal & Trim(.vspdData.Text)      & iColSep   
                    .vspdData.Col = C_MAJOR_FLAG   : 
         					IF .vspdData.Text = "1" Then
								strVal = strVal & "Y" & iColSep
						  ELSE
								strVal = strVal & "N" & iColSep
						  END IF

                    .vspdData.Col = C_GainCd       : strVal = strVal & Trim(.vspdData.Text)      & iRowSep  
                   lGrpCnt = lGrpCnt + 1
               Case ggoSpread.DeleteFlag                                      '☜: Delete
                                                     strDel = strDel & "D"                       & iColSep
                                                     strDel = strDel & lRow                      & iColSep
                    .vspdData.Col = C_CE_CD        : strDel = strDel & Trim(.vspdData.Text)      & iRowSep
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

    Set gActiveElement = document.ActiveElement                                                               '☜: Processing is OK
End Function


'========================================================================================================
Sub DbQueryOk()
	
    lgIntFlgMode = Parent.OPMD_UMODE    
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	Call SetToolbar("110011110011111")	
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
    If DbQuery() = False Then
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
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

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
Sub vspdData_Click(Col, Row)
    
   	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	

    gMouseClickStatus = "SPC"

    Set gActiveSpdSheet = frm1.vspdData

	
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort	Col			'Sort in ascending
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col	,lgSortKey	'Sort in descending
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
'    If Col <= C_CE_NM Or NewCol <= C_CE_NM Then
'        Cancel = True
'        Exit Sub
'    End If
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
           If DbQuery() = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
        End If
    End if
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	With frm1.vspdData
	
		.Row = Row
    
		Select Case Col
		    
			Case  C_DI_FLAG
				.Col = Col
				intIndex = .Value
				.Col = C_DI_FLAGCD
				.Value = intIndex
					
				
			Case  C_DI_FLAGCD
				.Col = Col
				intIndex = .Value
				.Col = C_DI_FLAG
				.Value = intIndex
				
			
			Case C_CE_TYPE
			    .Col = Col
				intIndex = .Value
				.Col = C_CE_TYPECD
				.Value = intIndex	
				
			Case  C_CE_TYPECD
				.Col = Col
				intIndex = .Value
				.Col = C_CE_TYPE
				.Value = intIndex

			Case C_GainNm
			    .Col = Col
				intIndex = .Value
				.Col = C_GainCd
				.Value = intIndex	
		End Select
	End With
End Sub


Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1

	With frm1.vspdData 
	
    ggoSpread.Source = frm1.vspdData
   
    If Row > 0 And Col = C_MAJOR_FLAG Then
        .Col = Col
        .Row = Row
        
	   End If
    
    End With
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>원가요소등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
							<script language =javascript src='./js/c1110ma1_vaSpread1_vspdData.js'></script></TD>
						</TR></TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	
	<TR <%=HEIGHT_TYPE_04%>>
		<TD WIDTH=100% HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hReqStatus" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


