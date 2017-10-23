
<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Cost
*  2. Function Name        : C/C별 배부규칙등록 
*  3. Program ID           : C1703MA1
*  4. Program Name         : C/C별 배부규칙등록 
*  5. Program Desc         : Multi Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/06/05
*  8. Modified date(Last)  : 2006/03/09
*  9. Modifier (First)     : Lee Tae Soo
* 10. Modifier (Last)      : Jeong Yong Kyun
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
Const BIZ_PGM_ID       = "C1703MB1.asp"                                      'Biz Logic ASP for Mmulti #1
Const BIZ_PGM_QRY_ID2  = "C1703MB9.asp"                                      'Biz Logic ASP for Mmulti #2
Const BIZ_PGM_QRY_ID3  = "cb009mb1.asp"
Const BIZ_PGM_JUMP_ID1 = "c1704ma1"

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================

'☜:--------Spreadsheet #1-----------------------------------------------------------------------------   
Dim C_Cost_Cd  
Dim C_Cost_PB  
Dim C_Cost_Nm  
Dim C_DstbFctr_Cd  
Dim C_DstbFctr_PB  
Dim C_DstbFctr_Nm  
Dim C_Dstb_Order  


'☜:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim C_CheckFlag  
Dim C_RecvCost_Cd  
Dim C_RecvCost_Nm  

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgCurrRow
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
Sub initSpreadPosVariables()  
	C_Cost_Cd			= 1
	C_Cost_PB			= 2
	C_Cost_Nm			= 3
	C_DstbFctr_Cd		= 4
	C_DstbFctr_PB		= 5
	C_DstbFctr_Nm		= 6
	C_Dstb_Order		= 7
End Sub


'========================================================================================================
Sub initSpreadPosVariables1()  
	C_CheckFlag		= 1
	C_RecvCost_Cd		= 2
	C_RecvCost_Nm		= 3
End Sub

'========================================================================================================
Sub InitVariables()
	lgBlnFlgChgValue   = False								    '⊙: Indicates that no value changed
    lgSortKey          = 1                                      '⊙: initializes sort direction
		
	lgIntFlgMode = Parent.OPMD_CMODE  
End Sub

'========================================================================================================

Sub SetDefaultVal()
    Call ggoOper.ClearField(Document, "1") 
End Sub
	
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

	<% Call loadInfTB19029A("I", "*", "COOKIE", "MA") %>
End Sub

'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Const CookieSplit = 4877						
    
    Dim IntRetCD 
	Dim strTemp, arrVal
	
	If Kubun = 1 Then									
		If frm1.vspdData.maxrows < 1 Then Exit Sub

		frm1.vspddata.Row = frm1.vspddata.ActiveRow
		frm1.vspddata.Col = C_Cost_Cd
		
		If frm1.vspddata.Row <> 0 Then	
	   		WriteCookie CookieSplit , frm1.txtVerCd.value  & Parent.gRowSep & frm1.vspddata.value
		End If
	ElseIf Kubun = 0 Then								
		strTemp = ReadCookie(CookieSplit)
			
		If strTemp = "" Then Exit Sub

		arrVal = Split(strTemp, Parent.gRowSep)

		If arrVal(0) = "" Then Exit Sub

		frm1.txtVerCd.value =  arrVal(0)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Sub
		End If

		Call MainQuery()		
			
		WriteCookie CookieSplit , ""
	End If
End Sub
	
'========================================================================================================
Sub InitComboBoxGrid()
	ggoSpread.source = frm1.vspdData2
	ggoSpread.SetCombo "Y" & vbtab & "N" , C_Checkflag
End Sub

'========================================================================================================
Sub InitData(pSpreadSheetNo)
	Dim intRow
	Dim intIndex 

    With frm1.vspdData
		Select Case UCase(pSpreadSheetNo)
			Case  "M"
                 For intRow = 1 To .MaxRows		                                  'Not from zero	
                       .Row = intRow
                       .Col = C_PrivatePublicCD  :  intIndex = .Value             ' .Value means that it is index of cell,not value in combo cell type
                       .Col = C_PrivatePublicNM  :  .Value = intindex					
                       .Col = C_CloseYNCD        :  intIndex = .Value             ' .Value means that it is index of cell,not value in combo cell type
                       .Col = C_CloseYNNM        :  .Value = intindex					
                 Next	
			Case  "M1"
		        For intRow = 1 To .MaxRows			
		              .Row = intRow  
		              .Col = C_StudyOnOffCd     :  intIndex = .Value             ' .Value means that it is index of cell,not value in combo cell type
		              .Col = C_StudyOnOffNm     :  .Value = intindex					
		        Next	
		End Select 
	End With
End Sub


'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Select Case UCase(pvSpdNo)
		Case "A" 
			Call initSpreadPosVariables()    
		
			With frm1.vspdData
				.MaxCols = C_Dstb_Order + 1
				.Col = .MaxCols
				.ColHidden = True

	   			ggoSpread.Source = frm1.vspdData
	   			ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread  
				
				ggoSpread.ClearSpreadData
				
				.ReDraw = False

				Call GetSpreadColumnPos("A")
				Call AppendNumberPlace("6","3","0")
				
				ggoSpread.SSSetEdit C_Cost_Cd, "코스트센타코드", 12,,,10,2
				ggoSpread.SSSetButton C_Cost_PB
				ggoSpread.SSSetEdit C_Cost_Nm, "코스트센타명", 21
				ggoSpread.SSSetEdit C_DstbFctr_CD, "배부요소코드", 10,,,2,2
				ggoSpread.SSSetButton C_DstbFctr_PB
				ggoSpread.SSSetEdit C_DstbFctr_NM, "배부요소명", 15
				ggoSpread.SSSetFloat C_Dstb_Order,"배부단계",11,"6"  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"P"

				call ggoSpread.MakePairsColumn(C_Cost_Cd,C_Cost_PB)
				call ggoSpread.MakePairsColumn(C_DstbFctr_CD,C_DstbFctr_PB)	
		    		
				.ReDraw = True
	        End with
		Case "B" 
			Call initSpreadPosVariables1()    
	
			With frm1
				.vspdData2.MaxCols = C_RecvCost_Nm	+1
				.vspdData2.Col = .vspdData2.MaxCols						
				.vspdData2.ColHidden = True 
								    
				ggoSpread.Source = .vspdData2
				ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread  
			
				ggoSpread.ClearSpreadData
			
				.vspdData2.ReDraw = false         

				Call GetSpreadColumnPos("B")   

				ggoSpread.SSSetCheck		C_Checkflag ,"대상여부",10 , ,"",true    
				ggoSpread.SSSetEdit		C_RecvCost_Cd, "배부대상", 10,,,10,2
				ggoSpread.SSSetEdit		C_RecvCost_Nm, "코스트센타명", 27

				.vspdData2.ReDraw = True
			End With 
     End Select      

	SetSpreadLock "I", 0, -1, -1
End Sub

'======================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow  , ByVal lRow2 )
    Dim objSpread

    With frm1
		Select Case Index
			Case 0
				ggoSpread.Source = .vspdData

				Set objSpread = .vspdData

				lRow2 = objSpread.MaxRows
				objSpread.Redraw = False

				ggoSpread.SpreadLock C_Cost_Cd, lRow, C_Cost_Cd, lRow2
				ggoSpread.SpreadLock C_Cost_PB, lRow, C_Cost_PB, lRow2	
		        ggoSpread.SpreadLock C_Cost_Nm, lRow, C_Cost_Cd, lRow2               
				ggoSpread.SpreadLock C_DstbFctr_Nm, lRow, C_DstbFctr_Nm, lRow2
		End Select
    
		ggoSpread.SSSetRequired C_DstbFctr_Cd, -1, C_DstbFctr_Cd
		ggoSpread.SSSetRequired C_Dstb_Order, -1, C_Dstb_Order
		ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1  

		objSpread.Redraw = True
		Set objSpread = Nothing
    End With
End Sub


'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False
										'Col				Row         Row2
		ggoSpread.SSSetRequired	C_Cost_Cd		,pvStartRow	,pvEndRow
		ggoSpread.SSSetProtected C_Cost_Nm		,pvStartRow	,pvEndRow	
		ggoSpread.SSSetRequired	C_DstbFctr_Cd	,pvStartRow	,pvEndRow 
		ggoSpread.SSSetProtected C_DstbFctr_Nm	,pvStartRow	,pvEndRow	
		ggoSpread.SSSetRequired	C_Dstb_Order	,pvStartRow	,pvEndRow

		.vspdData.ReDraw = True
    End With
End Sub

'======================================================================================================
Sub SetSpread2Color()
	Dim strStartRow, strEndRow

    With frm1
		strStartRow = 1
		strEndRow	= .vspdData2.MaxRows

		ggoSpread.Source	= .vspdData2
		.vspdData2.ReDraw	= False
		
		ggoSpread.SSSetProtected   C_RecvCost_Cd	,strStartRow, strEndRow
        ggoSpread.SSSetProtected   C_RecvCost_Nm	,strStartRow, strEndRow
 		
		.vspdData2.ReDraw = True
    End With
End Sub

'======================================================================================================
Sub SubSetErrPos(pSpreadSheetNo,iPosArr)
    Dim iDx
    Dim iRow

    iPosArr = Split(iPosArr,Parent.gColSep)

    Select Case pSpreadSheetNo
        Case "M"
        Case "M1"
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
    End  Select
End Sub

'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_Cost_Cd			= iCurColumnPos(1)
			C_Cost_PB			= iCurColumnPos(2)
			C_Cost_Nm			= iCurColumnPos(3)    
			C_DstbFctr_Cd       = iCurColumnPos(4)
			C_DstbFctr_PB		= iCurColumnPos(5)
			C_DstbFctr_Nm		= iCurColumnPos(6)
			C_Dstb_Order		= iCurColumnPos(7)
		Case "B"
			ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_CheckFlag  			= iCurColumnPos(1)
			C_RecvCost_Cd  			= iCurColumnPos(2)
			C_RecvCost_Nm  			= iCurColumnPos(3)    
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
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											 '⊙: Lock Field
	Call InitVariables
    Call SetDefaultVal
    Call InitSpreadSheet("A")
    Call InitSpreadSheet("B")                                                             'Setup the Spread sheet
	Call SetToolbar("110011010010111")	
	
	frm1.txtVerCd.focus 
    frm1.vspdData3.MaxRows = 0
    frm1.vspdData3.MaxCols = 5    

    Call InitComboBoxGrid
	Call CookiePage (0)                                                              '☜: Check Cookie

   	Set gActiveElement = document.activeElement			
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

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
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

    If Not chkField(Document, "1") Then									         '☜: This function check required field
		Exit Function
    End If

    Call InitVariables                                                           '⊙: Initializes local global variables

    If DbQuery() = False Then                                                      '☜: Query db data
		Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncQuery = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
 
    Set gActiveElement = document.ActiveElement   
    FncNew = True																 '☜: Processing is OK
End Function

'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Set gActiveElement = document.ActiveElement       
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    On Error Resume Next
    
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

    If DbSave = False Then                                                       '☜: Query db data
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    FncSave = True                                                              '☜: Processing is OK
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
	
	With Frm1
        .vspdData.Col  = C_Cost_CD  
        .vspdData.Row  = .vspdData.ActiveRow
        .vspdData.Text = ""
        .vspdData.Col  = C_Cost_Nm  
        .vspdData.Row  = .vspdData.ActiveRow
        .vspdData.Text = ""
        .vspdData.Col  = C_DstbFctr_Cd  
        .vspdData.Row  = .vspdData.ActiveRow
        .vspdData.Text = ""
        .vspdData.Col  = C_DstbFctr_Nm  
        .vspdData.Row  = .vspdData.ActiveRow
        .vspdData.Text = ""
	End With
	
	frm1.vspdData2.MaxRows =0

    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncCancel() 
    Dim iCostCd
    
    FncCancel = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	If frm1.vspdData.MaxRows < 1 Then Exit Function

	With frm1.vspdData
		.Row = .ActiveRow
        .Col = 0

		If .row = 0 Then 
			Exit Function
		End If

		If .Text = ggoSpread.InsertFlag Then
		    .Col = C_Cost_Cd
		    DeleteHSheet(.Text)
        ElseIf .Text = ggoSpread.UpdateFlag Then
		    .Col = C_Cost_Cd
		     CanCelHSheet(.Text)
		End If

		.Col = C_Cost_Cd
		iCostCd = .Text

		ggoSpread.Source = frm1.vspdData	
		ggoSpread.EditUndo

		If .activerow = 0 Then 
			Exit Function
		End If

		If .Text = ggoSpread.InsertFlag Then
            .Col = C_Cost_Cd
            If .text <> "" Then
				frm1.hCostCd.value = .Text
				frm1.vspdData2.MaxRows = 0
				If DbQuery3(.ActiveRow) = False Then
					Exit Function
				End If
			End If
        Else
            .Col = C_Cost_Cd
            frm1.hCostCd.value = .Text
            frm1.vspdData2.MaxRows = 0

            If DbQuery2(.ActiveRow) = False Then
				Exit Function
			End If
        End if
    End With
     
    Set gActiveElement = document.ActiveElement   
    FncCancel = False                                                            '☜: Processing is OK
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

	frm1.vspdData2.MaxRows= 0

    Set gActiveElement = document.ActiveElement

    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If   
End Function

'========================================================================================================
Function FncDeleteRow()
    Dim lDelRows
    Dim DelCostCd

    FncDeleteRow = False                                                         '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If Frm1.vspdData.MaxRows < 1 Then
       Exit function
	End If

    With Frm1.vspdData
    	.focus

    	If .maxrows = 0 Then Exit Function

        .Row = .ActiveRow
		.Col = C_Cost_Cd 
        DelCostCd = .Text
    	
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With

    DeleteHsheet DelCostCd

    Set gActiveElement = document.ActiveElement   
    FncDeleteRow = True                                                          '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrint()
    FncPrint = False                                                             '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

	Call parent.FncPrint()                                                       '☜: Protect system from crashing
    FncPrint = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Function FncPrev() 
    FncPrev = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    Set gActiveElement = document.ActiveElement   	
    FncPrev = True                                                               '☜: Processing is OK
End Function

'========================================================================================================
Function FncNext() 
    FncNext = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

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
' Name : FncFind
' Desc : This function is called by MainFind in Common.vbs
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
	Dim indx
	Dim lngActRow1, lngActRow2
	Dim lngActCol1, lngActCol2

	lngActRow1 = frm1.vspdData.ActiveRow
	lngActCol1 = frm1.vspdData.ActiveCol
	lngActRow2 = frm1.vspdData2.ActiveRow
	lngActCol2 = frm1.vspdData2.ActiveCol

	If gActiveSpdSheet.Name <> "" Then
		For indx = 0 To frm1.vspdData.MaxRows
			frm1.vspdData.Row = indx
			frm1.vspdData.Col = 0
			
			Select Case Trim(UCase(gActiveSpdSheet.Name))
				Case "VSPDDATA"
					frm1.vspdData.Row = lngActRow1 
					frm1.vspdData.Col = lngActCol1
					frm1.vspdData.Action = 0
					
			   		
				Case "VSPDDATA2"
					frm1.vspdData2.Row = lngActRow2
					frm1.vspdData2.Col = lngActCol2
					frm1.vspdData2.Action = 0
			End Select
		Next
	End If

	ggoSpread.Source = gActiveSpdSheet
	Call ggoSpread.RestoreSpreadInf()
 
    Select Case Trim(UCase(gActiveSpdSheet.Name))
		Case "VSPDDATA"
			Call InitSpreadSheet("A")
			Call ggoSpread.ReOrderingSpreadData()
			Call InitData()
			SetSpreadLock "I", 0, -1, -1
			
		Case "VSPDDATA2"
			Call ggoSpread.RestoreSpreadInf()
			Call InitSpreadSheet("B")			' 그리드2 초기화 
			Call InitComboBoxGrid()
			Call ggoSpread.ReOrderingSpreadData()
			Call SetSpread2Color()
	End Select

	If frm1.vspdData2.MaxRows <= 0 Then	
		Call DbQuery2(frm1.vspdData.ActiveRow)
	End If
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
Function ExeCopy() 
	Dim IntRetCD
	Dim strVal

	On Error Resume Next
	Err.Clear 

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        IntRetCD = DisplayMsgBox("900002","X","X","X")                               
        Exit Function
    End If

	If Not chkField(Document, "2") Then
		Exit Function
	End If	

	IntRetCD = DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X")

	If IntRetCD = vbNo Then
		Exit Function
	End If

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	ExeCopy = False

	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003
	strVal = strVal & "&txtVerCd=" & Trim(frm1.htxtVerCd.value)
	strVal = strVal & "&txtNewVerCd=" & Trim(frm1.txtNewVerCd.value)

	Call RunMyBizASP(MyBizASP, strVal)

	ExeCopy = True
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

	frm1.vspdData.MaxRows = 0
    frm1.vspdData2.MaxRows = 0
    frm1.vspdData3.MaxRows = 0 
    
    With frm1	    				
	    If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&txtVerCd=" & Trim(.htxtVerCd.value)
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	    Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001 
			strVal = strVal & "&txtVerCd=" & Trim(.txtVerCd.value)
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		End If

		Call RunMyBizASP(MyBizASP, strVal)
    End With	

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic

    DbQuery = True                                                               '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbQuery2(ByVal Row)
	Dim strVal
	Dim lngRows

    Err.Clear                                                                    '☜: Clear err status
    On Error Resume Next
    
    DbQuery2 = False                                                              '☜: Processing is NG

	
	With frm1
	    .vspdData.row = Row
	    .vspdData.col = C_Cost_Cd
	    .hCostCd.Value = .vspdData.Text

		If Trim(.hCostCd.Value) = "" Then
	        Exit Function
	    End If
	    
		If CopyFromData(.hCostCd.Value) = True Then
			    Exit Function
		End If
		    	
		Call LayerShowHide(1)
	
		DbQuery2 = False
		
		.vspdData.Row = Row
		.vspdData.Col = C_Cost_Cd
	
	    If lgIntFlgMode = Parent.OPMD_UMODE Then 
			strVal = BIZ_PGM_QRY_ID2 & "?txtMode=" & Parent.UID_M0001			
			strVal = strVal & "&txtVerCd=" & Trim(.htxtVerCd.value)		
     		strVal = strVal & "&txtCostCd=" & .hCostCd.Value    		
    		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
	    Else
			strVal = BIZ_PGM_QRY_ID2 & "?txtMode=" & Parent.UID_M0001	
			strVal = strVal & "&txtVerCd=" & Trim(.txtVerCd.value)
    		strVal = strVal & "&txtCostCd=" & .vspdData.text    		
    		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
	    End If
	End With

	Call RunMyBizASP(MyBizASP, strVal)
	
	DbQuery2 = True
End Function

'========================================================================================================
Function DbQueryOk2()
	With frm1
        ggoSpread.Source = .vspdData2
		SetSpread2Color 
    	End With
End Function

'========================================================================================================
Function DbQuery3(ByVal Row)
	Dim strVal
	Dim lngRows
	Dim i 

	DbQuery3 = False

	With frm1
	    .vspdData.row = Row
	    .vspdData.col = C_Cost_Cd
	    .hCostCd.Value = .vspdData.Text

		If CopyFromData(.hCostCd.Value) = True Then
		    Exit Function
		End If

		Call LayerShowHide(1)
	
		.vspdData.Row = Row
		.vspdData.Col = C_Cost_Cd
	    
   		strVal = BIZ_PGM_QRY_ID3 & "?txtMode=" & Parent.UID_M0001			
 		strVal = strVal & "&txtItemCd=" & .vspdData.Text
		strVal = strVal & "&txtMaxRows=" & .vspdData2.MaxRows
	End With

	Call RunMyBizASP(MyBizASP, strVal)
	
	DbQuery3 = True
End Function

'========================================================================================================
Function DbQueryOk3()
	Dim i

	With frm1
		ggoSpread.Source = .vspdData2
		SetSpread2Color 
    End With
End Function

'========================================================================================================
Function DbSave()
    Dim lngRows
    Dim itemRows        
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
	Frm1.txtFlgMode.value = lgIntFlgMode									
		
    ggoSpread.Source = frm1.vspdData

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	

	With frm1.vspdData    
		For lngRows = 1 To frm1.vspdData.MaxRows
			.Row = lngRows
			.Col = 0

			If .Text = ggoSpread.InsertFlag Then
				strVal = strVal & "CREATE" & iColSep & lngRows & iColSep
				.Col = C_Cost_Cd	
				strVal = strVal & Trim(.Text) & iColSep
				            
				.Col = C_DstbFctr_Cd	
				strVal = strVal & Trim(.Text) & iColSep
	
				.Col = C_Dstb_Order
				strVal = strVal & Trim(.Text) & iRowSep
				lGrpCnt = lGrpCnt + 1
			ElseIf .Text = ggoSpread.UpdateFlag Then
				strVal = strVal & "UPDATE" & iColSep & lngRows & iColSep
				.Col = C_Cost_Cd
				strVal = strVal & Trim(.Text) & iColSep
				            
				.Col = C_DstbFctr_Cd	
				strVal = strVal & Trim(.Text) & iColSep
	
				.Col = C_Dstb_Order	
				strVal = strVal & Trim(.Text) & iRowSep
  				lGrpCnt = lGrpCnt + 1
			ElseIf .Text = ggoSpread.DeleteFlag Then
			        strDel = strDel & "DELETE" & iColSep & lngRows & iColSep
				.Col = C_Cost_Cd
				strDel = strDel & Trim(.Text) & iRowSep
				lGrpCnt = lGrpCnt + 1
			End If
		Next	
    End With
	
	frm1.txtMaxRows.value     = lGrpCnt-1	
	frm1.txtSpread.value      = strDel & strVal

    lGrpCnt = 1
    strVal = ""
    strDel = ""
 
    With frm1.vspdData3
    	For itemRows = 1 To .MaxRows
	  		.Row = itemRows
			.Col = 0 

			Select Case .Text
				Case ggoSpread.DeleteFlag
					strDel = strDel & "D" & iColSep & itemRows & iColSep
					.Col = 1 	
					strDel = strDel & Trim(.Text) & iColSep
					.Col = 3  
					strDel = strDel & Trim(.Text) & iRowSep
						        
					lGrpCnt = lGrpCnt + 1
				Case ggoSpread.InsertFlag
					.Col = 2  
					If .text = "1" Then
						strVal = strVal & "C" & iColSep & itemRows & iColSep
						.Col = 1 
						strVal = strVal & Trim(.Text) & iColSep
						.Col =  3  
						strVal = strVal & Trim(.Text) & iRowSep
								
						lGrpCnt = lGrpCnt + 1
					End If		
			End Select
		Next
	End With
		
    frm1.txtMaxRows3.value = lGrpCnt-1	
    frm1.txtSpread3.value =   strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    DbSave = True                                                                '☜: Processing is OK

    Set gActiveElement = document.ActiveElement   
End Function

'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
    DbDelete = False                                                             '☜: Processing is NG

    DbDelete = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
Sub DbQueryOk()
    Err.Clear 
	On Error Resume Next
	
	With frm1
		SetSpreadLock "Q", 0, 1, ""
    
        lgIntFlgMode = Parent.OPMD_UMODE	
        
        Call ggoOper.LockField(Document, "I")	
        Call SetToolbar("110011110011111")	
        
        If .vspdData.MaxRows > 0 Then
            .vspdData.Row = 1
            .vspdData.Col = 1
			
			frm1.vspddata2.maxrows = 0

            Call DbQuery2(1)
       End If

       frm1.txtvercd.focus
    End With

    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================================
Sub DbSaveOk()
    Call InitVariables    															     '⊙: Initializes local global variables

    frm1.vspdData.MaxRows = 0 
	
    If DbQuery() = False Then
       Call RestoreToolBar()
       Exit Sub
    End if

    Set gActiveElement = document.ActiveElement   
End Sub

'========================================================================================================
Sub DbDeleteOk()

End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================
' Name : JumpAllocComp
' Desc : Jump to Dstb Rule By Acct
'========================================================================================================

Function JumpAllocComp()
    Dim IntRetCd, strVal
   
	ggoSpread.Source = frm1.vspdData

	If ggoSpread.SSCheckChange = True Then

		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")	

    	If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
		
    If frm1.vspdData.MaxRows = 0 Then 
		intRetCD = DisplayMsgBox("181216","x","x","x")
		Exit Function
	End If
		
	If frm1.vspdData.ActiveRow = 0 Then 
		intRetCD = DisplayMsgBox("181216","x","x","x")
		Exit Function
	End If
		
	Call PgmJump(BIZ_PGM_JUMP_ID1)	
End Function

'========================================================================================================
' Name : OpenPopup
' Desc : developer describe this line 
'========================================================================================================
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	Select Case iWhere
		Case 0
			arrParam(0) = "버전팝업"	
			arrParam(1) = "C_Dstb_Rule_by_CC"		
			arrParam(2) = strCode
			arrParam(3) = ""		
			arrParam(4) = ""	
			arrParam(5) = "버전"

			arrField(0) = "ver_cd"
    
			arrHeader(0) = "버전"	

		Case 1
			arrParam(0) = "코스트센타팝업"
			arrParam(1) = "B_Cost_Center"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""		
			arrParam(5) = "코스트센타"	

			arrField(0) = "COST_CD"	
			arrField(1) = "COST_NM"	
    
			arrHeader(0) = "코스트센타코드"
			arrHeader(1) = "코스트센타명"
			
		Case 2
			arrParam(0) = "배부요소팝업"
			arrParam(1) = "C_DSTB_FCTR"	
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""		
			arrParam(5) = "배부요소"

			arrField(0) = "DSTB_FCTR_CD"
			arrField(1) = "DSTB_FCTR_NM"	
    
			arrHeader(0) = "배부요소코드"
			arrHeader(1) = "배부요소명"
	End Select
    
    If iWhere = 0 Then
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=360px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End If

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtVerCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	
End Function

'========================================================================================================
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
				 frm1.txtVerCd.focus
				.txtVerCd.value = arrRet(0)
			Case 1
				.vspdData.Row = .vspdData.ActiveRow	
				.vspdData.Col = C_Cost_Cd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_Cost_Nm
				.vspdData.Text = arrRet(1)

				Call vspdData_Change(C_Cost_Cd, frm1.vspddata.activerow )
			Case 2
				.vspdData.Row = .vspdData.ActiveRow	
				.vspdData.Col = C_DstbFctr_Cd
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_DstbFctr_Nm
				.vspdData.Text = arrRet(1)
				Call vspdData_Change(C_DstbFctr_Cd, frm1.vspddata.activerow )
		End Select
	End With
End Function

'========================================================================================================
Function FindNumber(ByVal objSpread, ByVal intCol)
	Dim lngRows
	Dim lngPrevNum
	Dim lngNextNum

    FindNumber = 0

    lngPrevNum = 0
    lngNextNum = 0
    
    With frm1
        If objSpread.MaxRows = 0 Then
            Exit Function
        End If
        
        For lngRows = 1 To objSpread.MaxRows
            objSpread.Row = lngRows
            objSpread.Col = intCol
            lngNextNum = UniClng(objSpread.Text,0)
            
            If lngNextNum > lngPrevNum Then
                lngPrevNum = lngNextNum
            End If
        Next
    End With        
    
    FindNumber = lngPrevNum
End Function

'========================================================================================================
Function FindData()
	Dim strApNo
	Dim strItemSeq
	Dim strDtlSeq
	Dim lRows

    FindData = 0

    With frm1
        For lRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lRows
            .vspdData3.Col = 1
            strItemSeq = .vspdData3.Text
            .vspdData3.Col = 3
            strDtlSeq = .vspdData3.Text
            
            .vspdData.Row = frm1.vspdData.ActiveRow
            .vspdData2.Row = frm1.vspdData2.ActiveRow
            
            .vspdData.Col = C_Cost_Cd

            If strItemSeq = .vspdData.Text Then
                .vspdData2.Col = C_RecvCost_Cd
                If strDtlSeq = .vspdData2.Text Then
                    FindData = lRows
                    Exit Function
                End If
            End If    
        Next
    End With        
End Function

'========================================================================================================
Function CopyFromData(ByVal strItemSeq)
	Dim lngRows , i
	Dim boolExist
	Dim iCols

    boolExist = False

    frm1.vspdData2.maxrows = 0
    CopyFromData = boolExist
    
    With frm1
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = 1                

            If Trim(strItemSeq) = Trim(.vspdData3.Text) Then
                boolExist = True
                Exit For
            End If    
        Next
        
         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            .vspdData2.Redraw = False

            While lngRows <= .vspdData3.MaxRows
				.vspdData3.Row = lngRows

                .vspdData3.Col = 1

                If strItemSeq <> .vspdData3.Text Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData2.MaxRows = .vspdData2.MaxRows + 1
                    .vspdData2.Row = .vspdData2.MaxRows
                    .vspdData2.Col = 0
                    .vspdData3.Col = 0
                    .vspdData2.Text = .vspdData3.Text

                    .vspdData2.Col = C_CheckFlag
                    .vspdData3.Col = 2
                    .vspdData2.Text = .vspdData3.Text
                  
                    .vspdData2.Col = C_RecvCost_Cd
                    .vspdData3.Col = 3
                    .vspdData2.Text = .vspdData3.Text

                    .vspdData2.Col = C_RecvCost_Nm
                    .vspdData3.Col = 4
                    .vspdData2.Text = .vspdData3.Text
                End If   
                
                lngRows = lngRows + 1
            Wend
            
            ggoSpread.Source = frm1.vspdData2

			SetSpread2Color	

            frm1.vspdData.Row = lgCurrRow
            frm1.vspdData.Col = frm1.vspdData.MaxCols
            ggoSpread.Source = frm1.vspdData
            
            frm1.vspdData2.Redraw = True
        End If
    End With        
    
    CopyFromData = boolExist
End Function

'========================================================================================================
Function CancelHSheet(ByVal strItemSeq)
	Dim lngRows
 
    CancelHSheet = False
    lngRows = 1

    With frm1
		For lngRows = 1 To .vspdData3.MaxRows
		    .vspdData3.Row = lngRows
		    .vspdData3.Col = 1                

		    If strItemSeq = .vspdData3.Text Then
		        Exit For
            End If    
        Next
        
        While lngRows <= .vspdData3.MaxRows
			.vspdData3.Row = lngRows
            .vspdData3.Col = 1
            
            If Trim(strItemSeq) = Trim(.vspdData3.Text) Then
                .vspdData3.Col = 0
                
                If .vspdData3.Text = ggoSpread.InsertFlag or  .vspdData3.Text = ggoSpread.DeleteFlag Then
					.vspdData3.Text = ""						
					.vspdData3.Col = 2

					If .vspdData3.text = "0" Then
						.vspdData3.text = "1"
					Else
						.vspdData3.text = "0"
					End If	
				End If

				lngRows = lngRows + 1   
			Else
				lngRows = .vspdData3.MaxRows + 1
			End If
        Wend
    End With

    CancelHSheet = True
End Function    

'========================================================================================================
Sub CopyToHSheet(ByVal Row)
	Dim lRow
	Dim iCols

	With frm1 
	    lRow = FindData

	    If lRow > 0 Then
            .vspdData3.Row = lRow
            .vspdData2.Row = Row
            .vspdData3.Col = 0
            .vspdData2.Col = 0
            .vspdData3.Text = .vspdData2.Text
        
			.vspdData2.Col = C_CheckFlag
			.vspdData3.Col = 2
            .vspdData3.Text = .vspdData2.value

			.vspdData2.Col = C_RecvCost_Cd
			.vspdData3.Col = 3
            .vspdData3.Text = .vspdData2.value
			
			.vspdData2.Col = C_RecvCost_Nm
			.vspdData3.Col = 4
            .vspdData3.Text = .vspdData2.value
        End If
	End With

	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = 0

	If frm1.vspdData.Text <> ggoSpread.InsertFlag And frm1.vspdData.Text <> ggoSpread.DeleteFlag Then
   	    frm1.vspdData.Text = ggoSpread.UpdateFlag
	End If

End Sub

'========================================================================================================
Function DeleteHSheet(ByVal strItemSeq)
	Dim boolExist
	Dim lngRows
 
    DeleteHSheet = False
    boolExist = False

    frm1.vspdData2.MaxRows = 0

    With frm1
        For lngRows = 1 To .vspdData3.MaxRows
            .vspdData3.Row = lngRows
            .vspdData3.Col = 1                

            If strItemSeq = .vspdData3.Text Then
                boolExist = True
                Exit For
            End If    
        Next

         .vspdData3.Row = lngRows
        
        If boolExist = True Then
            While lngRows <= .vspdData3.MaxRows
                .vspdData3.Row = lngRows
                .vspdData3.Col = 1
                
                If strItemSeq <> .vspdData3.Text Then
                    lngRows = .vspdData3.MaxRows + 1
                Else
                    .vspdData3.Action = 5
                    .vspdData3.MaxRows = .vspdData3.MaxRows - 1
                End If   
            Wend
            
            ggoSpread.Source = frm1.vspdData2
            
            frm1.vspdData.Row = lgCurrRow
            frm1.vspdData.Col = frm1.vspdData.MaxCols
            ggoSpread.Source = frm1.vspdData
            
            frm1.vspdData2.Redraw = True
        End If
    End With
        
    DeleteHSheet = True
End Function    

'========================================================================================================
Function SortHSheet()
    With frm1
        .vspdData3.BlockMode = True
        .vspdData3.Col = 0
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 1
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.SortBy = 0 
        
        .vspdData3.SortKey(1) = 1
        .vspdData3.SortKey(2) = 2
        
        .vspdData3.SortKeyOrder(1) = 1 
        .vspdData3.SortKeyOrder(2) = 1
        
        .vspdData3.Col = 1
        .vspdData3.Col2 = .vspdData3.MaxCols
        .vspdData3.Row = 0
        .vspdData3.Row2 = .vspdData3.MaxRows
        .vspdData3.Action = 25 
        .vspdData3.BlockMode = False
    End With        
End Function

'========================================================================================================
Sub ShowHidden()
	Dim strHidden
	Dim lngRows
	Dim lngCols
    
    With frm1.vspdData3
        For lngRows = 1 To .MaxRows
            .Row = lngRows
            
            .Col = 1  
            strHidden = strHidden & Parent.gRowSep & .Text
            .Col = 2
            strHidden = strHidden & Parent.gRowSep & .Text
            .Col = 3
            strHidden = strHidden & Parent.gRowSep & .Text
            .Col = 4
            strHidden = strHidden & Parent.gRowSep & .Text
		
            .Col = 5  
            strHidden = strHidden & Parent.gRowSep & .Text		
   		
            strHidden = strHidden & Parent.gRowSep
        Next
    End With        
End Sub

'========================================================================================================
Sub SetSpreadFG( pobjSpread , ByVal pMaxRows )
    Dim lngRows 
    
    For lngRows = 1 To pMaxRows
        pobjSpread.Col = 0
        pobjSpread.Row = lngRows
        pobjSpread.Text = ""
    Next
End Sub

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
    Dim i

   	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else 
		Call SetPopupMenuItemInf("1101111111")
	End If	

    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.maxrows = 0 Then Exit Sub
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
    End If
    
    If Row = 0 Then
		Exit Sub
	End If
	
    ggoSpread.Source = frm1.vspdData
	frm1.vspddata.row = Row

	frm1.vspdData2.maxrows = 0

  	frm1.vspdData.Col = C_Cost_Cd
	
    If Len(Trim(frm1.vspdData.Text)) > 0 Then
		frm1.vspddata.Col = 0

		If frm1.vspddata.Text = ggoSpread.DeleteFlag Then
			Exit Sub
        End If
 
      	If frm1.vspddata.Text = ggoSpread.InsertFlag Then           
		 	If DbQuery3(Row) = False Then
		 		Exit Sub
		 	End If	
        Else           
		 	If DbQuery2(Row) = False Then
		 		Exit Sub	
		 	End If
    	End If	
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
'   Event Name : vspdData2_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
   
End Sub

'========================================================================================================
'   Event Name : vspdData2_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
	
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(Col, Row, ButtonDown)
	ggoSpread.Source = frm1.vspdData

	With frm1
		If Row > 0 Then
			Select Case Col
				Case  C_Cost_PB 
					.vspdData.Col = C_Cost_Cd
					.vspdData.Row = Row

					Call OpenPopUp(.vspdData.Text, 1 )
	 			Case C_DstbFctr_PB
					.vspdData.Col = C_DstbFctr_cd
					.vspdData.Row = Row

					Call OpenPopUp(.vspdData.Text, 2 )

			End Select	
		End If

		Call SetActiveCell(.vspdData,Col-1,.vspdData.ActiveRow ,"M","X","X")   	
	End With
End Sub

Sub vspdData2_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(Col , Row)
	Dim IntRetCd

    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)
 
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
    
    frm1.vspdData.Row = Row
    
    Select Case Col
		Case   C_Cost_Cd
			frm1.vspdData.Col = 0
			If  frm1.vspdData.Text = ggoSpread.InsertFlag Then
				frm1.vspdData.Col = C_Cost_Cd
				frm1.hCostCd.value = frm1.vspdData.Text

		        If Len(Trim(frm1.vspdData.Text)) > 0 Then
					frm1.vspdData.Row = Row
					frm1.vspdData.Col = C_Cost_Cd	
					DeleteHsheet frm1.vspdData.Text
		    
		            If DbQuery3 (Row) = False Then
						Exit Sub
					End If
		        End If  
			End If 
    End Select
End Sub

'========================================================================================================
Sub vspdData2_Change(ByVal Col, ByVal Row)
    ggoSpread.Source = frm1.vspdData2
	ggoSpread.UpdateRow Row

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = 0
	
	Select Case Col
		Case   C_CheckFlag
			If  frm1.vspdData2.Text <> ggoSpread.InsertFlag and frm1.vspdData2.Text <> ggoSpread.DeleteFlag Then
			    frm1.vspdData2.Col = C_CheckFlag
			        			
			    If frm1.vspdData2.value = "0" Then
					frm1.vspdData2.Row = Row
					frm1.vspdData2.Col = 0	
					frm1.vspdData2.text = ggoSpread.DeleteFlag
			    Else
					frm1.vspdData2.Row = Row
					frm1.vspdData2.Col = 0	
					frm1.vspdData2.text = ggoSpread.InsertFlag 
			    End If  
			Elseif frm1.vspdData2.Text = ggoSpread.DeleteFlag Then
				frm1.vspdData2.Col = C_CheckFlag

			    If frm1.vspdData2.value = "1" Then
					frm1.vspdData2.Row = Row
					frm1.vspdData2.Col = 0	
					frm1.vspdData2.text = frm1.vspdData2.Row
			    End If 
			Elseif frm1.vspdData2.Text = ggoSpread.InsertFlag Then
				frm1.vspdData2.Col = C_CheckFlag

			    If frm1.vspdData2.value = "0" Then
				    frm1.vspdData2.Row = Row
				    frm1.vspdData2.Col = 0	
				    frm1.vspdData2.text = frm1.vspdData2.Row
			    End If  
	    	End If  
	End Select

	CopyToHSheet Row
End Sub	

'==========================================================================================
'   Event Name : vspdData2_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData2_Click(ByVal Col, ByVal Row)
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("0001111111")
	Else 
		Call SetPopupMenuItemInf("0001111111")
	End If	

	gMouseClickStatus = "SP2C"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData2
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
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData2

    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
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
Sub vspdData2_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If
End Sub    

'========================================================================================================
Sub vspdData_onfocus()
    If lgIntFlgMode <> Parent.OPMD_UMODE Then    

    Else  
		Call SetToolbar("1100111100111111")	
    End If    
End Sub


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<HTML>
<HEAD>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>C/C별배부규칙등록</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">버전</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtVerCd" SIZE=10 MAXLENGTH=3 tag="12XXXU" ALT="버전"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnVerCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtVerCd.Value, 0)"></TD>
									<TD CLASS="TD6"></TD>
									<TD CLASS="TD6"></TD>
								</TR>               
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR >
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD WIDTH="65%" HEIGHT="100%" >
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
								<TD WIDTH="35%" HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>

							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>				
				<TR >
					<TD WIDTH=100% HEIGHT=100%>
						<FIELDSET CLASS="CLSFLD">					
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">신규버전</TD>
									<TD CLASS="TD6"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtNewVerCd" SIZE=10 MAXLENGTH=3 tag="22XXXU" ALT="신규버전">&nbsp;<BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeCopy()" Flag=1>복사실행</BUTTON></TD>
									<TD CLASS="TD6"></TD>
									<TD CLASS="TD6" ALIGN=RIGHT><a href="VBSCRIPT:JumpAllocComp()" ONCLICK="VBSCRIPT:CookiePage 1">계정별배부규칙등록</A></TD>
								</TR>
							</TABLE>
						</FIELDSET>						
					</TD>
				</TR>        				
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>

	<TR>
		<TD WIDTH="100%" HEIGHT=<%=BIZSIZE%>>
			<IFRAME NAME="MyBizASP" SRC="" WIDTH="100%" HEIGHT=<%=BIZSIZE%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSpread tag="24" TABINDEX= "-1"></TEXTAREA>
<TEXTAREA class=hidden name=txtSpread3 tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=hidden NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=hidden NAME="htxtVerCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=hidden NAME="hCostCd" tag="24" TABINDEX= "-1">
<INPUT TYPE=hidden NAME="txtCommandMode" tag="24" TABINDEX= "-1">

<INPUT TYPE=hidden NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX= "-1">

<INPUT  NAME="txtMaxRows3" tag="24">

<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> HEIGHT=0 name=vspdData3 width="100%" tag="2" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

