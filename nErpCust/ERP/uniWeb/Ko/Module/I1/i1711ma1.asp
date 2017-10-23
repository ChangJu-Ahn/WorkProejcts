<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Batch Posting 작업 
'*  3. Program ID           : i1711ma1
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/05/08
'*  8. Modified date(Last)  : 2003/06/03
'*  9. Modifier (First)     : lee hae ryong
'* 10. Modifier (Last)      : Lee Seung Wook 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">

Option Explicit	

Const BIZ_PGM_ID = "i1711mb2.asp"
Const BIZ_LOOKUP_ID = "i1711mb1.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgStrPrevKey2
Dim lgCheckall
DIm SetComboList

Dim C_Check
Dim C_ItemDocumentNo
Dim C_TrnsType
Dim C_MovType 
Dim C_DocumentDt
Dim C_PosDt
Dim IsOpenPop          

 '#########################################################################################################
'					2. Function부 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
	lgIntFlgMode		= Parent.OPMD_CMODE 
	lgBlnFlgChgValue	= False     	 
	lgIntGrpCount		= 0                
	
	lgStrPrevKey		= "" 
	lgStrPrevKey2		= ""               
	lgLngCurRows		= 0                 
	lgCheckall			= 0
End Sub

 '==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	Dim StartDate
	frm1.btnRun.Disabled = True
	frm1.txtBizCd.focus 

	StartDate = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtDocumentFrDt.Text	= UNIDateAdd("m", -1, StartDate, Parent.gDateFormat)
	frm1.txtDocumentToDt.Text	= StartDate
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>

End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
'========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20030410", , Parent.gAllowDragDropSpread
 		
		.ReDraw = false
		
 		.MaxCols = C_PosDt + 1
		.MaxRows = 0

 		Call GetSpreadColumnPos("A")		

		ggoSpread.SSSetCheck	C_Check,			"",						4,,,1
		ggoSpread.SSSetEdit		C_ItemDocumentNo,	"수불번호",			18
		ggoSpread.SSSetEdit		C_TrnsType,			"수불구분",			14,2
		ggoSpread.SSSetEdit		C_MovType,			"수불유형",			30		
		ggoSpread.SSSetDate		C_DocumentDt,		"수불발생일",		12,2,Parent.gDateFormat
		ggoSpread.SSSetDate		C_PosDt,			"회계전표발생일",	12,2,Parent.gDateFormat
		
  		Call ggoSpread.SSSetColHidden(C_PosDt, .MaxCols, True)
	    Call SetSpreadLock
		
		.ReDraw = True
   		ggoSpread.SSSetSplit2(2)  
 
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
	'ggoSpread.SpreadLock -1, -1
	ggoSpread.SpreadLockWithOddEvenRowColor()
	ggoSpread.SpreadUnLock C_Check, -1, C_Check
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD =" & FilterVar("I0002", "''", "S") & " ORDER BY MINOR_NM ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboTrnsType,lgF0,lgF1,Chr(11))
	SetComboList = lgF0 & Chr(12) & lgF1
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_Check				= 1
	C_ItemDocumentNo	= 2							
	C_TrnsType			= 3
	C_MovType			= 4
	C_DocumentDt		= 5
	C_PosDt				= 6
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
 		C_Check				= iCurColumnPos(1)
		C_ItemDocumentNo	= iCurColumnPos(2)
		C_TrnsType			= iCurColumnPos(3)
		C_MovType			= iCurColumnPos(4)
		C_DocumentDt		= iCurColumnPos(5)
		C_PosDt				= iCurColumnPos(6)
 	
 	End Select
 
End Sub

 '========================================== 2.4.2 Open???()  =============================================
Function OpenbizareaInfo()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장POPUP"
	arrParam(1) = "B_BIZ_AREA"		
	arrParam(2) = Trim(frm1.txtBizCd.value)
	arrParam(3) = ""						
	arrParam(4) = ""						
	arrParam(5) = "사업장"			
	
    arrField(0) = "BIZ_AREA_CD"				
    arrField(1) = "BIZ_AREA_NM"				
    
    arrHeader(0) = "사업장"				
    arrHeader(1) = "사업장명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBizCd.focus
	    Exit Function
	Else
		Call SetbizareaInfo(arrRet)
	End If	

End Function

 '------------------------------------------  OpenMovType()  -------------------------------------------------
Function OpenMovType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtMovType.ClassName)= UCase(Parent.UCN_PROTECTED) Then Exit Function	

	IsOpenPop = True

	arrParam(0) = "수불유형 팝업"					
	arrParam(1) = "I_MOVETYPE_CONFIGURATION A, B_MINOR B"
	arrParam(2) = Trim(frm1.txtMovType.Value)			
	arrParam(3) = ""                       				
	arrParam(4) = "A.MOV_TYPE = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("I0001", "''", "S") & ""
	arrParam(5) = "수불유형"
	
	arrField(0) = "A.MOV_TYPE"	
	arrField(1) = "B.MINOR_NM"	
	
	arrHeader(0) = "수불유형"		
	arrHeader(1) = "수불유형명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtMovType.focus
		Exit Function
	Else
		Call SetMovType(arrRet)
	End If
		
End Function

 '------------------------------------------  OpenMoveDtlRef()  -------------------------------------------------
Function OpenMoveDtlRef()
	Dim iCalledAspName
	Dim IntRetCD

	Dim arrRet
	Dim Param1  
	Dim Param2  
	Dim Param3  
	Dim Param4	
	Dim Param5	
	Dim Param6  
	Dim Param7  
	If IsOpenPop = True Then Exit Function
	
	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("800167","X","X","X")
		frm1.txtBizCd.focus
		Exit Function
	End If

	Param1 = Trim(frm1.txtBizCd.value)
	Param2 = Trim(frm1.txtBizNm.value)
	
	if Param1 = "" then
		Call DisplayMsgBox("169803","X", "X", "X")
		frm1.txtBizCd.focus
		Exit Function
	End If
	
	ggoSpread.Source = frm1.vspdData    
    Param5 = ""
	With frm1.vspdData	    
		If .MaxRows = 0 Then
		    Call DisplayMsgBox("169804","X", "X", "X")
			Exit Function
		else
		   .Col = C_ItemDocumentNo  : .Row = .ActiveRow : Param3 = Trim(.Text )
		   .Col = C_DocumentDt      : .Row = .ActiveRow : Param4 = Trim(.Text )
		   .Col = C_PosDt           : .Row = .ActiveRow : Param5 = Trim(.Text )
		   .Col = C_TrnsType        : .Row = .ActiveRow : Param6 = Trim(.Text )
		   .Col = C_MovType         : .Row = .ActiveRow : Param7 = Trim(.Text )
		End If	
    End With
    	
    if Param3 = "" then
       Call DisplayMsgBox("169804","X", "X", "X")
    	Exit Function
    End If
	
	IsOpenPop = True

	iCalledAspName = AskPRAspName("I1711RA1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"I1711RA1","x")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, Param1,Param2,Param3,Param4,Param5,Param6,Param7), _
		 "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")		    
    	
	IsOpenPop = False

	If arrRet(0) = "" Then
		Call SetActiveCell(frm1.vspdData,C_ItemDocumentNo,frm1.vspdData.ActiveRow,"M","X","X")
		Exit Function
	End If	
	
End Function

'------------------------------------------  SetbizareaInfo()  --------------------------------------------------
Function SetbizareaInfo(byRef arrRet)
	frm1.txtBizCd.Value    = arrRet(0)		
	frm1.txtBizNm.Value    = arrRet(1)		
	frm1.txtBizCd.focus
End Function

 '------------------------------------------  SetMovType()  --------------------------------------------------
Function SetMovType(byRef arrRet)
	frm1.txtMovType.Value   = arrRet(0)
	frm1.txtMovTypeNm.Value = arrRet(1)
	frm1.txtMovType.focus
End Function

'========================================================================================
' Function Name : Checkall()
'========================================================================================
Function Checkall()
	
 Dim IRowCount 
 Dim IClnCount
 ggoSpread.Source = frm1.vspdData
 With frm1.vspdData    
  IF lgCheckall = 0 Then 
   for IClnCount = 0 to C_Check
   	for IRowCount = 1 to .MaxRows
   	     if IClnCount <> 0 then   	     	 
        	 .Row = IRowCount 
        	 .Col = IClnCount	 
 	         .text = 1     
 	     Else
 	         .Row = IRowCount
 	         .Col = IClnCount
 	         .Text =ggoSpread.UpdateFlag
 	     End if
	next    
   next
   lgCheckall = 1
   lgBlnFlgChgValue = True
  Else
   
   for IClnCount = 0 to C_Check
   	for IRowCount = 1 to .MaxRows
   	     if IClnCount <> 0 then   	     	 
        	 .Row = IRowCount 
        	 .Col = IClnCount	 
 	         .text = 0     
 	     End if
	next    
   next
   lgCheckall = 0
   lgBlnFlgChgValue = False
  End If

 End With
 
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    
    Call LoadInfTB19029 
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    Call InitSpreadSheet                 
    Call InitVariables                   
    
    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("11000000000011")
		
End Sub

'=======================================================================================================
'   Event Name : txtDocumentFrDt_DblClick(Button)
'=======================================================================================================
Sub txtDocumentFrDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentFrDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentFrDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentFrDt_KeyPress()
'=======================================================================================================
Sub txtDocumentFrDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentToDt_DblClick(Button)
'=======================================================================================================
Sub txtDocumentToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentToDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentToDt_KeyPress()
'=======================================================================================================
Sub txtDocumentToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 And Col = C_Check Then
			.Col = Col
			.Row = Row									
			IF .Text = "1" Then
				.Col = 0
				.Text = ggoSpread.UpdateFlag
			Elseif .Text = "0" Then
				.Col = 0
				.Text = ""
			End if  							
		End If	
	End With
End Sub


'=========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   
	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub
	 '----------  Coding part  -------------------------------------------------------------   
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and lgStrPrevKey <> "" and lgStrPrevKey2 <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
		Call DisableToolBar(Parent.TBC_QUERY)
		If DbQuery = False Then
			Call RestoreToolBar()
			Exit Sub
		End If
	End if 
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	
 	Call SetPopupMenuItemInf("0000111111")         
 	
 	gMouseClickStatus = "SPC"   
    
 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData 
 		If lgSortKey = 1 Then
 			ggoSpread.SSSort Col
 			lgSortKey = 2
 		Else
 			ggoSpread.SSSort Col, lgSortKey	
 			lgSortKey = 1
 		End If
 		Exit Sub
 	End If
 	
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
    
 	If Row <= 0 Then Exit Sub
  	If frm1.vspdData.MaxRows = 0 Then Exit Sub
End Sub
 
'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 
 
'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   	
   	If NewCol = C_Check or Col = C_Check Then
		Cancel = True
		Exit Sub
	End If
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
End Sub 
 
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 
 
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
     ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
End Sub 

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
 
    Dim IntRetCD 
    
    Err.Clear

    FncQuery = False 
    
    If GetSetupMod(Parent.gSetupMod, "a") <> "Y" then
       Call DisplayMsgBox("169934","X", "X", "X")
       Exit Function
	End if	

    '-----------------------
    'Check previous data area
    '-----------------------
    If ggoSpread.SSCheckChange = True Then		
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not ChkField(Document, "1") Then	Exit Function
    '-----------------------
    'Erase contents area
    '-----------------------
 	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    Call InitVariables

    If ValidDateCheck(frm1.txtDocumentFrDt, frm1.txtDocumentToDt) = False Then Exit Function

	If 	CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA ", " BIZ_AREA_CD = " & Trim(FilterVar(frm1.txtBizCd.value," ","S")), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false Then
	
		Call DisplayMsgBox("124200","X","X","X")
		frm1.txtBizNm.value = ""
		frm1.txtBizCd.Focus		
		Exit function
    
    End If
    lgF0 = Split(lgF0,Chr(11))
	frm1.txtBizNm.value = lgF0(0)

	If Trim(frm1.txtMovType.Value) <> "" Then
		If 	CommonQueryRs(" B.MINOR_NM "," I_MOVETYPE_CONFIGURATION A, B_MINOR B ", _
		    " A.MOV_TYPE = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND A.MOV_TYPE = " & Trim(FilterVar(frm1.txtMovType.value, "''", "S")), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false Then
	
			Call DisplayMsgBox("169948","X","X","X")
			frm1.txtMovTypeNm.value = ""
			frm1.txtMovType.Focus		
			Exit function
    
		End If
		lgF0 = Split(lgF0,Chr(11))
		frm1.txtMovTypeNm.value = lgF0(0)
	Else
		frm1.txtMovTypeNm.value = ""
	End If
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then	Exit Function
       
    FncQuery = True	
 
End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
	Dim IntRetCD 
	
	FncSave = False
	
	Err.Clear
	'-----------------------
	'Precheck area
	'-----------------------
    ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = False Then
	   IntRetCD = DisplayMsgBox("169801","X", "X", "X")
 	   Exit Function
	End If
	
	'-----------------------
	'Save function call area
	'-----------------------
	If DbSave = False Then	Exit Function
	
	FncSave = True 

End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If
    FncExit = True

End Function

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then  Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : RemovedivTextArea
'========================================================================================
Function RemovedivTextArea()
	Dim i
	For i = 1 To divTextArea.children.length
		divTextArea.removeChild(divTextArea.children(0))
	Next
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
	
	Call LayerShowHide(1) 
    DbQuery = False
    
    Err.Clear

    Dim strVal
 
    With frm1
		strVal = BIZ_LOOKUP_ID &	"?txtBiZCd="			& Trim(.txtBizCd.value)			& _	
									"&txtDocumentFrDt="		& Trim(.txtDocumentFrDt.Text)	& _
									"&txtDocumentToDt="		& Trim(.txtDocumentToDt.Text)	& _
									"&cboTrnsType="			& Trim(.cboTrnsType.value)		& _
									"&txtMovType="			& Trim(.txtMovType.value)		& _	
									"&lgStrPrevKey="		& lgStrPrevKey					& _
									"&lgStrPrevKey2="		& lgStrPrevKey2					& _
									"&SetComboList="		& SetComboList					& _
									"&txtMaxRows="			& .vspdData.MaxRows

    Call RunMyBizASP(MyBizASP, strVal)
       
    End With
    
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()
		
	Call SetToolbar("11001000000111")	
    
    lgIntFlgMode = Parent.OPMD_UMODE
   	frm1.btnRun.Disabled = False    
	frm1.txtBizCd.focus
		
End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 
    
    Dim lRow        
    Dim strVal, strYear, strMonth, strDay
	Dim iRowSep, iColSep
	
	Dim strCUTotalvalLen
	Dim objTEXTAREA
	Dim iTmpCUBuffer
	Dim iTmpCUBufferCount
	Dim iTmpCUBufferMaxCount
	
	iRowSep = Parent.gRowSep
	iColSep = Parent.gColSep

    Err.Clear		
    DbSave = False

    On Error Resume Next
	
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
	iTmpCUBufferCount = -1
	strCUTotalvalLen = 0
	
    Call LayerShowHide(1)
        
	frm1.txtMode.value = Parent.UID_M0002
		
    '-----------------------
    'Data manipulate area
    '-----------------------
	With frm1.vspdData
    
		For lRow = 1 To .MaxRows
			strVal = ""
			.Row = lRow
			.Col = 0
			
			Select Case .Text
				Case ggoSpread.UpdateFlag
					
					.Col = C_ItemDocumentNo
		 	  		strVal = Trim(.Text) & iColSep
		          	
		          	.Col = C_DocumentDt
					Call ExtractDateFrom(.Text,Parent.gDateFormat,Parent.gComDateType,strYear,strMonth,strDay)
    		        strVal = strVal & strYear & iColSep
    		        strVal = strVal & lRow & iRowSep
    				        
					If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then
								                            
						Set objTEXTAREA = document.createElement("TEXTAREA")
						objTEXTAREA.name = "txtCUSpread"
						objTEXTAREA.value = Join(iTmpCUBuffer,"")
						divTextArea.appendChild(objTEXTAREA)     
									 
						iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT   
						ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
						iTmpCUBufferCount = -1
						strCUTotalvalLen  = 0
											
					End If
								       
					iTmpCUBufferCount = iTmpCUBufferCount + 1
								      
					If iTmpCUBufferCount > iTmpCUBufferMaxCount Then    
						iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
						ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
					End If   
											
					iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
					strCUTotalvalLen = strCUTotalvalLen + Len(strVal)

			End Select
		
		Next
	
	End With
	
	If iTmpCUBufferCount > -1 Then 
		Set objTEXTAREA = document.createElement("TEXTAREA")
		objTEXTAREA.name   = "txtCUSpread"
		objTEXTAREA.value = Join(iTmpCUBuffer,"")
		divTextArea.appendChild(objTEXTAREA)     
	else
		Call DisplayMsgBox("169801","X", "X", "X")
		Exit Function
	End If  
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	
    DbSave = True
    
End Function

Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()
   
	Call InitVariables	
   	ggoSpread.source = frm1.vspddata
    ggoSpread.ClearSpreadData
    Call FncQuery()
    Call SetToolbar("11000000000011")
    
End Function

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
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Batch Posting작업</font></TD>
								<TD background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></TD>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH=* align=right><A href="vbscript:OpenMoveDtlRef()">수불상세정보</A></TD>					
					<TD WIDTH=10>
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
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 tag="12XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenbizareaInfo()">&nbsp;
								                       <INPUT TYPE=TEXT NAME="txtBizNm" SIZE=20 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP>수불일자</TD>
								<TD CLASS="TD6" NOWRAP>
								      <script language =javascript src='./js/i1711ma1_fpDateTime1_txtDocumentFrDt.js'></script>
							          &nbsp;~&nbsp;
							          <script language =javascript src='./js/i1711ma1_fpDateTime2_txtDocumentToDt.js'></script>
							    </TD>      
							</TR>
							<TR>
					           	<TD CLASS="TD5" NOWRAP>수불구분</td>
								<TD CLASS="TD6" NOWRAP>
								<SELECT Name="cboTrnsType" ALT="수불구분" STYLE="WIDTH: 100px" tag="11"><OPTION Value=""></OPTION></SELECT>
								</TD>
								<TD CLASS="TD5" NOWRAP>수불유형</td>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT Name="txtMovType" SIZE="5" MAXLENGTH="3"  ALT="수불유형" tag="11XXXU"><IMG align=top height=20 name=btnMovType onclick="vbscript:OpenMovType()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtMovTypeNm" size="20" tag="14">
								</TD>
							</TR> 	
						</TABLE>
					</FIELDSET>
				</TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
	
			<TR HEIGHT=*>
				<TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
					<TABLE <%=LR_SPACE_TYPE_20%>>						
						<TR>
							<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
								<script language =javascript src='./js/i1711ma1_I429171580_vspdData.js'></script>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSMBTN" ONCLICK="vbscript:Checkall()">전체 선택/취소</BUTTON></TD>		
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

