<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : Batch Posting 취소 (Monthly)
'*  3. Program ID           : i1722ma1
'*  4. Program Name         : 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/05/08
'*  8. Modified date(Last)  : 2004/11/22
'*  9. Modifier (First)     : lee hae ryong
'* 10. Modifier (Last)      : Jeong Yong Kyun
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

Const BIZ_PGM_QRY_ID  = "i1722mb1.asp"
Const BIZ_PGM_QRY2_ID = "i1722mb2.asp"
Const BIZ_PGM_ID      = "i1722mb3.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgCheckall
DIm SetComboList

'****************** spread1 **************************************************************************
Dim C_Check
Dim C_MovType
Dim C_MovTypeNm
Dim C_TrnsType
Dim C_TrnsTypeNm
Dim C_GL_NO
Dim C_COST_MVMT
Dim C_COST_MVMT_NM
Dim C_REF_NO

'****************** spread2 **************************************************************************
Dim C_ItemDocumentNo 
Dim C_DocumentDt
Dim C_PosDt
Dim C_ParentRowNo
Dim C_Flag

Dim IsOpenPop         

'***************************** 2004-10-01 추가분 *****************************************************
Dim lgSpdHdrClicked

 '#########################################################################################################
'					2. Function부 
'######################################################################################################### 
 '==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
	lgIntFlgMode		= Parent.OPMD_CMODE
	lgBlnFlgChgValue	= False     	 
	lgIntGrpCount		= 0                
	
	lgStrPrevKey		= "" 
	lgLngCurRows		= 0                 
	lgCheckall			= 0
End Sub

 '==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	Dim StartDate
	frm1.btnRun.Disabled = True
	frm1.txtBizCd.focus 
	StartDate = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
	frm1.txtDocumentDt.Text	= StartDate
	Call ggoOper.FormatDate(frm1.txtDocumentDt, parent.gDateFormat, 2)
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
		ggoSpread.Spreadinit "V20041001", , Parent.gAllowDragDropSpread

		.ReDraw = False
 		.MaxCols = C_REF_NO + 1
		.MaxRows = 0

 		Call GetSpreadColumnPos("A")		

		ggoSpread.SSSetCheck	C_Check,			""                ,   	4,,,1
		ggoSpread.SSSetEdit		C_MovType,			"수불유형"    ,		10
		ggoSpread.SSSetEdit		C_MovTypeNm,		"수불유형명"  ,		20
		ggoSpread.SSSetEdit		C_TrnsType,			"수불유형명"  ,		20
		ggoSpread.SSSetEdit		C_TrnsTypeNm,		"수불구분명"  ,		10		
'		ggoSpread.SSSetEdit		C_TEMP_GL_NO,		"결의전표번호",		14,2
		ggoSpread.SSSetEdit		C_GL_NO,		    "전표번호",		    14,2	
		ggoSpread.SSSetEdit		C_COST_MVMT,		"수불경로"    ,		10,2	
		ggoSpread.SSSetEdit		C_COST_MVMT_NM,		"수불경로"    ,		10,2	
		ggoSpread.SSSetEdit		C_REF_NO,			"참조번호"    ,		10,2			

		Call ggoSpread.MakePairsColumn(C_MovType, C_MovTypeNm)
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_REF_NO, C_REF_NO, True)	
		Call ggoSpread.SSSetColHidden(C_TrnsType, C_TrnsType, True)
		Call ggoSpread.SSSetColHidden(C_COST_MVMT, C_COST_MVMT, True)				

	    Call SetSpreadLock
		.ReDraw = True
   		ggoSpread.SSSetSplit2(2)  
    End With
End Sub

'=============================================== 2.2.3 InitSpreadSheet2() ========================================
' Function Name : InitSpreadSheet2
'========================================================================================
Sub InitSpreadSheet2()

	Call InitSpreadPosVariables2()
	
	With frm1.vspdData2
	
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.Spreadinit "V20041001", , Parent.gAllowDragDropSpread
 		
		.ReDraw = false
		
 		.MaxCols = C_Flag + 1
		.MaxRows = 0

 		Call GetSpreadColumnPos("B")		

		ggoSpread.SSSetEdit		C_ItemDocumentNo,	"수불번호",			18
		ggoSpread.SSSetDate		C_DocumentDt,		"수불발생일",		12,2,Parent.gDateFormat
		ggoSpread.SSSetDate		C_PosDt,		     C_PosDt,		12,2,Parent.gDateFormat
		ggoSpread.SSSetEdit		C_ParentRowNo ,		"C_ParentRowNo", 5
		ggoSpread.SSSetEdit		C_Flag,				"C_Flag", 5
		
	    Call ggoSpread.SSSetColHidden(C_PosDt,C_PosDt, True)
	    Call ggoSpread.SSSetColHidden(C_ParentRowNo,C_ParentRowNo, True)
 		Call ggoSpread.SSSetColHidden(C_Flag, C_Flag, True)
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	    
	    Call SetSpreadLock2
		
		.ReDraw = True
 
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
	ggoSpread.SpreadUnLock C_Check, -1, C_Check
End Sub

'================================== 2.2.4 SetSpreadLock2() ==================================================
Sub SetSpreadLock2()
	'ggoSpread.SpreadLock -1, -1
	ggoSpread.SpreadLockWithOddEvenRowColor()
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
	C_MovType			= 2							
	C_MovTypeNm			= 3
	C_TrnsType          = 4	
	C_TrnsTypeNm        = 5
'	C_TEMP_GL_NO		= 6
	C_GL_NO				= 6
	C_COST_MVMT         = 7
	C_COST_MVMT_NM      = 8
	C_REF_NO			= 9
End Sub

'==========================================  2.2.7 InitSpreadPosVariables2()  =============================
Sub InitSpreadPosVariables2()
	C_ItemDocumentNo	= 1
	C_DocumentDt		= 2
	C_PosDt				= 3							
	C_ParentRowNo		= 4
	C_Flag				= 5
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 		Case "A"
 			ggoSpread.Source = frm1.vspdData 
 			
 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
 
 			C_Check				= iCurColumnPos(1)
			C_MovType			= iCurColumnPos(2)
			C_MovTypeNm			= iCurColumnPos(3)
			C_TrnsType          = iCurColumnPos(4)			
			C_TrnsTypeNm        = iCurColumnPos(5)
'			C_TEMP_GL_NO		= iCurColumnPos(5)
			C_GL_NO				= iCurColumnPos(6)
			C_COST_MVMT			= iCurColumnPos(7)
			C_COST_MVMT_NM		= iCurColumnPos(8)
			C_REF_NO            = iCurColumnPos(9)
		Case "B"
 			ggoSpread.Source = frm1.vspdData2

 			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

 			C_ItemDocumentNo	= iCurColumnPos(1)
			C_DocumentDt		= iCurColumnPos(2)
			C_PosDt				= iCurColumnPos(3)
			C_ParentRowNo		= iCurColumnPos(4)
			C_Flag				= iCurColumnPos(5)
 	End Select
End Sub

'========================================== 2.4.2 Open???()  =============================================
Function OpenbizareaInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장 팝업"
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
	ggoSpread.Source = frm1.vspdData2

    If frm1.vspdData.MaxRows = 0 Then
		Call DisplayMsgBox("169804","X", "X", "X")
		Exit Function	
	Else
		frm1.vspdData.Col = C_TrnsTypeNm  : frm1.vspdData.Row = frm1.vspdData.ActiveRow : Param6 = Trim(frm1.vspdData.Text )
		frm1.vspdData.Col = C_MovType     : frm1.vspdData.Row = frm1.vspdData.ActiveRow : Param7 = Trim(frm1.vspdData.Text )	
    End If
    
	With frm1.vspdData2	    
		If .MaxRows = 0 Then
		    Call DisplayMsgBox("169804","X", "X", "X")
			Exit Function
		else
		   .Col = C_ItemDocumentNo  : .Row = .ActiveRow : Param3 = Trim(.Text )
		   .Col = C_DocumentDt		: .Row = .ActiveRow : Param4 = Trim(.Text )
		   .Col = C_PosDt			: .Row = .ActiveRow : Param5 = Trim(.Text )
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

'======================================================================================================
' Function Name : OpenPopupGL
' Function Desc : This method Open The Popup window for GL
'=======================================================================================================
Function OpenPopupGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5120ra1")
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5120ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function
	
	frm1.vspddata.col = C_GL_NO
	
	arrParam(0) = Trim(frm1.vspddata.Text)							'회계전표번호 
	arrParam(1) = ""												'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'======================================================================================================
' Function Name : OpenPopupTempGL
' Function Desc : This method Open The Popup window for TempGL
'=======================================================================================================
Function OpenPopuptempGL()
	Dim arrRet
	Dim arrParam(1)	
	Dim iCalledAspName
	
	iCalledAspName = AskPRAspName("a5130ra1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a5130ra1", "X")
		IsOpenPop = False
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	frm1.vspddata.col = C_GL_NO
	
	arrParam(0) = Trim(frm1.vspddata.Text)							'결의전표번호 
	arrParam(1) = ""												'Reference번호 

	IsOpenPop = True
   
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
End Function

'========================================================================================
' Function Name : Checkall()
'========================================================================================
Function Checkall()
	Dim IRowCount 
	Dim IClnCount
	
	ggoSpread.Source = frm1.vspdData
	
	With frm1.vspdData    
		If lgCheckall = 0 Then 
			For IClnCount = 0 To C_Check
	  			For IRowCount = 1 To .MaxRows
	  				If IClnCount <> 0 Then   	     	 
	       			.Row = IRowCount 
	       			.Col = IClnCount	 
					    .text = 1     
					Else
					    .Row = IRowCount
					    .Col = IClnCount
					    .Text =ggoSpread.UpdateFlag
					End if
				Next    
			Next

			lgCheckall = 1
			lgBlnFlgChgValue = True
		Else
			For IClnCount = 0 To C_Check
	  			For IRowCount = 1 To .MaxRows
	  				If IClnCount <> 0 Then
	       			 .Row = IRowCount 
	       			 .Col = IClnCount	 
					    .text = 0     
					End If
				Next    
			Next

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
    
    Call InitSpreadSheet2
    Call InitVariables
                      
    Call SetDefaultVal
    Call InitComboBox
    Call SetToolbar("11000000000011")
		
End Sub

'=======================================================================================================
'   Event Name : txtBizCd_LostFocus()
'   Event Desc : 공장명과 최종마감년월을 찾는다.
'=======================================================================================================
Sub txtBizCd_LostFocus()
    Dim strYear
    Dim strMonth
    Dim strDay
	
	If frm1.txtBizCd.value <> "" Then
		If  CommonQueryRs(" PLANT_NM, CONVERT(CHAR(10), INV_CLS_DT, 21) "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value,"","S"), _
		 lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
			frm1.txtPlantNm.Value  = ""
			frm1.txtInvClsDt.text  = ""
			Exit Sub
		Else
			lgF0 = Split(lgF0,Chr(11))
			lgF1 = Split(lgF1,Chr(11))
			
			frm1.txtBizNm.Value = lgF0(0)
			Call ExtractDateFrom(lgF1(0), Parent.gServerDateFormat, Parent.gServerDateType , strYear, strMonth, strDay)

			If CommonQueryRs("CLOSE_DT","I_INV_CLOSING_HISTORY","CLOSE_FLAG = 'Y' AND PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value,"","S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) Then										
				frm1.btnCancel.disabled = false
			Else				
				frm1.btnCancel.disabled = true
			End if
		End If

		frm1.txtInvClsDt.Year  =  strYear
		frm1.txtInvClsDt.Month =  strMonth
	Else
		frm1.txtPlantNm.Value  = ""
		frm1.txtInvClsDt.text  = ""
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDt_DblClick(Button)
'=======================================================================================================
Sub txtDocumentDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDt_KeyPress()
'=======================================================================================================
Sub txtDocumentDt_KeyPress(KeyAscii)
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
	if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2, NewTop) and lgStrPrevKey <> "" Then
		Call DisableToolBar(Parent.TBC_QUERY)
		If DbdtlQuery = False Then
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

'=======================================================================================================
'   Event Name : vspdData_ScriptLeaveCell
'   Event Desc :
'=======================================================================================================
Sub vspdData_ScriptLeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	
	'If lgSpdHdrClicked = 1 Then	Exit Sub  '2003-03-01 Release 추가 
	
	If Row <> NewRow And NewRow > 0 Then
		'/* 다른 작업이 이루어지는 상황에서 다른 행 이동 시 조회가 이루어 지지 않도록 한다. - START */
		If CheckRunningBizProcess = True Then
			frm1.vspdData.Row = Row
			frm1.vspdData.Col = Col	
			frm1.vspdData.Action = 0
			Exit Sub
		End If
		'/* 다른 작업이 이루어지는 상황에서 다른 행 이동 시 조회가 이루어 지지 않도록 한다. - END */
			
		If DbDtlQuery(NewRow) = False Then	Exit Sub
	End If
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
	
	ggoSpread.Source = frm1.vspdData2
	ggoSpread.ClearSpreadData

    Call InitVariables

	If 	CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA ", " BIZ_AREA_CD = " & Trim(FilterVar(frm1.txtBizCd.value," ","S")), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false Then
	
		Call DisplayMsgBox("124200","X","X","X")
		frm1.txtBizNm.value = ""
		frm1.txtBizCd.Focus		
		Exit function
    
    End If

    lgF0 = Split(lgF0,Chr(11))
	frm1.txtBizNm.value = lgF0(0)

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
	Dim strYear, strMonth, strDay, strYyMm
	
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
	
	Call ExtractDateFrom(frm1.txtDocumentDt.Text,frm1.txtDocumentDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)
	
	strYyMm = strYear + strMonth
	frm1.hYyMm.value = strYyMm
	
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
    Dim strCostFlag
    Dim strYear, strMonth, strDay, strYyMm
    
    If frm1.RadioOutputType.rdoCase1.Checked Then
		strCostFlag = ""
	Elseif frm1.RadioOutputType.rdoCase2.Checked Then
		strCostFlag = "N"
	Elseif frm1.RadioOutputType.rdoCase3.Checked Then
		strCostFlag = "Y"			
	End If

	Call ExtractDateFrom(frm1.txtDocumentDt.Text,frm1.txtDocumentDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

	strYyMm = strYear + strMonth

    With frm1
		strVal = BIZ_PGM_QRY_ID &	"?txtBiZCd="& Trim(.txtBizCd.value)	& _	
									"&txtDocumentDt="& strYyMm & _
									"&cboTrnsType="& Trim(.cboTrnsType.value) & _
									"&txtCostMvmt="& Trim(strCostFlag) & _									
									"&txtMaxRows="& .vspdData.MaxRows

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
	
	If DbDtlQuery(1) = False Then Exit Function
End Function

'=======================================================================================================
' Function Name : DbDtlQuery																				
' Function Desc : This function is data query and display												
'=======================================================================================================
Function DbDtlQuery(ByVal Row)
	Dim strVal
	Dim strCostFlag,strMoveType
    Dim strYear, strMonth, strDay, strYyMm

	DbDtlQuery = False

	Call LayerShowHide(1)

'	If Trim(lgStrPrevKey) = "" Then
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.ClearSpreadData
'	End If
	
	Call ExtractDateFrom(frm1.txtDocumentDt.Text,frm1.txtDocumentDt.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

	strYyMm = strYear + strMonth

	With frm1
		.vspdData.Row = Row
		.vspdData.Col = C_MovType
		strMoveType = .vspdData.Text

		strVal = BIZ_PGM_QRY2_ID & "?txtMoveType=" & strMoveType _
								 & "&txtBiZCd=" & Trim(.txtBizCd.value)	_
								 & "&txtDocumentDt="	& strYyMm _
								 & "&lgStrPrevKey=" & lgStrPrevKey _
								 & "&txtMaxRows=" & .vspdData2.MaxRows
	End With

	Call RunMyBizASP(MyBizASP, strVal)
	DbDtlQuery = True
End Function

'=======================================================================================================
' Function Name : DbDtlQueryOk
' Function Desc : DbDtlQuery가 성공적일 경우 MyBizASP 에서 호출되는 Function
'=======================================================================================================
Function DbDtlQueryOk()
	DbDtlQueryOk = false	
	frm1.vspdData.Focus
	DbDtlQueryOk = true
End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim strVal, strYear, strMonth, strDay
    Dim strCostFlag
	Dim iRowSep, iColSep
	
'	Dim strCUTotalvalLen
'	Dim objTEXTAREA
'	Dim iTmpCUBuffer
'	Dim iTmpCUBufferCount
'	Dim iTmpCUBufferMaxCount
	
	iRowSep = Parent.gRowSep
	iColSep = Parent.gColSep

    Err.Clear		
    DbSave = False

    On Error Resume Next
	
'	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
'	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
'	iTmpCUBufferCount = -1
'	strCUTotalvalLen = 0
	strVal = ""

    Call LayerShowHide(1)

	frm1.txtMode.value = Parent.UID_M0002
	frm1.hTrnsType.value = frm1.cboTrnsType.value
	frm1.hCostMvmt.value = frm1.cboCostMvmt.value

    '-----------------------
    'Data manipulate area
    '-----------------------
	With frm1.vspdData
		For lRow = 1 To .MaxRows
			.Row = lRow
			.Col = 0

			Select Case .Text
				Case ggoSpread.UpdateFlag
					.Col = C_Check
					If .Text = "1" Then
						strVal = strval & lRow & iColSep 
						.col = C_MovType
						strVal = strVal & Trim(.Text) & iColSep		 	  			
						.col = C_COST_MVMT
						strVal = strVal & Trim(.Text) & iColSep		 	  									
						.Col = C_REF_NO
		 	  			strVal = strVal & Trim(.Text) & iRowSep		 	  			
					End If

'					If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then
'						Set objTEXTAREA = document.createElement("TEXTAREA")
'						objTEXTAREA.name = "txtCUSpread"
'						objTEXTAREA.value = Join(iTmpCUBuffer,"")
'						divTextArea.appendChild(objTEXTAREA)     
'									 
'						iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT   
'						ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
'						iTmpCUBufferCount = -1
'						strCUTotalvalLen  = 0
'					End If
'								       
'					iTmpCUBufferCount = iTmpCUBufferCount + 1
'								      
'					If iTmpCUBufferCount > iTmpCUBufferMaxCount Then    
'						iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
'						ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
'					End If   
'											
'					iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
'					strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
			End Select
		Next
	End With

	frm1.txtSpread.value = strval
'	If iTmpCUBufferCount > -1 Then 
'		Set objTEXTAREA = document.createElement("TEXTAREA")
'		objTEXTAREA.name   = "txtCUSpread"
'		objTEXTAREA.value = Join(iTmpCUBuffer,"")
'		divTextArea.appendChild(objTEXTAREA)     
'	Else
'		Call DisplayMsgBox("169801","X", "X", "X")
'		Exit Function
'	End If  
	
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0">
							<TR>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif"><IMG SRC="../../../CShared/image/table/seltab_up_left.gif" WIDTH="9" HEIGHT="23"></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="center" CLASS="CLSMTAB"><FONT COLOR=white>Batch Posting 취소(Monthly)</FONT></TD>
								<TD BACKGROUND="../../../CShared/image/table/seltab_up_bg.gif" ALIGN="right"><IMG SRC="../../../CShared/image/table/seltab_up_right.gif" WIDTH="10" HEIGHT="23"></TD>
						    	</TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPopupGL()">회계전표정보</A> | <A href="vbscript:OpenPopuptempGL()">결의전표정보</A> | <A href="vbscript:OpenMoveDtlRef()">수불상세정보</A></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD  WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizCd" SIZE=10 MAXLENGTH=10 tag="12XXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBizCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenbizareaInfo()">&nbsp;
															<INPUT TYPE=TEXT NAME="txtBizNm" SIZE=40 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>수불년월</TD>
									<TD CLASS="TD6" NOWRAP>
								      <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMM name=txtDocumentDt CLASSID=<%=gCLSIDFPDT%> tag="12x1" ALT="수불년월" VIEWASTEXT id=OBJECT3> <PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""> </OBJECT>');</SCRIPT>
									</TD>      
								</TR>
								<TR>
					           		<TD CLASS="TD5" NOWRAP>수불구분</TD>
									<TD CLASS="TD6" NOWRAP>
										<SELECT Name="cboTrnsType" ALT="수불구분" STYLE="WIDTH: 150px" tag="11"><OPTION Value=""></OPTION></SELECT>
									</TD>
					           		<TD CLASS="TD5" NOWRAP>수불경로</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X" checked><LABEL FOR="rdoCase1">전체</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X"><LABEL FOR="rdoCase2">재고</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase3" TAG="1X"><LABEL FOR="rdoCase3">원가보정</LABEL>										
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
					<TD HEIGHT=*  WIDTH=100% VALIGN=TOP>						
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT=100% WIDTH=100% Colspan=4>
									<TABLE <%=LR_SPACE_TYPE_20%>>
										<TR HEIGHT="*">
											<TD HEIGHT=100% WIDTH=65% Colspan=4>
												<TABLE <%=LR_SPACE_TYPE_20%>>
													<TR HEIGHT="*" WIDTH="100%">
														<TD WIDTH="100%" Colspan=4>
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="22" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MAXCOLs" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
														</TD>
													</TR>
												</TABLE>
											</TD>
											<TD HEIGHT=100% WIDTH=35%>
												<TABLE <%=LR_SPACE_TYPE_20%>>
													<TR HEIGHT="*">
														<TD WIDTH="100%">
															<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="22" TITLE="SPREAD" id=OBJECT2> <PARAM NAME="MAXCOLs" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
														</TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noreSIZE framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>	
<INPUT TYPE=HIDDEN NAME="SpdCount" tag="24" TABINDEX="-1">	
<TEXTAREA CLASS=HIDDEN NAME=txtSpread Width=100% tag="24" TABINDEX="-1"></TEXTAREA>
	<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hYyMm" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hCostMvmt" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="hTrnsType" tag="24" TABINDEX="-1">	
	<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
	<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=320 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
</BODY>
</HTML>
