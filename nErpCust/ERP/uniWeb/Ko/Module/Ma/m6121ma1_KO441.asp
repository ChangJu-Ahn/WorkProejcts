<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : m6121ma1
'*  4. Program Name         : 부대비일괄배부취소 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2004/11/04
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Byun Jee Hyun
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. History              :
'*                            
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		
Const BIZ_PGM_ID = "m6121mb1_KO441.asp"			

'==========================================  1.2.1 Global 상수 선언  ======================================
Dim C_DocumentNo
Dim C_RefNo 
Dim C_PlantCd 
Dim C_PlantNm
Dim C_DisbDt 
Dim C_BatchJobDt 
Dim C_TotDisbAmt
Dim C_DisbFrDt 
Dim C_DisbToDt 	
Dim C_ProcessStepCd 
Dim C_ProcessStepNm 

'==========================================  1.2.2 Global 변수 선언  =====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim IsOpenPop

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE 
    lgBlnFlgChgValue = False  
    lgIntGrpCount = 0         
    lgStrPrevKey = ""         
    lgLngCurRows = 0     
    frm1.vspdData.operationmode = 3     
    frm1.vspdData.MaxRows = 0
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadPosVariables() ========================================
Sub InitSpreadPosVariables()
	C_DocumentNo = 1
	C_RefNo = 2
	C_PlantCd = 3
	C_PlantNm = 4
	C_DisbDt = 5
	C_BatchJobDt = 6
	C_TotDisbAmt = 7
	C_DisbFrDt = 8
	C_DisbToDt 	= 9
	C_ProcessStepCd = 10 
	C_ProcessStepNm = 11
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20061105",,Parent.gAllowDragDropSpread  
		.ReDraw = false

		.MaxCols = C_ProcessStepNm+1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols:    .ColHidden = True
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit	 C_DocumentNo		, "재고처리번호",18,,,,2
		ggoSpread.SSSetEdit	 C_RefNo		, "배부참조번호",18,,,,2
		ggoSpread.SSSetEdit  C_PlantCd		, "공장",10,,,,2
		ggoSpread.SSSetEdit  C_PlantNm		, "공장명",20
		ggoSpread.SSSetDate  C_DisbDt		, "작업일", 10, 2, Parent.gDateFormat
		ggoSpread.SSSetDate  C_BatchJobDt	, "배부년월", 10, 2, Parent.gDateFormat
		SetSpreadFloatLocal  C_TotDisbAmt	, "총배부금액", 20,1,2
		ggoSpread.SSSetDate  C_DisbFrDt		, "배부대상기간(From)", 20, 2, Parent.gDateFormat
		ggoSpread.SSSetDate  C_DisbToDt		, "배부대상기간(To)", 20, 2, Parent.gDateFormat
		ggoSpread.SSSetEdit  C_ProcessStepCd	, "경비발생단계",13,,,,2
		ggoSpread.SSSetEdit  C_ProcessStepNm	, "경비발생단계명",20	

		Call SetSpreadLock 
    
		.ReDraw = true
    End With
end sub

'------------------------------------------  OpenPlant()  ------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function
    
	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.value= arrret(1)
	End If	
End Function

Function OpenProcessStep()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "경비발생단계"					
	arrParam(1) = "B_minor"						
	arrParam(2) = frm1.txtProcessStep.value	
	arrParam(3) = ""							
	arrParam(4) = "major_cd=" & FilterVar("M9014", "''", "S") & ""			
	arrParam(5) = "경비발생단계"			
	
    arrField(0) = "minor_cd"					
    arrField(1) = "minor_nm"					
    
    arrHeader(0) = "경비발생단계"				
    arrHeader(1) = "경비발생단계명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtProcessStep.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtProcessStep.value		= arrRet(0)
		frm1.txtProcessStepNm.value	= arrRet(1)
		frm1.txtProcessStep.focus
		Set gActiveElement = document.activeElement
	End If	

End Function

'==========================================================================================
'   Event Name : SetSpreadFloatLocal
'==========================================================================================
Sub SetSpreadFloatLocal(ByVal iCol , ByVal Header , ByVal dColWidth , ByVal HAlign , ByVal iFlag )
   Select Case iFlag
        Case 2                                                              '금액 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign
        Case 3                                                              '수량 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggQtyNo       ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"P"
        Case 4                                                              '단가 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggUnitCostNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"P"
        Case 5                                                              '환율 
            ggoSpread.SSSetFloat iCol, Header, dColWidth, Parent.ggExchRateNo  ,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec, HAlign,,"P"
    End Select
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029                                               '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                 '⊙: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                                                   '⊙: Setup the Spread sheet
    Call InitVariables 
    Call GetValue_ko441()
    
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPGCd
	End If
    
    frm1.txtPlantCd.focus
    Set gActiveElement = document.activeElement
    
    Call SetToolbar("1110000000001111")                                                   '⊙: Initializes local global variables
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
    With frm1
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			
			C_DocumentNo = iCurColumnPos(1)  
			C_RefNo	= iCurColumnPos(2)
			C_PlantCd = iCurColumnPos(3)
			C_PlantNm = iCurColumnPos(4)
			C_DisbDt = iCurColumnPos(5)
			C_BatchJobDt = iCurColumnPos(6)
			C_TotDisbAmt = iCurColumnPos(7)
			C_DisbFrDt = iCurColumnPos(8)
			C_DisbToDt 	= iCurColumnPos(9)
			C_ProcessStepCd = iCurColumnPos(10)
			C_ProcessStepNm = iCurColumnPos(11)	
	End Select

End Sub	
   
'==========================================================================================
'   Event Name : txtDisbDt  	 
'==========================================================================================
Sub txtFrDisbDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDisbDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFrDisbDt.Focus
	End if
End Sub

Sub txtToDisbDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDisbDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToDisbDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : txtBatchJobDt
'==========================================================================================
Sub txtFrBatchJobDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrBatchJobDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFrBatchJobDt.Focus
	End if
End Sub

Sub txtToBatchJobDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToBatchJobDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToBatchJobDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtFrDisbDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

Sub txtToDisbDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

Sub txtFrBatchJobDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

Sub txtToBatchJobDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If
End Sub

'========================================================================================
' Function Name : FncSplitColumn
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
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery()
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear  
    
    Call InitVariables                                                             '☜: Protect system from crashing

	ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
	Set gActiveElement = document.activeElement
    FncQuery = True																'⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    Err.Clear                                                               '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
        
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
 	
	Set gActiveElement = document.activeElement
    FncNew = True                                                           '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    Err.Clear                                                               '☜: Protect system from crashing
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If
    
    'ggoSpread.Source = frm1.vspdData
    
    'If ggoSpread.SSCheckChange = False Then
     '   IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
     '   Exit Function
    'End If
        
	If DbSave = False Then Exit Function
    
	Set gActiveElement = document.activeElement
    FncSave = True                                                          '⊙: Processing is OK
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

Function FncExcel()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncExport(Parent.C_SINGLE)												<%'☜: 화면 유형 %>
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_SINGLE , False)                                     <%'☜:화면 유형, Tab 유무 %>
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	Set gActiveElement = document.activeElement
    FncExit = True
End Function


'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      
    
    DbQuery = False
    
    If LayerShowHide(1) = False Then Exit Function
    
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
    
    With frm1
    
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtPlantCd=" & .txtPlantCd.value
	    strVal = strVal & "&txtProcessStep=" & .txtProcessStep.value
		strVal = strVal & "&txtFrDisbDt=" & .txtFrDisbDt.text           
		strVal = strVal & "&txtToDisbDt=" & .txtToDisbDt.text		   
		strVal = strVal & "&txtFrBatchJobDt=" & .txtFrBatchJobDt.text
		strVal = strVal & "&txtToBatchJobDt=" & .txtToBatchJobDt.text
	    strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows   

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  
	   
	Call RunMyBizASP(MyBizASP, strVal)
        
    End With
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
    lgIntFlgMode = Parent.OPMD_UMODE												'⊙: Indicates that current mode is Update mode
    
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
    
    Call SetToolbar("11101001000111")
	
	frm1.vspddata.focus
	Set gActiveElement = document.activeElement

End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 
	Dim ColSep, RowSep
	
    DbSave = False                                                        '⊙: Processing is NG
    
    If LayerShowHide(1) = False Then Exit Function
 
	With frm1
		.txtMode.value = Parent.UID_M0002
		.hdnDocumentNo.value = Trim(GetSpreadText(.vspdData,C_DocumentNo,.vspdData.ActiveRow,"X","X"))
		.hdnDisbBatchJobDt.value = Trim(GetSpreadText(.vspdData,C_BatchJobDt,.vspdData.ActiveRow,"X","X"))
		.hdnProcessStep.value = Trim(GetSpreadText(.vspdData,C_ProcessStepCd,.vspdData.ActiveRow,"X","X"))
        .hdnPlantCd.value = Trim(GetSpreadText(.vspdData,C_PlantCd,.vspdData.ActiveRow,"X","X"))
        .hdnDisbQryDt.value = Trim(GetSpreadText(.vspdData,C_DisbToDt,.vspdData.ActiveRow,"X","X"))
        .hdnDisbFrQryDt.value = Trim(GetSpreadText(.vspdData,C_DisbFrDt,.vspdData.ActiveRow,"X","X"))
        .hdnRefNo.value = Trim(GetSpreadText(.vspdData,C_RefNo,.vspdData.ActiveRow,"X","X"))
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
	
	End With
	
    DbSave = True                                                           '⊙: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()													'☆: 저장 성공후 실행 로직 
	Call InitVariables
	
	lgBlnFlgChgValue = False
	
	Call FncQuery()
	
End Function

'------------------------------------------  OpenDisbRef()  -------------------------------------------------
'	Name : OpenDisbRef()
'	Description :배부내역참조 
'---------------------------------------------------------------------------------------------------------
Function OpenDisbRef()

	Dim strRet
	Dim arrParam(5)
	Dim iCalledAspName
	Dim IntRetCD
	
	if lgIntFlgMode = Parent.OPMD_CMODE then
		Call DisplayMsgBox("900002", "X", "X", "X")
		Exit Function
	End if 
	
	if Trim(GetSpreadText(frm1.vspdData,C_DocumentNo,frm1.vspdData.ActiveRow,"X","X")) = "Y" then
		Call DisplayMsgBox("209001", "X", "X", "X")
		Exit Function
	End if
	
	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True
	
	arrParam(0) = Trim(GetSpreadText(frm1.vspdData,C_PlantCd,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(1) = Trim(GetSpreadText(frm1.vspdData,C_PlantNm,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(2) = Trim(GetSpreadText(frm1.vspdData,C_BatchJobDt,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(3) = Trim(GetSpreadText(frm1.vspdData,C_DisbDt,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(4) = Trim(GetSpreadText(frm1.vspdData,C_DocumentNo,frm1.vspdData.ActiveRow,"X","X"))
	arrParam(5) = Trim(GetSpreadText(frm1.vspdData,C_RefNo,frm1.vspdData.ActiveRow,"X","X"))
	
	iCalledAspName = AskPRAspName("M6121RA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M6121RA1", "X")
		lblnWinEvent = False
		Exit Function
	End If

	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	
	IsOpenPop = False
	
	If strRet(0,0) = "" Then
		Exit Function
	Else
		Call SetReqRef(strRet)
	End If
		
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
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 border="0">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSLTAB"><font color=white>부대비일괄배부취소</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenDisbRef">배부내역참조</A> </TD>
					<TD WIDTH=10>&nbsp;</TD>
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
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" ALT="공장" SIZE=10 MAXLENGTH=4  tag="11N1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT ID="txtPlantNm" ALT="공장" NAME="txtPlantNm" tag="14X"></TD>
								<TD CLASS="TD5" NOWRAP>경비발생단계</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtProcessStep" ALT="경비발생단계" SIZE=10 MAXLENGTH=5  tag="11N1"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDocumentNo1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenProcessStep()">
													   <INPUT TYPE=TEXT ID="txtProcessStepNm" ALT="발생단계명" NAME="txtProcessStepNm" tag="14X"></TD>							
							<TR>
								<TD CLASS="TD5" NOWRAP>작업일</TD>
								<TD CLASS="TD6" NOWRAP>
								<table cellspacing=0 cellpadding=0>
									<tr>
										<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtFrDisbDt NAME="txtFrDisbDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="작업일"></OBJECT>');</SCRIPT></TD>
										</td>
										<TD> ~ </TD>
										<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtToDisbDt NAME="txtToDisbDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="작업일"></OBJECT>');</SCRIPT></TD>
										</td>
									</tr>
								</table></TD>
								<TD CLASS="TD5" NOWRAP>배부년월</TD>
								<TD CLASS="TD6" NOWRAP>
								<table cellspacing=0 cellpadding=0>
									<tr>
										<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtFrBatchJobDt name="txtFrBatchJobDt" CLASS=FPDTYYYYMM title="FPDATETIME" ALT="배부년월" tag="11XXXU"></OBJECT>');</SCRIPT></TD> 
										</td>
										<TD> ~ </TD>
										<td>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtToBatchJobDt name="txtToBatchJobDt" CLASS=FPDTYYYYMM title="FPDATETIME" ALT="배부년월" tag="11XXXU"></OBJECT>');</SCRIPT></TD>
										</td>
									</tr>
									</table></TD>
							</TR>
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
			<TR>
				<TD WIDTH=100% valign=top>
					<TABLE <%=LR_SPACE_TYPE_20%>>
						<TR>
							<TD HEIGHT="100%">
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hProcessStep" tag="24"><INPUT TYPE=HIDDEN NAME="hFrDisbDt" tag="24"><INPUT TYPE=HIDDEN NAME="hToDisbDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hFrBatchJobDt" tag="24"><INPUT TYPE=HIDDEN NAME="hToBatchJobDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDocumentNo" tag="24"><INPUT TYPE=HIDDEN NAME="hdnDisbBatchJobDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnProcessStep" tag="24"><INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hdnDisbQryDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnDisbFrQryDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRefNo" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
