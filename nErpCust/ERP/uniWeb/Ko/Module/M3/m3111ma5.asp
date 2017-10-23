
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111ma5
'*  4. Program Name         : 발주마감 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/14
'*  8. Modified date(Last)  : 2000/04/14
'*  9. Modifier (First)     : Shin Jin Hyun
'* 10. Modifier (Last)      : Min, HJ
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*							  2000/04/14
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		

Const BIZ_PGM_ID = "m3111mb5.asp"											

Dim C_Check
Dim C_CloseFlg
Dim C_PoNo
Dim C_PoSeq
Dim C_PlantCd
Dim C_PlantNm
Dim C_ItemCd
Dim C_ItemNm
Dim C_ItemSpec
Dim C_dlvy_dt
Dim C_PoDt
Dim C_PoQty
Dim C_PoUnit
Dim C_GrQty
Dim C_IvQty
Dim C_LcQty
Dim C_BlQty
Dim C_CcQty
Dim C_InspectQty
Dim C_SupplierCd
Dim C_SupplierNm

Const C_SHEETMAXROWS=100

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim StartDate,EndDate
 
EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  

Dim IsOpenPop          


'==========================================   Selection()  ======================================
'	Name : Selection()
'	Description : 일괄선택버튼의 Event 합수 
'================================================================================================
Sub Selection()

Dim index,Count
	
	frm1.vspdData.ReDraw = false
	
	Count = frm1.vspdData.MaxRows 
	
	For index = 1 to Count
		
		frm1.vspdData.Row = index
		frm1.vspdData.Col = C_Check
		
		if frm1.vspdData.Text = "1" then
			frm1.vspdData.Text = "0"
		else
			frm1.vspdData.Text = "1"
		End if
		
		frm1.vspdData.Col = 0 
		
		if ggoSpread.UpdateFlag = frm1.vspdData.Text then    
			'ggoSpread.EditUndo index
			frm1.vspdData.Text=""
	    else
	    	ggoSpread.UpdateRow Index
		End if
				
	Next 
	
	frm1.vspdData.ReDraw = true
		
End Sub


'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE  
    lgBlnFlgChgValue = False   
    lgIntGrpCount = 0          
    lgStrPrevKey = ""          
    lgLngCurRows = 0           
    frm1.vspdData.MaxRows = 0
    
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtPur_Grp.focus
	Set gActiveElement = document.activeElement
	frm1.txtPur_Grp.Value = Parent.gPurGrp
    frm1.txtFrDt.Text = StartDate
    frm1.txtToDt.Text = EndDate
    frm1.txtCfmFlag.value = frm1.rdoCfmAll.value
	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True

	Call SetToolbar("1110000000001111")
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
	C_Check		= 1
	C_CloseFlg  = 2
	C_PoNo		= 3
	C_PoSeq		= 4
	C_PlantCd	= 5
	C_PlantNm	= 6
	C_ItemCd	= 7
	C_ItemNm	= 8
	C_ItemSpec	= 9
	C_dlvy_dt	= 10
	C_PoDt		= 11
	C_PoQty		= 12
	C_PoUnit	= 13
	C_GrQty		= 14
	C_IvQty		= 15
	C_LcQty		= 16
	C_BlQty		= 17
	C_CcQty		= 18
	C_InspectQty= 19
	C_SupplierCd= 20
	C_SupplierNm= 21
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	With frm1.vspdData

    ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20061213",,Parent.gAllowDragDropSpread  

	.ReDraw = false	
	
    .MaxCols = C_SupplierNm + 1														
	.Col = .MaxCols:    .ColHidden = True											
    .MaxRows = 0

	Call GetSpreadColumnPos("A")

	ggoSpread.SSSetCheck C_Check, "선택",10,,,true
	ggoSpread.SSSetEdit C_CloseFlg, "마감여부", 10
	ggoSpread.SSSetEdit C_PoNo,"발주번호",18
    ggoSpread.SSSetEdit C_PoSeq, "발주순번", 10
	ggoSpread.SSSetEdit C_PlantCd, "공장",10
	ggoSpread.SSSetEdit C_PlantNm, "공장명",20
	ggoSpread.SSSetEdit C_ItemCd, "품목",10
	ggoSpread.SSSetEdit C_ItemNm, "품목명",20
	ggoSpread.SSSetEdit C_ItemSpec, "품목규격",20
	ggoSpread.SSSetDate C_dlvy_dt,"납기일", 10, 2, Parent.gDateFormat
    ggoSpread.SSSetDate C_PoDt,"발주일", 10, 2, Parent.gDateFormat
    SetSpreadFloat	 	C_PoQty, "발주수량", 15, 1, 3
    ggoSpread.SSSetEdit C_PoUnit, "단위", 10
    SetSpreadFloat	 	C_GrQty, "입고수량", 15, 1, 3
    SetSpreadFloat	 	C_IvQty, "매입수량", 15, 1, 3
    SetSpreadFloat	 	C_LcQty, "L/C수량", 15, 1, 3
    SetSpreadFloat	 	C_BlQty, "B/L수량", 15, 1, 3
    SetSpreadFloat	 	C_CcQty, "통관수량", 15, 1, 3
    SetSpreadFloat	 	C_InspectQty, "검사중수량", 15, 1, 3
    ggoSpread.SSSetEdit C_SupplierCd, "공급처", 10
    ggoSpread.SSSetEdit C_SupplierNm, "공급처명", 20
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
    With frm1
		Dim IRow_1, Maxrow
    
		.vspdData.ReDraw = False
    
		Maxrow = .vspdData.Maxrows
		ggoSpread.SpreadLock	-1, -1
        
		For IRow_1 =1 to Maxrow
				ggoSpread.SpreadUnLock	C_Check,	IRow_1,		C_Check,	IRow_1
		Next
	
		.vspdData.ReDraw = True

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

			C_Check		= iCurColumnPos(1)
			C_CloseFlg  = iCurColumnPos(2)
			C_PoNo		= iCurColumnPos(3)
			C_PoSeq		= iCurColumnPos(4)
			C_PlantCd	= iCurColumnPos(5)
			C_PlantNm	= iCurColumnPos(6)
			C_ItemCd	= iCurColumnPos(7)
			C_ItemNm	= iCurColumnPos(8)
			C_ItemSpec	= iCurColumnPos(9)
			C_dlvy_dt	= iCurColumnPos(10)
			C_PoDt		= iCurColumnPos(11)
			C_PoQty		= iCurColumnPos(12)
			C_PoUnit	= iCurColumnPos(13)
			C_GrQty		= iCurColumnPos(14)
			C_IvQty		= iCurColumnPos(15)
			C_LcQty		= iCurColumnPos(16)
			C_BlQty		= iCurColumnPos(17)
			C_CcQty		= iCurColumnPos(18)
			C_InspectQty= iCurColumnPos(19)
			C_SupplierCd= iCurColumnPos(20)
			C_SupplierNm= iCurColumnPos(21)
	End Select

End Sub	

'------------------------------------------  OpenPoNo()  -------------------------------------------------
Function OpenPoNo()
	
	Dim strRet
	Dim arrParam
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	IsOpenPop = True
	
	Redim arrParam(2)
	arrParam(0) = "N"  'Return Flag
	arrParam(1) = "Y"  'Release Flag
	arrParam(2) = ""  'STO Flag

	iCalledAspName = AskPRAspName("M3111PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If strRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
	End If	
		
End Function

'------------------------------------------  OpenSupplier()  -------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"						
	arrParam(1) = "B_BIZ_PARTNER"					

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)	
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "공급처"				
    arrHeader(1) = "공급처명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)	
	End If	
End Function

'------------------------------------------  OpenPurGrp()  -------------------------------------------------
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"					
	arrParam(1) = "B_PUR_GRP"					

	arrParam(2) = Trim(frm1.txtPur_Grp.Value)	
'	arrParam(3) = Trim(frm1.txtPur_Grp_Nm.Value)
	
	arrParam(4) = ""							
	arrParam(5) = "구매그룹"					
	
    arrField(0) = "PUR_GRP"						
    arrField(1) = "PUR_GRP_NM"					
    
    arrHeader(0) = "구매그룹"				
    arrHeader(1) = "구매그룹명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPur_Grp.Value		= arrRet(0)		
		frm1.txtPur_Grp_Nm.Value	= arrRet(1)	
	End If	
End Function
'------------------------------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
			 
	arrParam(0) = "공장" 
	arrParam(1) = "B_Plant"    
	 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	 
	arrParam(4) = ""   
	arrParam(5) = "공장"
	 
	arrField(0) = "Plant_Cd" 
	arrField(1) = "Plant_NM" 
	    
	arrHeader(0) = "공장"
	arrHeader(1) = "공장명"  
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)  
		frm1.txtPlantNm.Value= arrRet(1)  
		frm1.txtPlantCd.focus
	End If  
End Function

'------------------------------------------  OpenPoType()  -------------------------------------------------
'	Name : OpenPoType()
'	Description : OpenPoType PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPotype()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPotypeCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "발주형태"					
	arrParam(1) = "M_CONFIG_PROCESS"			
	
	arrParam(2) = Trim(frm1.txtPotypeCd.Value)
	'arrParam(3) = Trim(frm1.txtPotypeNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") &  " "							
	'arrParam(4) = "USAGE_FLG='Y'"							
	arrParam(5) = "발주형태"					
	
    arrField(0) = "PO_TYPE_CD"					
    arrField(1) = "PO_TYPE_NM"					
    
    arrHeader(0) = "발주형태"				
    arrHeader(1) = "발주형태명"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPotypeCd.focus
		Exit Function
	Else
		frm1.txtPoTypeCd.Value    = arrRet(0)		
		frm1.txtPoTypeNm.Value    = arrRet(1)
		lgBlnFlgChgValue = True

		'Call PotypeRef()
		frm1.txtPotypeCd.focus
	End If	
	Set gActiveElement = document.activeElement 
End Function


'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemInfo()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemInfo(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName

	If IsOpenPop = True Then Exit Function

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = strCode						' Item Code
	arrParam(2) = ""						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"							' Default Value
	
	arrField(0) = 1 '"ITEM_CD"					' Field명(0)
	arrField(1) = 2 '"ITEM_NM"					' Field명(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
		End With
	End If	
	
	Call SetFocusToDocument("M")
		frm1.txtItemCd.focus
	

End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029                         
    Call ggoOper.LockField(Document, "N")       
    
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitSpreadSheet                        
    Call InitVariables                          
    Call SetDefaultVal
End Sub
'==========================================================================================
'   Event Name : Form_QueryUnload
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'==========================================================================================
'   Event Name : btnPosting_OnClick()
'   Event Desc : 출고처리 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnSelect_OnClick()
	Dim i, ClsFlg
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Row = i
			
			'If ClsFlg <> "Y" Then
				frm1.vspdData.Col = C_Check
				frm1.vspdData.value = 1
				Call vspdData_ButtonClicked(C_Check, i, 1)
			'End If
		Next	
		
	End If
End Sub

'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : 출고처리취소 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnDisSelect_OnClick()
	Dim i,ClsFlg_1
	
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Row = i
			
			'frm1.vspdData.Col = C_CloseFlg
			'ClsFlg_1 = frm1.vspdData.text
			
			'If ClsFlg_1 <> "Y" Then
				frm1.vspdData.Col = C_Check
				frm1.vspdData.value = 0
				Call vspdData_ButtonClicked(C_Check, i, 0)
			'End If
		Next	
	End If
End Sub

'==========================================================================================
'   Event Name : txtFrDt
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")  
        frm1.txtFrDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : txtToDt
'==========================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")  
        frm1.txtToDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : txtStartDt
'==========================================================================================
Sub txtStartDt_DblClick(Button)
	if Button = 1 then
		frm1.txtStartDt.Action = 7
		Call SetFocusToDocument("M")  
        frm1.txtStartDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : txtEndDt
'==========================================================================================
Sub txtEndDt_DblClick(Button)
	if Button = 1 then
		frm1.txtEndDt.Action = 7
		Call SetFocusToDocument("M")  
        frm1.txtEndDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'==========================================================================================
Sub txtFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
Sub txtStartDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtEndDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Function rdoCfmAll_OnClick()
	frm1.txtCfmFlag.value = frm1.rdoCfmAll.value
End Function

Function rdoCfmYes_OnClick()
	frm1.txtCfmFlag.value = frm1.rdoCfmYes.value
End Function

Function rdoCfmNo_OnClick()
	frm1.txtCfmFlag.value = frm1.rdoCfmNo.value
End Function


Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Col = C_Check And Row > 0 Then
	    Select Case ButtonDown
	    Case 1
			ggoSpread.Source = frm1.vspdData
			ggoSpread.UpdateRow Row
			lgBlnFlgChgValue = True		
	    Case 0
			ggoSpread.Source = frm1.vspdData
			frm1.vspdData.Col = 0
			frm1.vspdData.Row = Row 
			frm1.vspdData.text = "" 
			lgBlnFlgChgValue = False					
	    End Select
	End If
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	If frm1.vspdData.MaxRows > 0 Then
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("0000111111")
	End If   
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
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
'   Event Name : vspdDatchange
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    
'    ggoSpread.Source = frm1.vspdData
    
'	Frm1.vspdData.Row = Row
'	Frm1.vspdData.Col = Col
	
	'If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
	'	If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
	'		Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
	'	End If
	'End If
	
'	Call CheckMinNumSpread(frm1.vspdData, Col, Row) 
	
'	frm1.vspdData.Row = Row
'	frm1.vspdData.Col = 0
	
'	if Col = C_Check And ggoSpread.UpdateFlag = frm1.vspdData.Text then
'		ggoSpread.EditUndo
'	elseif Col = C_Check And ggoSpread.UpdateFlag <> frm1.vspdData.Text then
'		ggoSpread.UpdateRow Row
'	elseif Col <> C_Check then
'		ggoSpread.UpdateRow Row
'	End if
			
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
     
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then	
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
    End if
    
End Sub


'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                        
    
    Err.Clear                                               
	
	ggoSpread.Source = frm1.vspdData
	
    If ggoSpread.SSCheckChange = true Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")					
    Call InitVariables
    														
    If Not chkField(Document, "1") Then						
       Exit Function
    End If
    
	with frm1
		if (UniConvDateToYYYYMMDD(.txtFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDt.text,Parent.gDateFormat,"")) and Trim(.txtFrDt.text)<>"" and Trim(.txtToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")			
			Exit Function
		End if   
	End with
	
    If DbQuery = False Then Exit Function
       
    FncQuery = True											
    
End Function

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    Call ggoOper.ClearField(Document, "A")                  
    Call ggoOper.LockField(Document, "N")                   
    Call InitVariables                                      
    Call SetDefaultVal
    
    FncNew = True                                                           

End Function

'========================================================================================
' Function Name : FncDelete
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                                       
    
    Err.Clear                                                               
    'On Error Resume Next                                                   
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                                      
        Call DisplayMsgBox("900002", "X", "X", "X")                         
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    Call ggoOper.ClearField(Document, "A")                                  
    
    FncDelete = True                                                        
    
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear    
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                                                           
    
  
	ggoSpread.Source = frm1.vspdData                         
    If ggoSpread.SSCheckChange = False Then                  
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")    
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                         
    If Not ggoSpread.SSDefaultCheck  Then              
       Exit Function
    End If
    
    If DbSave = False Then Exit Function
    
    FncSave = True                                     
    
End Function

'========================================================================================
' Function Name : FncCopy
'========================================================================================
Function FncCopy() 
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow
End Function

'========================================================================================
' Function Name : FncCancel
'========================================================================================
Function FncCancel()
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                 
End Function

'========================================================================================
' Function Name : FncInsertRow
'========================================================================================
Function FncInsertRow() 
End Function

'========================================================================================
' Function Name : FncDeleteRow
'========================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
	if frm1.vspdData.Maxrows < 1	then exit function
    
    With frm1.vspdData 
    
    .focus
    ggoSpread.Source = frm1.vspdData 
    
	lDelRows = ggoSpread.DeleteRow

    lgBlnFlgChgValue = True
    
    End With
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint()
	ggoSpread.Source = frm1.vspdData 
	Call parent.FncPrint()
End Function
'========================================================================================
' Function Name : FncPrev
'========================================================================================
Function FncPrev() 
    On Error Resume Next                               
End Function
'========================================================================================
' Function Name : FncNext
'========================================================================================
Function FncNext() 
    On Error Resume Next                               
End Function
'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncExport(Parent.C_MULTI)						
End Function
'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData
    Call parent.FncFind(Parent.C_MULTI , False)                
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")              
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
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
    
    Err.Clear                                         

	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtSupplier=" & Trim(.hdnSupplier.value)
	    strVal = strVal & "&txtGroup=" & Trim(.hdnGroup.value)
	    strVal = strVal & "&txtPoNo=" & Trim(.hdnPoNo.value)
		strVal = strVal & "&txtFrDt=" & Trim(.hdnFrDt.value)
		strVal = strVal & "&txtToDt=" & Trim(.hdnToDt.value)
		strVal = strVal & "&txtStartDt=" & Trim(.hdnStartDt.value)
		strVal = strVal & "&txtEndDt=" & Trim(.hdnEndDt.value)
		strVal = strVal & "&txtClsFlag=" & Trim(.txtHCfmFlag.value)
		strVal = strVal & "&txtPlantCd=" & Trim(.hdnPlantCd.value)
		strVal = strVal & "&txtPotypeCd=" & Trim(.hdnPoTypeCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.hdnItemCd.value)
	Else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtSupplier=" & Trim(.txtSuppliercd.value)
	    strVal = strVal & "&txtGroup=" & Trim(.txtPur_Grp.value)
	    strVal = strVal & "&txtPoNo=" & Trim(.txtPoNo.value)
		strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)
		strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)
		strVal = strVal & "&txtStartDt=" & Trim(.txtStartDt.text)
		strVal = strVal & "&txtEndDt=" & Trim(.txtEndDt.text	)
		strVal = strVal & "&txtClsFlag=" & Trim(.txtCfmFlag.value)
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&txtPotypeCd=" & Trim(.txtPotypeCd.value)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
	End If 
	
	Call RunMyBizASP(MyBizASP, strVal)				
        
    End With
    
    DbQuery = True

End Function
'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()									
	
    lgIntFlgMode = Parent.OPMD_UMODE							
    
    Call ggoOper.LockField(Document, "Q")				
	Call SetSpreadLock
	
	frm1.btnSelect.disabled = False
	frm1.btnDisSelect.disabled = False
	
	Call SetToolbar("11101000000111")

End Function
'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	Dim ColSep, RowSep
	
    DbSave = False    
    
    If LayerShowHide(1) = False Then Exit Function
    

	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		
    
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    
    strVal = ""
    
    '-----------------------
    'Data manipulate area
    '-----------------------

	For lRow = 1 To .vspdData.MaxRows
		
        .vspdData.Row = lRow
		.vspdData.Col = 0
		
        if .vspdData.Text = ggoSpread.UpdateFlag then
	   				
			'strVal = strVal & "U" & Parent.gColSep
			.vspdData.Col = C_CloseFlg
			if Trim(.vspdData.Text) = "Y" then
				strVal = strVal & "Y" & Parent.gColSep	
				.vspdData.Col = C_PoNo
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				.vspdData.Col = C_PoSeq
				strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				strVal = strVal & lRow & Parent.gRowSep	
			else
				strDel = strDel & "N" & Parent.gColSep	
				.vspdData.Col = C_PoNo
				strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
				.vspdData.Col = C_PoSeq
				strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
				strDel = strDel & lRow & Parent.gRowSep	
			end if
			
			lGrpCnt = lGrpCnt + 1
		end if
	Next

	.txtMaxRows.value = lGrpCnt-1
	.txtSpread.value = strDel + strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)					
	
	End With
	
    DbSave = True                                       
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
'========================================================================================
Function DbSaveOk()										
	Call InitVariables()
	Call MainQuery()
	Call SetSpreadLock()
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>발주마감</font></td>
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
					<FIELDSET CLASS="CLSFLD"><TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>구매그룹</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT ALT="구매그룹"  NAME="txtPur_Grp" SIZE=10 MAXLENGTH=4 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPur_Grp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrp()">
               			 					 <INPUT TYPE=TEXT ALT="구매그룹명" NAME="txtPur_Grp_Nm" SIZE=20 MAXLENGH=20 tag="14XXXU"></TD>
						<TD CLASS="TD5" NOWRAP>발주번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="발주번호"  NAME="txtPoNo" SIZE=26 MAXLENGTH=18 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>공급처</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="공급처"  NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
           									   <INPUT TYPE=TEXT AlT="공급처" ID="txtSupplierNm" NAME="arrCond" tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>발주일</TD>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주일 NAME="txtFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주일 NAME="txtToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
								<tr>
							</table>
						</TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>발주형태</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="발주형태" NAME="txtPotypeCd"  MAXLENGTH=5 SIZE=10 tag="11NXXU" STYLE = "text-transform:uppercase"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPotype()">
											   <INPUT TYPE=TEXT AlT="발주형태" NAME="txtPotypeNm" SIZE=20 tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
											   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=25 tag="14X"></TD>
									
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE="TEXT" NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>	
						<TD CLASS=TD5 NOWRAP>납기일</TD>
						<TD CLASS=TD6 NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=납기일 NAME="txtStartDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발주일 NAME="txtEndDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 style="HEIGHT: 20px; WIDTH: 100px" tag="11X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
								<tr>
							</table>
						</TD>
					</TR>	
					<TR>
						<TD CLASS=TD5 NOWRAP>진행단계</TD>
						<TD CLASS=TD6 NOWRAP>
							<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoCfmAll" value="A" tag = "11X" checked>
								<label for="rdoCfmAll">전체</label>&nbsp;&nbsp;
							<input type=radio CLASS="RADIO" name="rdoStatusflag" id="rdoCfmYes" value="Y" tag = "11X">
								<label for="rdoCfmYes">마감</label>&nbsp;&nbsp;
							<input type=radio CLASS = "RADIO" name="rdoStatusflag" id="rdoCfmNo" value="N" tag = "11X">
								<label for="rdoCfmNo">미마감</label>
						</TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>		
					</TR>			
					</TABLE></FIELDSET></TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% valign=top><TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%">
						    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
						</TD>
					</TR></TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
	    <td WIDTH="100%">
	    	<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td WIDTH="*" align="left">
					    <button name="btnSelect" class="clsmbtn">일괄선택</button>&nbsp;
					    <BUTTON NAME="btnDisSelect" CLASS="CLSMBTN">일괄선택취소</BUTTON>
					</td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
	    	</table>
	    </td>
    </tr>
<!--
    <tr HEIGHT="20">
	    <td WIDTH="100%">
	    	<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td WIDTH="*" align="left"><button name="btnAutoSel" class="clsmbtn" ONCLICK="Selection()">일괄선택</button></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
	    	</table>
	    </td>
    </tr>
-->    
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCfmFlag" tag="14" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtHCfmFlag" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnStartDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnEndDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSupplier" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoTypeCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
</FORM>

    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>

</BODY>
</HTML>
