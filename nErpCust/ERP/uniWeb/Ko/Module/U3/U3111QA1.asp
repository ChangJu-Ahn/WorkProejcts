<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : U3111QA1.asp
'*  4. Program Name         : ��ǰ������Ȳ - Query Delivery Summary
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004-07-25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : NHG
'* 10. Modifier (Last)      : 
'* 11. Comment              : 
'* 12. History              : 
'********************************************************************************************** -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'########################################################################################################## -->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* -->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'================================================================================================================================
Const BIZ_PGM_QRY1_ID	= "U3111QB1.asp"							'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_QRY2_ID	= "U3111QB2.asp"							'��: �����Ͻ� ���� ASP�� 

'================================================================================================================================
' Grid 1(vspdData1) - Order
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_TrackingNo
Dim C_PlantCd
Dim C_PlantNm
Dim C_MvmtUnit
Dim C_RcptQty
Dim C_RcptAmt
Dim C_RetQty
Dim C_RetAmt
Dim C_TotQty
Dim C_TotAmt
Dim C_GrpFlag

' Grid 2(vspdData2) - Result
Dim C_PONo
Dim C_POSeq
Dim C_ItemCd2
Dim C_ItemNm2
Dim C_Spec2
Dim C_TrackingNo2
Dim C_DvryDt
Dim C_SLCD
Dim C_SLNM
Dim C_DvryQty
Dim C_DvryPrice
Dim C_DvryAmt

'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================
Dim IsOpenPop 
Dim lgStrPrevKey1
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow
Dim lgSortKey1
Dim lgSortKey2

Dim strDate
Dim iDBSYSDate
Dim lgStrColorFlag
Dim lgBPCD

'================================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgStrPrevKey = ""
    lgStrPrevKey1 = ""
    lgLngCurRows = 0
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey1 = 1
    lgSortKey2 = 1
End Sub

'================================================================================================================================
Sub SetDefaultVal()
	Dim strDate
	Dim BaseDate
	Dim strYear
	Dim strMonth
	Dim strDay

	BaseDate = "<%=GetSvrDate%>"

	Call ExtractDateFrom(BaseDate, parent.gServerDateFormat, parent.gServerDateType, strYear, StrMonth, StrDay)
	strDate = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	
	frm1.txtDvFrDt.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	frm1.txtDvToDt.text   = strDate
	Call SetBPCD()
End Sub

Sub SetBPCD()

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
		Call ggoOper.SetReqAttr(frm1.txtPlantCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtItemCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtDvToDt,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoAppflg1,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoAppflg2,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoAppflg3,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoGbnflg1,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoGbnflg2,"Q")
		Call ggoOper.SetReqAttr(frm1.rdoGbnflg3,"Q")
		Call ggoOper.SetReqAttr(frm1.txtSlCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTrackingNo,"Q")
		
		frm1.rdoAppflg1.checked = False
		frm1.rdoGbnflg1.checked = False
		
		Call DisplayMsgBox("210033","X","X","X")
		Call SetToolBar("10000000000011")
		Exit Sub
	Else
	    Call SetToolBar("11000000000011")										'��: ��ư ���� ���� 
	End If

	lgF0 = Split(lgF0, Chr(11))
	frm1.txtBpCd.value = parent.gUsrId
	frm1.txtBpNm.value = lgF0(1)

End Sub

'================================================================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","MA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "MA") %>
End Sub

'================================================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData1 
			
			ggoSpread.Source = frm1.vspdData1
			ggoSpread.Spreadinit "V20021224", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_GrpFlag + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit 	C_ItemCd,        "ǰ��"			,18
			ggoSpread.SSSetEdit 	C_ItemNm,        "ǰ���"		,20
			ggoSpread.SSSetEdit 	C_Spec,			 "�԰�"			,20
			ggoSpread.SSSetEdit 	C_Trackingno,	 "Tracking No."	,20
			ggoSpread.SSSetEdit 	C_PlantCd,		 "��ǰó"		,6
			ggoSpread.SSSetEdit 	C_PlantNm,       "��ǰó��"		,10
			ggoSpread.SSSetEdit 	C_MvmtUnit,		 "����"			,4
			ggoSpread.SSSetFloat 	C_RcptQty,		 "��ǰ����"		,12,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_RcptAmt,		 "��ǰ�ݾ�"		,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_RetQty,		 "��ǰ����"		,12,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_RetAmt,		 "��ǰ�ݾ�"		,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_TotQty,		 "�԰��ѷ�"		,12,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_TotAmt,		 "�԰��ѱݾ�"	,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_GrpFlag,       ""				,1
			
			Call ggoSpread.SSSetColHidden( C_GrpFlag, C_GrpFlag, True)
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("A")
			
			.Col = 1 : .ColMerge = 2
			.Col = 2 : .ColMerge = 2
			.Col = 3 : .ColMerge = 2
			.Col = 4 : .ColMerge = 2
			.Col = 5 : .ColMerge = 2
			
			.ReDraw = true    
    
		End With
	
    End If
    
    If pvSpdNo = "B" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData2 
			
			ggoSpread.Source = frm1.vspdData2
			ggoSpread.Spreadinit "V20021225", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_DvryAmt + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit 	C_PONo,			"���ֹ�ȣ"		,15
			ggoSpread.SSSetEdit		C_POSeq,		"���"			,4
			ggoSpread.SSSetEdit 	C_ItemCd2,      "ǰ��"			,15
			ggoSpread.SSSetEdit 	C_ItemNm2,      "ǰ���"		,20
			ggoSpread.SSSetEdit 	C_Spec2,		"�԰�"			,20
			ggoSpread.SSSetEdit 	C_Trackingno2,  "Tracking No."	,20
			ggoSpread.SSSetDate 	C_DvryDt,		"��ǰ��"		,10, 2, parent.gDateFormat		 
			ggoSpread.SSSetEdit 	C_SLCD,			"��ǰâ��"		,12
			ggoSpread.SSSetEdit 	C_SLNM,			"��ǰâ���"	,18			
			ggoSpread.SSSetFloat	C_DvryQty,		"��ǰ����"		,12,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_DvryPrice,	"��ǰ�ܰ�"		,12,parent.ggUnitCostNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_DvryAmt,		"��ǰ�ݾ�"		,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("B")
			
			.ReDraw = true    
    
		End With
    End If
    
End Sub

'================================================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)
	If pvSpdNo = "A" Then
		'--------------------------------
		'Grid 1
		'--------------------------------
		ggoSpread.Source = frm1.vspdData1
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
		
	If pvSpdNo = "B" Then 
		'--------------------------------
		'Grid 2
		'--------------------------------
		ggoSpread.Source = frm1.vspdData2
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If	
End Sub

'================================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'================================================================================================================================
Sub InitComboBox()

End Sub

'================================================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData1)
		C_ItemCd		= 1
		C_ItemNm		= 2
		C_Spec			= 3
		C_Trackingno	= 4
		C_PlantCd		= 5
		C_PlantNm		= 6
		C_MvmtUnit		= 7
		C_RcptQty		= 8
		C_RcptAmt		= 9
		C_RetQty		= 10
		C_RetAmt		= 11
		C_TotQty		= 12
		C_TotAmt		= 13
		C_GrpFlag		= 14
	End If	
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2)
		C_PONo				= 1
		C_POSeq				= 2
		C_ItemCd2			= 3
		C_ItemNm2			= 4
		C_Spec2				= 5
		C_Trackingno2		= 6
		C_DvryDt			= 7
		C_SLCD				= 8
		C_SLNM				= 9
		C_DvryQty			= 10
		C_DvryPrice			= 11
		C_DvryAmt			= 12
		
	End If	

End Sub

'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
      
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData1
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
			C_Trackingno	= iCurColumnPos(4)
			C_PlantCd		= iCurColumnPos(5)
			C_PlantNm		= iCurColumnPos(6)
			C_MvmtUnit		= iCurColumnPos(7)
			C_RcptQty		= iCurColumnPos(8)
			C_RcptAmt		= iCurColumnPos(9)
			C_RetQty		= iCurColumnPos(10)
			C_RetAmt		= iCurColumnPos(11)
			C_TotQty		= iCurColumnPos(12)
			C_TotAmt		= iCurColumnPos(13)
			C_GrpFlag		= iCurColumnPos(14)
			
		Case "B"
		
			ggoSpread.Source = frm1.vspdData2
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_PONo				= iCurColumnPos(1)
			C_POSeq				= iCurColumnPos(2)
			C_ItemCd2			= iCurColumnPos(3)
			C_ItemNm2			= iCurColumnPos(4)
			C_Spec2				= iCurColumnPos(5)
			C_Trackingno2		= iCurColumnPos(6)
			C_DvryDt			= iCurColumnPos(7)
			C_SLCD				= iCurColumnPos(8)
			C_SLNM				= iCurColumnPos(9)
			C_DvryQty			= iCurColumnPos(10)
			C_DvryPrice			= iCurColumnPos(11)
			C_DvryAmt			= iCurColumnPos(12)
			
    End Select

End Sub    

'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��ǰó�˾�"
	arrParam(1) = "(			SELECT	DISTINCT B.PLANT_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_DTL B, M_PUR_ORD_HDR C "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO AND A.SPLIT_SEQ_NO = 0 "
	arrParam(1) = arrParam(1) & "AND	A.PO_NO = C.PO_NO AND C.BP_CD = '" & frm1.txtBpCd.value & "') A, B_PLANT B"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD"			
	arrParam(5) = "��ǰó"			
	
    arrField(0) = "A.PLANT_CD"	
    arrField(1) = "B.PLANT_NM"	
    
    arrHeader(0) = "��ǰó"		
    arrHeader(1) = "��ǰó��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'================================================================================================================================
Function OpenItemInfo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ���˾�"
	arrParam(1) = "(			SELECT	DISTINCT ITEM_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_HDR B "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.SPLIT_SEQ_NO = 0 AND B.BP_CD = '" & frm1.txtBpCd.value & "') A, B_ITEM B"
	arrParam(2) = Trim(frm1.txtItemCd.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = "A.ITEM_CD = B.ITEM_CD "
	arrParam(5) = "ǰ��"
	 
    arrField(0) = "A.ITEM_CD"												' Field��(0)
    arrField(1) = "B.ITEM_NM"												' Field��(1)
    
    arrHeader(0) = "ǰ��"													' Header��(0)
    arrHeader(1) = "ǰ���"													' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetItemInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtItemCd.focus

End Function

'================================================================================================================================
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus()		
End Function

'================================================================================================================================
Function SetItemInfo(Byval arrRet)
	frm1.txtItemCd.value = arrRet(0)
	frm1.txtItemNm.value = arrRet(1)
	frm1.txtItemCd.focus()
End Function


'================================================================================================================================
Function OpenSlInfo(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemCd.className) = "PROTECTED" Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "â���˾�"
	arrParam(1) = "B_STORAGE_LOCATION "
	arrParam(2) = Trim(frm1.txtSlCd.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = ""
	arrParam(5) = "â��"
	 
    arrField(0) = "SL_CD"												' Field��(0)
    arrField(1) = "SL_NM"												' Field��(1)
    
    arrHeader(0) = "â��"													' Header��(0)
    arrHeader(1) = "â���"													' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetSlInfo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtSlCd.focus

End Function

'================================================================================================================================
Function SetSlInfo(Byval arrRet)
    With frm1
		.txtSlCd.value = arrRet(0)
		.txtSlNm.value = arrRet(1)
    End With
End Function

'--------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo()
	
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	arrParam(3) = frm1.txtDvFrDt.Text
	arrParam(4) = frm1.txtDvToDt.Text
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetTrackingNo(arrRet)
	End If
	
	Call SetFocusToDocument("M")
	frm1.txtTrackingNo.focus
	
End Function

'------------------------------------------  SetTrackingNo()  -----------------------------------------
'	Name : SetTrackingNo()
'	Description : Tracking No. Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetTrackingNo(Byval arrRet)
    frm1.txtTrackingNo.Value = arrRet(0)
End Function

Function CookiePage(ByVal Kubun)

End Function

'================================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    Call LockObjectField(frm1.txtDvFrDt,"R")
    Call LockObjectField(frm1.txtDvToDt,"R")
    Call FormatDATEField(frm1.txtDvFrDt)
    Call FormatDATEField(frm1.txtDvToDt)
    
    Call InitSpreadSheet("*")
   
    Call SetDefaultVal
    Call InitVariables
    Call InitComboBox
 
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement

End Sub

'================================================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'================================================================================================================================
Sub txtDvFrDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtDvFrDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtDvFrDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtDvToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtDvToDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtDvToDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtDvFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub txtDvToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub vspdData1_GotFocus()
    ggoSpread.Source = frm1.vspdData1
End Sub

'================================================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey <> "" Then
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'================================================================================================================================
Sub vspdData1_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData1

    If frm1.vspdData1.MaxRows = 0 Then
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey1 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey1 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey1
            lgSortKey1 = 1
        End If
   
    End If
    
    If lgOldRow <> Row Then
				
		frm1.vspdData2.MaxRows = 0 
		lgStrPrevKey1 = ""
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Sub
		End If
		
		lgOldRow = frm1.vspdData1.ActiveRow
			
	End If
    
End Sub

'================================================================================================================================
Sub vspdData2_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SP2C"
	
	Set gActiveSpdSheet = frm1.vspdData2

    If frm1.vspdData2.MaxRows = 0 Then
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey2 = 1 Then
            ggoSpread.SSSort Col
            lgSortKey2 = 2
        Else
            ggoSpread.SSSort Col, lgSortKey2
            lgSortKey2 = 1
        End If
    Else
        
    End If
    
End Sub

'================================================================================================================================
Sub vspdData1_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'================================================================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SP2C" Then
       gMouseClickStatus = "SP2CR"
    End If

End Sub

'================================================================================================================================
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'================================================================================================================================
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'================================================================================================================================
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'================================================================================================================================
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub

'================================================================================================================================
Sub vspdData1_Change(ByVal Col , ByVal Row )

End Sub

'================================================================================================================================
Sub vspdData2_Change(ByVal Col , ByVal Row )

End Sub

'================================================================================================================================
Sub vspdData1_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData1 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub

'================================================================================================================================
Sub vspdData2_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData2 
		If Row >= NewRow Then
		    Exit Sub
		End If
    End With
End Sub

'================================================================================================================================
Function FncQuery() 
    Dim IntRetCD 

    FncQuery = False
    Err.Clear

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = "" 
	End If	
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = "" 
	End If

	If ValidDateCheck(frm1.txtDvFrDt, frm1.txtDvToDt) = False Then Exit Function
		
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.ClearSpreadData
    ggoSpread.Source = frm1.vspdData2
    ggoSpread.ClearSpreadData
    Call InitVariables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then
		Call RestoreToolBar()
		Exit Function														'��: Query db data
	End If
	
    FncQuery = True															'��: Processing is OK
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
	On Error Resume Next
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncSave() 
	On Error Resume Next    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
	On Error Resume Next	
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next													'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next													'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)									'��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)								'��: Protect system from crashing
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
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'******************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  **************************
'	���� : 
'**************************************************************************************** 

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)

	Dim strPlantcd
	Dim strItemCd
	Dim dtMvmtDt

   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   Select Case pOpt
       Case "M"
       
				With frm1
					If lgIntFlgMode = parent.OPMD_UMODE Then
						lgKeyStream = UCase(Trim(.hPlantCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.hItemCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.hBPCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hDvryFromDt.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hDvryToDt.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hrdoAppflg.value)  & Parent.gColSep
						If .rdogbnflg1.checked = true Then
							lgKeyStream = lgKeyStream & "A" & Parent.gColSep
						ElseIf .rdoGbnflg2.checked = true Then
							lgKeyStream = lgKeyStream & "N" & Parent.gColSep
						Else
							lgKeyStream = lgKeyStream & "Y" & Parent.gColSep
						End If						
						lgKeyStream = lgKeyStream & Trim(.hSlCd.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hTrackingNo.value)  & Parent.gColSep
					Else
						lgKeyStream = UCase(Trim(.txtPlantCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtItemCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtBPCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtDvFrDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtDvToDt.Text)  & Parent.gColSep
						If .rdoAppflg(0).checked = true Then
							lgKeyStream = lgKeyStream & "A" & Parent.gColSep
						ElseIf .rdoAppflg(1).checked = true Then
							lgKeyStream = lgKeyStream & "Y" & Parent.gColSep
						Else
							lgKeyStream = lgKeyStream & "N" & Parent.gColSep
						End If
						
						If .rdogbnflg1.checked = true Then
							lgKeyStream = lgKeyStream & "A" & Parent.gColSep
						ElseIf .rdoGbnflg2.checked = true Then
							lgKeyStream = lgKeyStream & "N" & Parent.gColSep
						Else
							lgKeyStream = lgKeyStream & "Y" & Parent.gColSep
						End If
						
						If .rdoAppflg(0).checked = true Then
							.hrdoAppflg.value = "A"
						ElseIf .rdoAppflg(1).checked = true Then
							.hrdoAppflg.value = "Y"
						Else
							.hrdoAppflg.value = "N"
						End If
						
						lgKeyStream = lgKeyStream & Trim(.txtSlCd.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtTrackingNo.value)  & Parent.gColSep
						
						.hPlantCd.value		= .txtPlantCd.value
						.hItemCd.value		= .txtItemCd.value
						.hBPCd.value		= .txtBPCd.value
						.hDvryFromDt.value	= .txtDvFrDt.Text
						.hDvryToDt.value	= .txtDvToDt.Text
						.hSlCd.value		= .txtSlCd.value
						.hBPCd.value		= .txtBPCd.value
						
					End If
				End With
			
       Case "S"
				With frm1
					.vspdData1.Row = .vspdData1.ActiveRow
					.vspdData1.Col = C_PlantCd
					strPlantcd = .vspdData1.text
					If strPLANTcd = "" Then
						strPLANTcd = UCase(Trim(.hPLANTCd.value))
					End If
					
					.vspdData1.Col = C_ItemCd
					strItemCd = .vspdData1.text
					If strItemCd = "" Then
						strItemCd = UCase(Trim(.hItemCd.value))
					End If
					
					'.vspdData1.Col = C_MvmtDt
					'dtMvmtDt = .vspdData1.text
					
					lgKeyStream = UCase(Trim(strPlantcd))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(strItemCd))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(.hBPCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hDvryFromDt.value)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hDvryToDt.value)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hrdoAppflg.value)  & Parent.gColSep
					lgKeyStream = lgKeyStream & Trim(.hSlCd.value)  & Parent.gColSep
					
					If .rdogbnflg1.checked = true Then
						lgKeyStream = lgKeyStream & "A" & Parent.gColSep
					ElseIf .rdoGbnflg2.checked = true Then
						lgKeyStream = lgKeyStream & "N" & Parent.gColSep
					Else
						lgKeyStream = lgKeyStream & "Y" & Parent.gColSep
					End If
					
					lgKeyStream = lgKeyStream & Trim(.hTrackingNo.value)  & Parent.gColSep
					
				End With

	End Select
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
   
End Sub    

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 

	Dim strVal

    DbQuery = False

	Call LayerShowHide(1)
    
    Call MakeKeyStream("M")
    
	strVal = BIZ_PGM_QRY1_ID & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey
	strVal = strVal & "&txtMaxRows="	& frm1.vspddata1.MaxRows
    
    Call RunMyBizASP(MyBizASP, strVal)
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()

	Call SetToolBar("11000000000111")														'��: ��ư ���� ���� 
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData1,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If

	Call SetQuerySpreadColor
	
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	lgOldRow = 1
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		If DbDtlQuery = False Then	
			Call RestoreToolBar()
			Exit Function
		End If
	End If
	
	lgIntFlgMode = parent.OPMD_UMODE														'��: Indicates that current mode is Update mode
			
End Function

Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)

	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)
		
		With frm1.vspdData1	
		.Col = -1
		.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				'.BackColor = RGB(204,255,153) '���� 
			Case "2"
				.BackColor = RGB(176,234,244) '�ϴû� 
				.ForeColor = vbBlue
			Case "3"
				.BackColor = RGB(224,206,244) '������ 
			Case "4"  
				.BackColor = RGB(251,226,153) '����Ȳ 
			Case "5" 
				.BackColor = RGB(255,255,153) '����� 
				.ForeColor = vbRed
		End Select
		End With
	Next

End Sub

'========================================================================================
' Function Name : DbDtlQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbDtlQuery() 
    
    Dim strVal
	Dim strPlantcd
	Dim strItemCd
	Dim dtMvmtDt
	
    DbDtlQuery = False

	Call LayerShowHide(1)
    
    Call MakeKeyStream("S")
    
	strVal = BIZ_PGM_QRY2_ID & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="     & lgKeyStream
 
    Call RunMyBizASP(MyBizASP, strVal)													'��: �����Ͻ� ASP �� ���� 
    
    DbDtlQuery = True
    
    
End Function

'========================================================================================
Function DbDtlQueryOk()

End Function

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
' Function Desc : �׸��� �����¸� �����Ѵ�.
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Function Desc : �׸��带 ���� ���·� �����Ѵ�.
'========================================================================================
Sub PopRestoreSpreadColumnInf()
	Dim LngRow

    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet(gActiveSpdSheet.Id)
	Call ggoSpread.ReOrderingSpreadData()
	
End Sub 


'================================================================================================================================
Function BtnPreview()
    
    Dim strEbrFile
    Dim objName
    
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	
	dim strUrl
	dim arrParam, arrField, arrHeader

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	Call BtnDisabled(1)
	
	If frm1.hBPCd.value = "" Then
		var1 = "%"
	Else
		var1 = Trim(frm1.hBPCd.value)
	End If
	
	If frm1.rdoAppflg(0).checked = true Then
		var2 = "%"
	ElseIf frm1.rdoAppflg(1).checked = true Then
		var2 = "Y"
	Else
		var2 = "N"
	End If
	
	If frm1.hDvryFromDt.value = "" Then
		var3 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var3 = UniConvDateAtoB(frm1.hDvryFromDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
	If frm1.hDvryToDt.value = "" Then
		var4 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var4 = UniConvDateAtoB(frm1.hDvryToDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
	
		var5 = UNIGetLastDay(frm1.hDvryToDt.value,parent.gServerDateFormat)
	
	strUrl =  "BP|" & var1
	strUrl = strUrl & "|TYPE|" & var2
	strUrl = strUrl & "|FRDT|" & var3
	strUrl = strUrl & "|TODT|" & var4
	strUrl = strUrl & "|DT|"   & var5



	strEbrFile = "U3111QA1"
	objName = AskEBDocumentName(strEbrFile,"ebr")

	call FncEBRPreview(objName, strUrl)
	
	Call BtnDisabled(0)
	
	frm1.btnRun(0).focus
	Set gActiveElement = document.activeElement
	
End Function

'================================================================================================================================
Function BtnPrint()
	
	Dim strEbrFile
    Dim objName
    
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	
	dim strUrl
	dim arrParam, arrField, arrHeader

	If lgIntFlgMode = Parent.OPMD_CMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End If

	Call BtnDisabled(1)
	
	If frm1.hBPCd.value = "" Then
		var1 = "%"
	Else
		var1 = Trim(frm1.hBPCd.value)
	End If
	
	If frm1.rdoAppflg(0).checked = true Then
		var2 = "%"
	ElseIf frm1.rdoAppflg(1).checked = true Then
		var2 = "Y"
	Else
		var2 = "N"
	End If
	
	
	If frm1.hDvryFromDt.value = "" Then
		var3 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var3 = UniConvDateAtoB(frm1.hDvryFromDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
	If frm1.hDvryToDt.value = "" Then
		var4 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var4 = UniConvDateAtoB(frm1.hDvryToDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
		var5 = UNIGetLastDay(frm1.hDvryToDt.value,parent.gServerDateFormat)
	strUrl =  "BP|" & var1
	strUrl = strUrl & "|TYPE|" & var2
	strUrl = strUrl & "|FRDT|" & var3
	strUrl = strUrl & "|TODT|" & var4
	strUrl = strUrl & "|DT|"   & var5
	
	strEbrFile = "U3111QA1"
	objName = AskEBDocumentName(strEbrFile,"ebr")
	
	call FncEBRprint(EBAction, objName, strUrl)
	
	Call BtnDisabled(0)	
	
	frm1.btnRun(1).focus
	Set gActiveElement = document.activeElement

End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
    '----------  Coding part  -------------------------------------------------------------


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ǰ������Ȳ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=500>&nbsp;</TD>
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
			 						<TD CLASS=TD5 NOWRAP>��ü</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="14" ALT="��ü">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=25 tag="14" ALT="��ü��"></TD>
			 						<TD CLASS=TD5 NOWRAP>��ǰó</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="11xxxU" ALT="��ǰó"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="��ǰó��"></TD>
								</TR>
			 					<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>��ǰ��</TD> 
									<TD CLASS=TD6>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtDvFrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="��ǰ������"></OBJECT>');</SCRIPT>
										&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtDvToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="��ǰ������"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
			 					<TR>
									<TD CLASS="TD5" NOWRAP>����/��ǰ����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoGbnflg" id = "rdoGbnflg1" Value="A" checked tag="11"><label for="rdoGbnflg1">&nbsp;��ü&nbsp;</label>
														   <INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoGbnflg" id = "rdoGbnflg2" Value="N" tag="11"><label for="rdoGbnflg2">&nbsp;����&nbsp;</label>
														   <INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoGbnflg" id = "rdoGbnflg3" Value="Y" tag="11"><label for="rdoGbnflg3">&nbsp;��ǰ&nbsp;</label></TD>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoAppflg" id = "rdoAppflg1" Value="A" checked tag="11"><label for="rdoAppflg1">&nbsp;��ü&nbsp;</label>
														   <INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoAppflg" id = "rdoAppflg2" Value="N" tag="11"><label for="rdoAppflg2">&nbsp;���&nbsp;</label>
														   <INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoAppflg" id = "rdoAppflg3" Value="Y" tag="11"><label for="rdoAppflg3">&nbsp;�Ϲ�&nbsp;</label></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��ǰâ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSlCd" SIZE=15 MAXLENGTH=18 tag="11xxxU" ALT="��ǰâ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSlCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSlInfo frm1.txtSlCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtSlNm" SIZE=28 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11xxxU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNoBtn" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo()"></TD>
								</TR>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> name=vspdData1 Id = "A" HEIGHT="100%" width="100%" tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> name=vspdData2 Id = "B" HEIGHT="100%" width="100%" tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	<!--<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>				
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=0>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=0>�μ�</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>-->
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hrdoAppflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hDvryFromDt" tag="24"><INPUT TYPE=HIDDEN NAME="hDvryToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hSlCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
</FORM>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<input type="hidden" name="uname">
	<input type="hidden" name="dbname">
	<input type="hidden" name="filename">
	<input type="hidden" name="condvar">
	<input type="hidden" name="date">                 
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</H    