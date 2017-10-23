<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : i1611ma1_KO391.asp
'*  4. Program Name         : ǰ�񺰼�����Ȳ��ȸ(S) - Query Goods Movement By Item
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2004-10-08
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park, BumSoo
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'================================================================================================================================
Const BIZ_PGM_QRY_ID	= "i1611mb1_KO391.asp"							'��: �����Ͻ� ���� ASP�� 

'================================================================================================================================
' Grid (vspdData)
Dim C_ItemCd			'ǰ�� 
Dim C_ItemNm			'ǰ��� 
Dim C_Spec				'�԰� 
Dim C_MoveDt			'����
Dim C_BaseQty			'���ʼ��� 
Dim C_RcptQty			'�԰���� 
Dim C_IssueQty			'������ 
Dim C_OnhandQty			'������
Dim C_BadRcptQty		'�ҷ��԰� 
Dim C_BadIssueQty		'�ҷ���� 
Dim C_BadOnhandQty		'�ҷ���� 
Dim C_TransType			'���ұ���
Dim C_MoveType			'��������
Dim C_DocumentNo		'���ҹ�ȣ
Dim	C_SlCd				'â��
Dim	C_SlNm				'â���
Dim	C_Bp_cd				'����ó
Dim	C_Bp_Nm				'����ó��
Dim	C_WcCd				'�۾���
Dim	C_WcNm				'�۾����
Dim C_TrnsSlCd			'�̵�â��
Dim C_TrnsSlNm			'�̵�â���
Dim C_DocumentText		'���Һ��
Dim	C_PoNo				'���ֹ�ȣ
Dim	C_PoRcptNo			'����������ȣ
Dim	C_ProdtOrderNo		'����������ȣ
Dim	C_DnNo				'����ȣ
Dim	C_Remark			'�����
Dim C_LotNo				'�����ȣ
Dim C_Price				'�ܰ�
Dim C_Amount			'�ݾ�
Dim	C_insrtUserId		'�����

'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================
Dim IsOpenPop 
Dim lgAfterQryFlg
Dim lgLngCnt
Dim lgOldRow

Dim strDate
Dim iDBSYSDate
Dim lgStrColorFlag
Dim lgQueryType
Dim lgOnhandQty, lgBadOnhandQty, lgRcptQty, lgBadRcptQty, lgIssueQty, lgBadIssueQty

'================================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgStrPrevKey = ""
    lgLngCurRows = 0
    lgAfterQryFlg = False
    lgLngCnt = 0
    lgOldRow = 0
    lgSortKey = 1

    lgOnhandQty = 0
    lgBadOnhandQty = 0
    lgRcptQty = 0
    lgBadRcptQty = 0
    lgIssueQty = 0
    lgBadIssueQty = 0

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
	
	frm1.txtReportFrDt.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
End Sub

'================================================================================================================================
Sub LoadInfTB19029()     
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q","I","NOCOOKIE","MA") %>
End Sub

'================================================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

	Call InitSpreadPosVariables(pvSpdNo)

	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		'------------------------------------------
		' Grid 1 - Operation Spread Setting
		'------------------------------------------
		With frm1.vspdData
			
			ggoSpread.Source = frm1.vspdData
			ggoSpread.Spreadinit "V20021224", ,Parent.gAllowDragDropSpread
					
			.ReDraw = false
					
			.MaxCols = C_insrtUserId + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit 	C_ItemCd,		"ǰ��"			,12
			ggoSpread.SSSetEdit 	C_ItemNm,		"ǰ���"		,20
			ggoSpread.SSSetEdit 	C_Spec,			"�԰�"			,25
			ggoSpread.SSSetEdit 	C_MoveDt,		"����"			,10, 2
			ggoSpread.SSSetFloat 	C_BaseQty,		"�̿�����"		,8,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat 	C_RcptQty,		"��ǰ�԰�"		,8,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat 	C_IssueQty,		"��ǰ���"		,8,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat 	C_OnhandQty,	"��ǰ���"		,8,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat 	C_BadRcptQty,	"�ҷ��԰�"		,8,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat 	C_BadIssueQty,	"�ҷ����"		,8,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetFloat 	C_BadOnhandQty,	"�ҷ����"		,8,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetEdit 	C_TransType,	"���ұ���"		,8
			ggoSpread.SSSetEdit 	C_MoveType,		"��������"		,10
			ggoSpread.SSSetEdit 	C_DocumentNo,	"���ҹ�ȣ"		,12
			ggoSpread.SSSetEdit 	C_SlCd,			"â��"			,6
			ggoSpread.SSSetEdit 	C_SlNm,			"â���"		,12
			ggoSpread.SSSetEdit 	C_Bp_Cd,		"�ŷ�ó"		,6
			ggoSpread.SSSetEdit 	C_Bp_Nm,		"�ŷ�ó��"		,10
			ggoSpread.SSSetEdit 	C_WcCd,			"�۾���"		,6
			ggoSpread.SSSetEdit 	C_WcNm,			"�۾����"		,10
			ggoSpread.SSSetEdit 	C_TrnsSlCd,		"�̵�â��"		,7
			ggoSpread.SSSetEdit 	C_TrnsSlNm,		"�̵�â���"	,12
			ggoSpread.SSSetEdit 	C_DocumentText,	"���Һ��"		,20
			ggoSpread.SSSetEdit 	C_PoNo,			"���ֹ�ȣ"		,12
			ggoSpread.SSSetEdit 	C_PoRcptNo,		"���������ȣ"		,12
			ggoSpread.SSSetEdit 	C_ProdtOrderNo,	"����������ȣ"	,12
			ggoSpread.SSSetEdit 	C_DnNo,			"����ȣ"		,12
			ggoSpread.SSSetEdit 	C_Remark,		"�����"		,20
			ggoSpread.SSSetEdit 	C_LotNo,		"LOT NO"		,12
			ggoSpread.SSSetFloat	C_Price,		"�ܰ�"			,12,parent.ggUnitCostNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_Amount,		"�ݾ�"			,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit 	C_insrtUserId,	"�����"		,10
		
			Call ggoSpread.SSSetColHidden( C_ItemCd, C_Spec, True)
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(1)
			
			Call SetSpreadLock("A")
			
			.Col = 1 : .ColMerge = 2
			.Col = 2 : .ColMerge = 2
			.Col = 3 : .ColMerge = 2
'			.Col = 4 : .ColMerge = 2

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
		ggoSpread.Source = frm1.vspdData
'		ggoSpread.SpreadLockWithOddEvenRowColor()
		frm1.vspdData.ReDraw = False
   		ggoSpread.SpreadLock	 C_ItemCd, -1, C_insrtUserId
		frm1.vspdData.ReDraw = True

	End If
End Sub

'================================================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

End Sub

'================================================================================================================================
Sub InitComboBox()

	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = 'I0003'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboSLGroup,lgF0  ,lgF1  ,Chr(11))

End Sub

'================================================================================================================================
Sub InitSpreadPosVariables(ByVal pvSpdNo)
	If pvSpdNo = "A" Or pvSpdNo = "*" Then
		' Grid 1(vspdData)
		C_ItemCd		= 1
		C_ItemNm		= 2
		C_Spec			= 3
		C_MoveDt		= 4
		C_BaseQty		= 5
		C_RcptQty		= 6
		C_IssueQty		= 7
		C_OnhandQty		= 8
		C_BadRcptQty	= 9
		C_BadIssueQty	= 10
		C_BadOnhandQty	= 11
		C_TransType		= 12
		C_MoveType		= 13
		C_DocumentNo	= 14
		C_SlCd			= 15
		C_SlNm			= 16
		C_Bp_cd			= 17
		C_Bp_Nm			= 18
		C_WcCd			= 19
		C_WcNm			= 20
		C_TrnsSlCd		= 21
		C_TrnsSlNm		= 22
		C_DocumentText	= 23
		C_PoNo			= 24
		C_PoRcptNo		= 25
		C_ProdtOrderNo	= 26
		C_DnNo			= 27
		C_Remark		= 28
		C_LotNo			= 29
		C_Price			= 30
		C_Amount		= 31
		C_insrtUserId	= 32
	End If	
	
End Sub

'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
      
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCd		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_Spec			= iCurColumnPos(3)
			C_MoveDt		= iCurColumnPos(4)
			C_BaseQty		= iCurColumnPos(5)
			C_RcptQty		= iCurColumnPos(6)
			C_IssueQty		= iCurColumnPos(7)
			C_OnhandQty		= iCurColumnPos(8)
			C_BadRcptQty	= iCurColumnPos(9)
			C_BadIssueQty	= iCurColumnPos(10)
			C_BadOnhandQty	= iCurColumnPos(11)
			C_TransType		= iCurColumnPos(12)
			C_MoveType		= iCurColumnPos(13)
			C_DocumentNo	= iCurColumnPos(14)
			C_SlCd			= iCurColumnPos(15)
			C_SlNm			= iCurColumnPos(16)
			C_Bp_cd			= iCurColumnPos(17)
			C_Bp_Nm			= iCurColumnPos(18)			
			C_WcCd			= iCurColumnPos(19)
			C_WcNm			= iCurColumnPos(20)
			C_TrnsSlCd		= iCurColumnPos(21)
			C_TrnsSlNm		= iCurColumnPos(22)
			C_DocumentText	= iCurColumnPos(23)
			C_PoNo			= iCurColumnPos(24)
			C_PoRcptNo		= iCurColumnPos(25)
			C_ProdtOrderNo	= iCurColumnPos(26)
			C_DnNo			= iCurColumnPos(27)
			C_Remark		= iCurColumnPos(28)
			C_LotNo			= iCurColumnPos(29)
			C_Price			= iCurColumnPos(30)
			C_Amount		= iCurColumnPos(31)
			C_insrtUserId	= iCurColumnPos(32)
			
    End Select

End Sub    

'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "����"		
    arrHeader(1) = "�����"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPlant(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPlantCd.focus
	
End Function

'------------------------------------------  OpenSLCd()  -------------------------------------------------
Function OpenSLCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus    
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "â���˾�"                                                                    
	arrParam(1) = "B_STORAGE_LOCATION"                           
	arrParam(2) = Trim(frm1.txtSlCd.Value)                     
	arrParam(3) = ""
	If frm1.cboSLGroup.value = "" Then
		arrParam(4) = "Plant_cd= " & FilterVar(frm1.txtPlantCd.value,"","S") 
	ELse
		arrParam(4) = "Plant_cd= " & FilterVar(frm1.txtPlantCd.value,"","S") & " and Sl_Group_Cd = " & FilterVar(frm1.cboSLGroup.value,"","S")
	End IF
	arrParam(5) = "â��"
	 
	arrField(0) = "Sl_Cd" 
	arrField(1) = "Sl_Nm" 
	 
	arrHeader(0) = "â��"  
	arrHeader(1) = "â���"  

	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	  
	IsOpenPop = False
	 
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtSlCd.Value = arrRet(0)
		frm1.txtSlNm.Value = arrRet(1)
		frm1.txtSlCd.focus
	End If 
	Set gActiveElement = document.activeElement 
End Function

'-----------------------  OpenItem()  -------------------------------------------------
Function OpenItem()
 
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value= "" Then
		Call Displaymsgbox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If
	
	iCalledAspName = AskPRAspName("B1B11PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)	' Item Code
	arrParam(2) = ""				' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""				' Default Value
	
	arrField(0) = 1 '"ITEM_CD"			' Field��(0)
	arrField(1) = 2 '"ITEM_NM"			' Field��(1)
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
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
End Function

'================================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
        
    Call InitSpreadSheet("*")   
    Call SetDefaultVal
    Call InitVariables
    Call InitComboBox
 
    Call SetToolBar("11000000000011") 
    
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtItemCd.focus
	Else
		frm1.txtPlantCd.focus 
	End If
	
	Set gActiveElement = document.activeElement

End Sub

'================================================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'================================================================================================================================
Sub txtReportFrDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtReportFrDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtReportFrDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtReportFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub txtReportToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtReportToDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtReportToDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtReportToDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then
        Exit Sub
	End If
        
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If DbQuery = False Then	
				Call RestoreToolBar()
				Exit Sub
			End If
		End If     
    End if
    
End Sub

'================================================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("0000111111")
	'----------------------
	'Column Split
	'----------------------
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then
       Exit Sub
   	End If
   	
   	If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
   
    End If
    
End Sub

'================================================================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
		
	If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'================================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'================================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'================================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

End Sub

'================================================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData
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
		
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables

    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkfield(Document, "1") Then
       Exit Function
    End If
    
    lgQueryType = "NORM"

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

    FncPrev = False                                                        '��: Processing is NG
    
    Err.Clear                                                              '��: Protect system from crashing

    lgQueryType = "PREV"

    If frm1.txtPlantCd.value = "" Then
	frm1.txtPlantNm.value = "" 
    End If	
	
    If frm1.txtItemCd.value = "" Then
	frm1.txtItemNm.value = "" 
    End If
		
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
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
      
    FncPrev = True															'��: Protect system from crashing

End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 

    FncNext = False                                                        '��: Processing is NG
    
    Err.Clear                                                              '��: Protect system from crashing

    lgQueryType = "NEXT"

    If frm1.txtPlantCd.value = "" Then
	frm1.txtPlantNm.value = "" 
    End If	
	
    If frm1.txtItemCd.value = "" Then
	frm1.txtItemNm.value = "" 
    End If
		
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")
    ggoSpread.Source = frm1.vspdData
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
      
    FncNext = True		

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

   Select Case pOpt
       Case "M"
       
				With frm1
					If lgIntFlgMode = parent.OPMD_UMODE Then
						lgKeyStream = UCase(Trim(.hPlantCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.hItemCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtSlCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hReportFrDt.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hReportToDt.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.hcboSLGroup.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim("NORM")  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgOnhandQty)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgBadOnhandQty)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgRcptQty)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgBadRcptQty)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgIssueQty)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgBadIssueQty)  & Parent.gColSep
					Else
						lgKeyStream = UCase(Trim(.txtPlantCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtItemCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtSlCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtReportFrDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtReportToDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.cboSLGroup.value)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgQueryType)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgOnhandQty)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgBadOnhandQty)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgRcptQty)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgBadRcptQty)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgIssueQty)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(lgBadIssueQty)  & Parent.gColSep

						.hPlantCd.value			= .txtPlantCd.value
						.hItemCd.value			= .txtItemCd.value
						.hSlCd.value			= .txtSlCd.value
						.hReportFrDt.value		= .txtReportFrDt.Text
						.hReportToDt.value		= .txtReportToDt.Text
						.hcboSLGroup.value		= .cboSLGroup.value

					End If
				End With
			
	End Select
   
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
    
	strVal = BIZ_PGM_QRY_ID & "?txtMode="	& parent.UID_M0001
	strVal = strVal & "&txtKeyStream="     & lgKeyStream
    strVal = strVal & "&lgStrPrevKey="  & lgStrPrevKey
    strVal = strVal & "&txtMaxRows="	& frm1.vspddata.MaxRows
    
    Call RunMyBizASP(MyBizASP, strVal)
    
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()

	Call SetToolBar("11000000110111")														'��: ��ư ���� ���� 
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call SetActiveCell(frm1.vspdData,1,1,"M","X","X")
		Set gActiveElement = document.activeElement
	End If

	Call SetQuerySpreadColor

	lgIntFlgMode = parent.OPMD_UMODE														'��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = False
	lgAfterQryFlg = True
	lgOldRow = 1
		
End Function

Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt, iMaxCnt

	With frm1.vspdData	

	.Redraw = False
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)

	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)
		
		.Col = -1
		.Row =  iArrColor2(0)
	
		Select Case iArrColor2(1)			
			Case "1"
				.BackColor = RGB(255,255,230)
			Case "2"				
				.BackColor = vbWhite' RGB(225,230,255) '������
'				.BackColor = RGB(245,250,255)
'				.ForeColor = RGB(200,5,200)
			Case "3"
				.BackColor = RGB(230,255,255)
'				.BackColor = RGB(235,240,245)
'				.ForeColor = RGB(5,5,200)
			Case "4"
				.BackColor = RGB(230,255,255)
'				.BackColor = RGB(225,230,235)
'				.ForeColor = RGB(200,5,5)
		End Select

	Next
	
	.Row =  -1
	.Col = C_OnhandQty
	.ForeColor = vbRed
	.Col = C_BadOnhandQty
	.ForeColor = vbRed
	
	iMaxCnt = .MaxRows
	
	For iLoopCnt=1 to iMaxCnt

		.Row =  iLoopCnt
		
		If iLoopCnt = 1 Then
			.Col = C_RcptQty
			.ForeColor = RGB(255,255,230)
			.Col = C_IssueQty
			.ForeColor = RGB(255,255,230)
			.Col = C_BadRcptQty
			.ForeColor = RGB(255,255,230)
			.Col = C_BadIssueQty
			.ForeColor = RGB(255,255,230)
			.Col = C_Price
			.ForeColor = RGB(255,255,230)
			.Col = C_Amount
			.ForeColor = RGB(255,255,230)
		ElseIf iLoopCnt = iMaxCnt Then
			.Col = C_RcptQty
			.ForeColor = vbBlue 
			.Col = C_IssueQty
			.ForeColor = RGB(200,50,200)
			.Col = C_BadRcptQty
			.ForeColor = vbBlue 
			.Col = C_BadIssueQty
			.ForeColor = RGB(200,50,200)
			.Col = C_BaseQty
			.ForeColor = RGB(230,255,255)
			.Col = C_Price
			.ForeColor = RGB(230,255,255)
			.Col = C_Amount
			.ForeColor = RGB(230,255,255)
		Else
			.Col = C_RcptQty
			If UNICDbl(.Text) = 0 Then
				.ForeColor = vbWhite
			Else
				.ForeColor = vbBlue 
			End If
			.Col = C_IssueQty
			If UNICDbl(.Text) = 0 Then
				.ForeColor = vbWhite
			Else
				.ForeColor = RGB(200,50,200) '������ 
			End If
			.Col = C_BadRcptQty
			If UNICDbl(.Text) = 0 Then
				.ForeColor = vbWhite
			Else
				.ForeColor = vbBlue 
			End If
			.Col = C_BadIssueQty
			If UNICDbl(.Text) = 0 Then
				.ForeColor = vbWhite
			Else
				.ForeColor = RGB(200,50,200) '������ 
			End If
			.Col = C_BaseQty
			.ForeColor = vbWhite
		End If

	Next

	.Redraw = True

	End With

End Sub

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

Function OpenReference()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	DIM IntRetCD
	Dim arrpb(0)            
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")		 '��: "Will you destory previous data"
		Exit Function
	End If
	
	
	
	If IsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("XI1611RA_KO244")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "XI1611RA_KO244", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	frm1.vspddata.row = frm1.vspddata.activerow
	frm1.vspddata.col = C_DocumentNo
	arrParam(0) = Trim(frm1.vspddata.value)
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.Parent, arrParam ), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0)= "" Then		
		Exit Function
	Else	

		
	End if	

	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>ǰ�񺰼�����Ȳ��ȸ</font></td>
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
			 						<TD CLASS=TD5 NOWRAP>����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12xxxU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="�����"></TD>
									<TD CLASS=TD5 NOWRAP>ǰ��</TD>
									<TD CLASS=TD6 NOWRAP>	<INPUT TYPE=TEXT NAME="txtItemCd" SIZE="15" MAXLENGTH="18" STYLE="Text-Transform: uppercase" ALT="ǰ��" tag="12XXXU" ><IMG align=top height=20 name="btnItemCd" onclick="vbscript:OpenItem()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtItemNm" CLASS=protected readonly=true TABINDEX="-1" SIZE="20" tag="14" >
								</TR>
			 					<TR>
									<TD CLASS=TD5 NOWRAP>������</TD> 
									<TD CLASS=TD6>
										<OBJECT classid=<%=gCLSIDFPDT%> name=txtReportFrDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="12" ALT="������" id=OBJECT1></OBJECT>
										&nbsp;~&nbsp;
										<OBJECT classid=<%=gCLSIDFPDT%> name=txtReportToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="11" ALT="������" id=OBJECT2></OBJECT>
									</TD>
									<TD CLASS=TD5 NOWRAP>â��׷�</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboSLGroup" ALT="â��׷�" STYLE="Width: 98px;" tag="11"><OPTION VALUE = ""></OPTION></SELECT></TD>
								</TR>
			 					<TR>
									<TD CLASS=TD5 NOWRAP></TD> 
									<TD CLASS=TD6>
									<TD CLASS=TD5 NOWRAP>â��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSlCd" SIZE="15" MAXLENGTH="7" STYLE="Text-Transform: uppercase" tag="11XXXU" ALT = "â��"><IMG align=top height=20 name="btnFrSlCd" onclick="vbscript:OpenSlCd()" src="../../../CShared/image/btnPopup.gif" width=16  TYPE="BUTTON">&nbsp;<input TYPE=TEXT NAME="txtSlNm" CLASS=protected readonly=true TABINDEX="-1" SIZE="20" tag="14" ></TD>
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
					<TABLE <%=LR_SPACE_TYPE_60%>>
						<TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�԰�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtspec" SIZE=40 MAXLENGTH=40 tag="24xxxU" ALT="�԰�"></TD>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtUnit" SIZE=10 MAXLENGTH=3 tag="24xxxU" ALT="����"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemAcct" SIZE=20 MAXLENGTH=20 tag="24xxx" ALT="ǰ�����"></TD>
								<TD CLASS=TD5 NOWRAP>���ޱ���</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtProcureType" SIZE=15 MAXLENGTH=20 tag="24xxx" ALT="���ޱ���"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>Location</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLocation" SIZE=20 MAXLENGTH=40 tag="24xxxU" ALT="Location"></TD>
								<TD CLASS=TD5 NOWRAP></TD>
								<TD CLASS=TD6 NOWRAP></TD>
							</TR>
							<TD HEIGHT="100%" colspan=4>
								<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData ID = "A" WIDTH=100% HEIGHT=100% tag="21" TITLE="SPREAD">
								<PARAM NAME="MaxCols" VALUE="0">
								<PARAM NAME="MaxRows" VALUE="0">
								</OBJECT>
							</TD>
						</TR>
				</TR>
			</TABLE>
		</TD>
	</TR>
	
	<!-- 	<TR HEIGHT=20>
	
		<TD WIDTH=100%>
			<TABLE  CLASS="BasicTB" CELLSPACING=0>
			    <tr>
                   <TD WIDTH=10>&nbsp;</TD>
                    <TD WIDTH=* ALIGN=RIGHT>
						<a href = "VBSCRIPT:OpenReference()">������������</A>
					</TD>	
				    <TD WIDTH=10>&nbsp;</TD>
                </tr>
			</TABLE>
		</TD>
	</TR>-->
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=20 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX = "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hReportFrDt" tag="24"><INPUT TYPE=HIDDEN NAME="hReportToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hSlCd" tag="24"><INPUT TYPE=HIDDEN NAME="hcboSLGroup" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>