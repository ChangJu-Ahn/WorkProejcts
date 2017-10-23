<%@ LANGUAGE="VBSCRIPT" %>
<!--**********************************************************************************************
'*  1. Module Name          : PROCUREMENT
'*  2. Function Name        : 
'*  3. Program ID           : U3116QA1.asp
'*  4. Program Name         : �ŷ�ó������Ȳ - Query Delivery Summary
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : NHG
'* 10. Modifier (Last)      : NHG
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'================================================================================================================================
Const BIZ_PGM_QRY1_ID	= "U3116QB1.asp"							'��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_QRY2_ID	= "U3116QB2.asp"							'��: �����Ͻ� ���� ASP�� 

'================================================================================================================================
' Grid 1(vspdData1) - Order
Dim C_IV_BIZ_AREA
Dim C_IV_DT
Dim C_IV_TYPE
Dim C_IV_CUR
Dim C_NET_DOC_AMT
Dim C_VAT_AMT
Dim C_GROSS_DOC_AMT
Dim C_VAT_RATE
Dim C_VAT_TYPE
Dim C_IV_NO


' Grid 2(vspdData2) - Result
Dim C_ITEM_CD
Dim C_ITEM_NM
Dim C_SPEC
Dim C_IV_UNIT
Dim C_IV_QTY
Dim C_IV_PRC
Dim C_IV_DOC_AMT
Dim C_VAT_DOC_AMT
Dim C_TOT_IV_AMT

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
	
	frm1.txtIvFrDt.text = UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, "01")
	frm1.txtIvToDt.text   = strDate
	Call SetBPCD()
End Sub

Sub SetBPCD()

	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(parent.gUsrId, "", "S"), lgF0) = False Then
		Call ggoOper.SetReqAttr(frm1.txtPlantCd,"Q")
		Call ggoOper.SetReqAttr(frm1.txtIvType,"Q")
		Call ggoOper.SetReqAttr(frm1.txtIvFrDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtIvToDt,"Q")
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
					
			.MaxCols = C_IV_NO + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("A")

			ggoSpread.SSSetEdit 	C_IV_BIZ_AREA,		"�����"		,15
			ggoSpread.SSSetDate 	C_IV_DT,			"��������"		,10, 2, parent.gDateFormat		 
			ggoSpread.SSSetEdit 	C_IV_TYPE,			"��������"		,20
			ggoSpread.SSSetEdit 	C_IV_CUR,			"ȭ�����"		,6
			ggoSpread.SSSetFloat 	C_NET_DOC_AMT,		"����ݾ�"		,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_VAT_AMT,			"VAT�ݾ�"		,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z" 
			ggoSpread.SSSetFloat 	C_GROSS_DOC_AMT,	"�Ѹ���ݾ�"	,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_VAT_RATE,			"VAT��"			,15, parent.ggExchRateNo,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec
			ggoSpread.SSSetEdit 	C_VAT_TYPE,			"VAT����"		,12
			ggoSpread.SSSetEdit 	C_IV_NO,			"�����ȣ"		,18

			
			Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
			ggoSpread.SSSetSplit2(2)
			
			Call SetSpreadLock("A")
			
			'.Col = 1 : .ColMerge = 2
			'.Col = 2 : .ColMerge = 2
			'.Col = 3 : .ColMerge = 2
			'.Col = 4 : .ColMerge = 2
			'.Col = 5 : .ColMerge = 2
			
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
					
			.MaxCols = C_TOT_IV_AMT + 1    
			.MaxRows = 0    
			
			Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit 	C_ITEM_CD,		"ǰ��"			,15
			ggoSpread.SSSetEdit		C_ITEM_NM,		"ǰ���"		,20
			ggoSpread.SSSetEdit 	C_SPEC,			"�԰�"			,15
			ggoSpread.SSSetEdit 	C_IV_UNIT,      "�������"		,6
			ggoSpread.SSSetFloat 	C_IV_QTY,		"�������"		,12,parent.ggQtyNo ,ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat 	C_IV_PRC,		"����ܰ�"		,15,parent.ggUnitCostNo,   ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_IV_DOC_AMT,	"����ݾ�"		,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_VAT_DOC_AMT,	"VAT�ݾ�"		,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_TOT_IV_AMT,	"�Ѹ���ݾ�"	,15,parent.ggAmtOfMoneyNo, ggStrIntegeralPart,	ggStrDeciPointPart,	parent.gComNum1000,	parent.gComNumDec,,,"Z"

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
		C_IV_BIZ_AREA		= 1
		C_IV_DT				= 2
		C_IV_TYPE			= 3
		C_IV_CUR			= 4
		C_NET_DOC_AMT		= 5
		C_VAT_AMT			= 6
		C_GROSS_DOC_AMT		= 7
		C_VAT_RATE			= 8
		C_VAT_TYPE			= 9
		C_IV_NO				= 10

	End If	
	
	If pvSpdNo = "B" Or pvSpdNo = "*" Then
		' Grid 2(vspdData2)
		C_ITEM_CD			= 1
		C_ITEM_NM			= 2
		C_SPEC				= 3
		C_IV_UNIT			= 4
		C_IV_QTY			= 5
		C_IV_PRC			= 6
		C_IV_DOC_AMT		= 7
		C_VAT_DOC_AMT		= 8
		C_TOT_IV_AMT		= 9
	End If	

End Sub

'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
      
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
		Case "A"
		
 			ggoSpread.Source = frm1.vspdData1
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
			C_IV_BIZ_AREA		= iCurColumnPos(1)
			C_IV_DT				= iCurColumnPos(2)
			C_IV_TYPE			= iCurColumnPos(3)
			C_IV_CUR			= iCurColumnPos(4)
			C_NET_DOC_AMT		= iCurColumnPos(5)
			C_VAT_AMT			= iCurColumnPos(6)
			C_GROSS_DOC_AMT		= iCurColumnPos(7)
			C_VAT_RATE			= iCurColumnPos(8)
			C_VAT_TYPE			= iCurColumnPos(9)
			C_IV_NO				= iCurColumnPos(10)
						
		Case "B"
		
			ggoSpread.Source = frm1.vspdData2
		
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ITEM_CD			= iCurColumnPos(1)
			C_ITEM_NM			= iCurColumnPos(2)
			C_SPEC				= iCurColumnPos(3)
			C_IV_UNIT			= iCurColumnPos(4)
			C_IV_QTY			= iCurColumnPos(5)
			C_IV_PRC			= iCurColumnPos(6)
			C_IV_DOC_AMT		= iCurColumnPos(7)
			C_VAT_DOC_AMT		= iCurColumnPos(8)
			C_TOT_IV_AMT		= iCurColumnPos(9)
			
    End Select

End Sub    

'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "������˾�"
	arrParam(1) = "B_BIZ_AREA"
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "�����"			
	
    arrField(0) = "BIZ_AREA_CD"	
    arrField(1) = "BIZ_AREA_NM"	
    
    arrHeader(0) = "�����"		
    arrHeader(1) = "������"		
    
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
Function OpenIVType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������"
	arrParam(1) = "S_BILL_TYPE_CONFIG"
	arrParam(2) = Trim(frm1.txtIvType.Value)								' Code Condition
	arrParam(3) = ""														' Name Cindition
	arrParam(4) = ""
	arrParam(5) = "��������"
	 
    arrField(0) = "BILL_TYPE"												' Field��(0)
    arrField(1) = "BILL_TYPE_NM"												' Field��(1)
    
    arrHeader(0) = "��������"													' Header��(0)
    arrHeader(1) = "�������¸�"													' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) <> "" Then
		Call SetIVType(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtIvType.focus

End Function

'================================================================================================================================
Function SetPlant(byval arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)
	frm1.txtPlantCd.focus()		
End Function

'================================================================================================================================
Function SetIVType(Byval arrRet)
	frm1.txtIvType.value = arrRet(0)
	frm1.txtIvTypeNm.value = arrRet(1)
	frm1.txtIvType.focus()
End Function

'================================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    Call LockObjectField(frm1.txtIvFrDt,"R")
    Call LockObjectField(frm1.txtIvToDt,"R")
    Call FormatDATEField(frm1.txtIvFrDt)
    Call FormatDATEField(frm1.txtIvToDt)
    
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
Sub txtIvFrDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtIvFrDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtIvFrDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtIvToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtIvToDt.Action = 7 
		Call SetFocusToDocument("M")
		frm1.txtIvToDt.Focus
	End If 
End Sub

'================================================================================================================================
Sub txtIvFrDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub

'================================================================================================================================
Sub txtIvToDt_KeyDown(keycode, shift)
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
	
	If frm1.txtIvType.value = "" Then
		frm1.txtIvTypeNm.value = "" 
	End If

	If ValidDateCheck(frm1.txtIvFrDt, frm1.txtIvToDt) = False Then Exit Function
		
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

	Dim strIvNo
	Dim dtMvmtDt

   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   Select Case pOpt
       Case "M"
       
				With frm1
					
						lgKeyStream = UCase(Trim(.txtPlantCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtIvType.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & UCase(Trim(.txtBPCd.value))  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtIvFrDt.Text)  & Parent.gColSep
						lgKeyStream = lgKeyStream & Trim(.txtIvToDt.Text)  & Parent.gColSep
						
						.hPlantCd.value		= .txtPlantCd.value
						.hIvType.value		= .txtIvType.value
						.hBPCd.value		= .txtBPCd.value
						.hIvFrDt.value		= .txtIvFrDt.Text
						.hIvToDt.value	= .txtIvToDt.Text
						
				
				End With
			
       Case "S"
				With frm1
					.vspdData1.Row = .vspdData1.ActiveRow
					.vspdData1.Col = C_IV_NO
					 strIvNo = .vspdData1.text
				
					lgKeyStream = UCase(Trim(strIvNo))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(.txtPlantCd.value))  & Parent.gColSep
					lgKeyStream = lgKeyStream & UCase(Trim(.hBPCd.value))  & Parent.gColSep
					
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
	
	If frm1.hPlantCd.value = "" Then
		var1 = "%"
	Else
		var1 = Trim(frm1.hPlantCd.value)
	End If
	If frm1.hIvType.value = "" Then
		var2 = "%"
	Else
		var2 = Trim(frm1.hIvType.value)
	End If
	If frm1.hBPCd.value = "" Then
		var3 = "%"
	Else
		var3 = Trim(frm1.hBPCd.value)
	End If
	If frm1.hIvFrDt.value = "" Then
		var4 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var4 = UniConvDateAtoB(frm1.hIvFrDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
	If frm1.hIvToDt.value = "" Then
		var5 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var5 = UniConvDateAtoB(frm1.hIvToDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
	
	strUrl = strUrl & "PLANT|" & var1
	strUrl = strUrl & "|TYPE|" & var2
	strUrl = strUrl & "|BP|" & var3
	strUrl = strUrl & "|FRDT|" & var4
	strUrl = strUrl & "|TODT|" & var5

	
	strEbrFile = "U3113QA1"
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

	If frm1.hPlantCd.value = "" Then
		var1 = "%"
	Else
		var1 = Trim(frm1.hPlantCd.value)
	End If
	If frm1.hIvType.value = "" Then
		var2 = "%"
	Else
		var2 = Trim(frm1.hIvType.value)
	End If
	If frm1.hBPCd.value = "" Then
		var3 = "%"
	Else
		var3 = Trim(frm1.hBPCd.value)
	End If
	If frm1.hIvFrDt.value = "" Then
		var4 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var4 = UniConvDateAtoB(frm1.hIvFrDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
	If frm1.hIvToDt.value = "" Then
		var5 = UniConvDateAtoB(UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31"),parent.gDateFormat,parent.gServerDateFormat)
	Else
		var5 = UniConvDateAtoB(frm1.hIvToDt.value,parent.gDateFormat,parent.gServerDateFormat)
	End If
	
	strUrl = strUrl & "PLANT|" & var1
	strUrl = strUrl & "|TYPE|" & var2
	strUrl = strUrl & "|BP|" & var3
	strUrl = strUrl & "|FRDT|" & var4
	strUrl = strUrl & "|TODT|" & var5
	
	strEbrFile = "U3113QA1"
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ŷ�ó������Ȳ</font></td>
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
			 						<TD CLASS=TD5 NOWRAP>��������</TD> 
									<TD CLASS=TD6>
										<script language =javascript src='./js/u3116qa1_OBJECT1_txtIvFrDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/u3116qa1_OBJECT2_txtIvToDt.js'></script>
									</TD>
								</TR>
			 					<TR>
									<TD CLASS=TD5 NOWRAP>�����</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="11xxxU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14" ALT="������"></TD>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtIvType" SIZE=18 MAXLENGTH=18 tag="11xxxU" ALT="��������"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvType()">&nbsp;<INPUT TYPE=TEXT NAME="txtIvTypeNm" SIZE=25 tag="14"></TD>
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
									<script language =javascript src='./js/u3116qa1_A_vspdData1.js'></script>
								</TD>
							</TR>
							<TR HEIGHT="50%">
								<TD WIDTH="100%">
									<script language =javascript src='./js/u3116qa1_B_vspdData2.js'></script>
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
					<TD ><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>�μ�</BUTTON>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
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
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hIvFrDt" tag="24"><INPUT TYPE=HIDDEN NAME="hIvToDt" tag="24"><INPUT TYPE=HIDDEN NAME="hIvType" tag="24"><INPUT TYPE=HIDDEN NAME="hBpCd" tag="24">
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