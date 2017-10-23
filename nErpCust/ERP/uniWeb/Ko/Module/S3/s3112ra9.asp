<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : ����																		*
'*  2. Function Name        : 																			*
'*  3. Program ID           : S3112RA9																	*
'*  4. Program Name         : Ŭ��������(���ֳ������)													*
'*  5. Program Desc         :																			*
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/02/06																*
'*  8. Modified date(Last)  : 																			*
'*  9. Modifier (First)     : Hwang Seong Bae															*
'* 10. Modifier (Last)      :																			*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									*
'*                            this mark(��) Means that "may  change"									*
'*                            this mark(��) Means that "must change"									*
'* 13. History              : 																			*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '��: �ش� ��ġ�� ���� �޶���, ��� ��� %>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit					
	

Const BIZ_PGM_ID = "s3112rb9.asp"        

Const C_PopClassCd		= 1
Const C_PopCharValueCd1	= 2
Const C_PopCharValueCd2	= 3
Const C_PopPlantCd		= 4

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim gblnWinEvent															   
Dim	lgStrCharCd1		' Class�� �Ҵ�� ���1(Class ���ý� ����)
Dim lgStrCharCd2		' Class�� �Ҵ�� ���2(Class ���ý� ����)
Dim lgObjCaller			' ȣ�� Window(���ֳ��� document)
Dim lgLngTotalInsertedRows		' �߰��� Row �� 
Dim lgBlnChgClass

'��:--------Spreadsheet #1-----------------------------------------------------------------------------   
Dim C_PlantCd
Dim C_PlantNm
Dim C_ClassCd
Dim C_ClassNm
Dim C_CharValueCd11
Dim C_CharValueNm11
Dim C_CharValueCd12
Dim C_CharValueNm12
Dim C_Unit
Dim C_UnitPopup
Dim C_Qty
Dim C_Price
Dim C_Amt
Dim C_InvQty
Dim C_RcptQty
Dim C_ItemCd
Dim C_ItemNm
Dim C_Spec
Dim C_HsCd
Dim C_VatType
Dim C_VatNm
Dim C_VatRate
Dim C_OldQty
Dim C_OldAmt
Dim C_Pointer

'��:--------Spreadsheet #2-----------------------------------------------------------------------------   
Dim C_ItemCd2			'ǰ��  
Dim C_ItemNm2			'ǰ��� 
Dim C_Spec2				'�԰�	 
Dim C_Unit2				'���� 
Dim C_Qty2				'���� 
Dim C_Price2			'�ܰ� 
Dim C_Amt2				'�ݾ� 
Dim C_PlantCd2			'�����ڵ� 
Dim C_PlantNm2			'����� 
Dim C_HsCd2				'HS��ȣ 
Dim C_VatType2			'	
Dim C_VatNm2			
Dim C_VatRate2				

'======================================  Global Variable�� ����  ==================================
Set lgObjCaller = window.dialogArguments
Set PopupParent = lgObjCaller.parent
top.document.title = PopupParent.gActivePRAspName


'========================================================================================================
Sub initSpreadPosVariables()
   	If gMouseClickStatus = "N" Or gMouseClickStatus = "SPCRP" Then
		C_PlantCd		= 1
		C_PlantNm		= 2
		C_ClassCd		= 3
		C_ClassNm		= 4
		C_CharValueCd11	= 5
		C_CharValueNm11	= 6
		C_CharValueCd12	= 7
		C_CharValueNm12	= 8
		C_Unit			= 9
		C_UnitPopup		= 10
		C_Qty			= 11
		C_Price			= 12
		C_Amt			= 13
		C_InvQty		= 14
		C_RcptQty		= 15
		C_ItemCd		= 16
		C_ItemNm		= 17
		C_Spec			= 18
		C_HsCd			= 19
		C_VatType		= 20
		C_VatNm			= 21
		C_VatRate		= 22
		C_OldQty		= 23
		C_OldAmt		= 24
		C_Pointer		= 25
	End If
	
   	If gMouseClickStatus = "N" Then
		C_ItemCd2		= 1
		C_ItemNm2		= 2
		C_Spec2		= 3
		C_Unit2			= 4
		C_Qty2			= 5
		C_Price2		= 6
		C_Amt2			= 7
		C_PlantCd2		= 8
		C_PlantNm2		= 9
		C_HsCd2			= 10
		C_VatType2		= 11
		C_VatNm2	= 12
		C_VatRate2		= 13
	End If
	
End Sub

'========================================================================================================
Function InitVariables()
	lgIntFlgMode = PopupParent.OPMD_CMODE								
	lgIntGrpCount = 0										
	lgStrPrevKey = ""										
	lgBlnChgClass = False		
	gblnWinEvent = False
End Function
	
'========================================================================================================
Sub SetDefaultVal()
	lgLngTotalInsertedRows = 0
	
	With frm1
		.txtCurrency.value = lgObjCaller.frm1.txtCurrency.value	' ȭ����� 
		If Trim(lgObjCaller.frm1.txtPlant.value) <> "" Then			' ���� 
			.txtConPlantCd.value = Trim(lgObjCaller.frm1.txtPlant.value)
		Else
			.txtConPlantCd.value = PopupParent.gPlant
		End If
	End With

	Call SetReqAttr()
End Sub


'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "PA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "PA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables() 
	
	If gMouseClickStatus = "N" Or gMouseClickStatus = "SPCRP" Then
		With frm1.vspdData		
				
		   	'��:--------Spreadsheet #1-----------------------------------------------------------------------------   
			ggoSpread.Source = frm1.vspdData

		    ggoSpread.Spreadinit "V20021214",,PopupParent.gAllowDragDropSpread    		
			.ReDraw = false
				
			.MaxRows = 0 : .MaxCols = 0
			.MaxCols = C_Pointer + 1            '��: �ִ� Columns�� �׻� 1�� ������Ŵ 
			    
		    Call GetSpreadColumnPos("A")
			    
			ggoSpread.SSSetEdit		C_PlantCd,		"����", 10,,,4,2 
			ggoSpread.SSSetEdit		C_PlantNm,		"�����", 18
			ggoSpread.SSSetEdit		C_ClassCd,		"Ŭ����", 10,,,16,2 
			ggoSpread.SSSetEdit		C_ClassNm,		"Ŭ������", 18
			ggoSpread.SSSetEdit		C_CharValueCd11,"���1", 10,,,16,2 
			ggoSpread.SSSetEdit		C_CharValueNm11,"����1", 18
			ggoSpread.SSSetEdit		C_CharValueCd12,"���2", 10,,,16,2 
			ggoSpread.SSSetEdit		C_CharValueNm12,"����2", 18
			ggoSpread.SSSetEdit		C_Unit,			"����", 8,2,,3,2
		    ggoSpread.SSSetButton	C_UnitPopup
			ggoSpread.SSSetFloat	C_Qty,			"����" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Price,		"�ܰ�",15,PopupParent.ggUnitCostNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_Amt,			"�ݾ�",15,PopupParent.ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_InvQty,		"������" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat	C_RcptQty,		"�԰�����" ,15,PopupParent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec,,,"Z"
			ggoSpread.SSSetEdit		C_ItemCd,		"ǰ��", 18,,,18,2 
			ggoSpread.SSSetEdit		C_ItemNm,		"ǰ���", 40
			ggoSpread.SSSetEdit		C_Spec,			"�԰�", 40
			ggoSpread.SSSetEdit		C_HsCd,			"", 1
			ggoSpread.SSSetEdit		C_VatType,		"", 1
			ggoSpread.SSSetEdit		C_VatNm,		"", 1
			ggoSpread.SSSetEdit		C_VatRate,		"", 1
			ggoSpread.SSSetEdit		C_OldQty,		"", 1
			ggoSpread.SSSetEdit		C_OldAmt,		"", 1
			ggoSpread.SSSetEdit		C_Pointer,		"", 1

			Call ggoSpread.MakePairsColumn(C_PlantCd,C_ItemNm)
			Call ggoSpread.MakePairsColumn(C_ClassCd,C_ClassNm)
			Call ggoSpread.MakePairsColumn(C_CharValueCd11,C_CharValueNm11)
			Call ggoSpread.MakePairsColumn(C_CharValueCd12,C_CharValueNm12)
			Call ggoSpread.MakePairsColumn(C_Unit,C_UnitPopup)
			    
		    Call ggoSpread.SSSetColHidden(C_HsCd,.MaxCols,True)   '��: ������Ʈ�� ��� Hidden Column
			    
   		    Call SetSpreadLock()

			.ReDraw = True
		End With
	End If
	
   	'��:--------Spreadsheet #2-----------------------------------------------------------------------------   
   	If gMouseClickStatus = "N" Then
		With frm1.vspdData2		
			
			ggoSpread.Source = frm1.vspdData2
			'patch version
		    ggoSpread.Spreadinit "V20021214",,PopupParent.gAllowDragDropSpread    		
			.ReDraw = false
			
			.MaxRows = 0 : .MaxCols = 0
			.MaxCols = C_VatRate2 + 1            '��: �ִ� Columns�� �׻� 1�� ������Ŵ 
		    
		    Call GetSpreadColumnPos("B")

			ggoSpread.SSSetEdit		C_ItemCd2,		"", 1
			ggoSpread.SSSetEdit		C_ItemNm2,	"", 1
			ggoSpread.SSSetEdit		C_Spec2,	"", 1
			ggoSpread.SSSetEdit		C_Qty2,			"", 1
			ggoSpread.SSSetEdit		C_Price2,		"", 1
			ggoSpread.SSSetEdit		C_Amt2,			"", 1
			ggoSpread.SSSetEdit		C_PlantCd2,		"", 1
			ggoSpread.SSSetEdit		C_PlantNm2,		"", 1
			ggoSpread.SSSetEdit		C_HsCd2,		"", 1
			ggoSpread.SSSetEdit		C_VatType2,		"", 1
			ggoSpread.SSSetEdit		C_VatNm2,	"", 1
			ggoSpread.SSSetEdit		C_VatRate2,		"", 1
		    
			.ReDraw = True
		End With
	End If
End Sub

'========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_PlantCd		= iCurColumnPos(1)
			C_PlantNm		= iCurColumnPos(2)
			C_ClassCd		= iCurColumnPos(3)
			C_ClassNm		= iCurColumnPos(4)
			C_CharValueCd11	= iCurColumnPos(5)
			C_CharValueNm11	= iCurColumnPos(6)
			C_CharValueCd12	= iCurColumnPos(7)
			C_CharValueNm12	= iCurColumnPos(8)
			C_Unit			= iCurColumnPos(9)
			C_UnitPopup		= iCurColumnPos(10)
			C_Qty			= iCurColumnPos(11)
			C_Price			= iCurColumnPos(12)
			C_Amt			= iCurColumnPos(13)
			C_InvQty		= iCurColumnPos(14)
			C_RcptQty		= iCurColumnPos(15)
			C_ItemCd		= iCurColumnPos(16)
			C_ItemNm		= iCurColumnPos(17)
			C_Spec			= iCurColumnPos(18)
			C_HsCd			= iCurColumnPos(19)
			C_VatType		= iCurColumnPos(20)
			C_VatNm			= iCurColumnPos(21)
			C_VatRate		= iCurColumnPos(22)
			C_OldQty		= iCurColumnPos(23)
			C_OldAmt		= iCurColumnPos(24)
			C_Pointer		= iCurColumnPos(25)
			
       Case "B"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_ItemCd2		= iCurColumnPos(1)
			C_ItemNm2		= iCurColumnPos(2)
			C_Spec2			= iCurColumnPos(3)
			C_Unit2			= iCurColumnPos(4)
			C_Qty2			= iCurColumnPos(5)
			C_Price2		= iCurColumnPos(6)
			C_Amt2			= iCurColumnPos(7)
			C_PlantCd2		= iCurColumnPos(8)
			C_PlantNm2		= iCurColumnPos(9)
			C_HsCd2			= iCurColumnPos(10)
			C_VatType2		= iCurColumnPos(11)
			C_VatNm2		= iCurColumnPos(12)
			C_VatRate2		= iCurColumnPos(13)

    End Select    
End Sub
	
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock C_PlantCd, -1, C_CharValueNm12
	ggoSpread.SpreadLock C_Amt, -1
End Sub

'========================================================================================================	
Sub SetQuerySpreadColor(ByVal pvStartRow)
	ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		ggoSpread.SSSetRequired  C_UNIT			, pvStartRow, .MaxRows
		ggoSpread.SSSetRequired  C_QTY			, pvStartRow, .MaxRows
		ggoSpread.SSSetRequired  C_PRICE		, pvStartRow, .MaxRows
	End With
	
End Sub

' ���2�� ���� Class�� ���2�� Hidden ó���Ѵ�.
Sub SetColHiddenByClass()
	Dim iHiddenType
	ggoSpread.Source = frm1.vspdData
	 
	With frm1.vspdData
		' ���2�� ���� ��쿡�� �ش� �ʵ带 Hidden ó�� 
		.Col = C_CharValueCd12
		If lgStrCharCd2 = "" Then
			If Not .ColHidden Then
				Call ggoSpread.SSSetColHidden(C_CharValueCd12, C_CharValueNm12, True)
			End If
		Else
			If .ColHidden Then
				Call ggoSpread.GetHiddenCol(iHiddenType)
				' User Hiddenó������ ���� ��쿡�� �����ش� 
				' iHiddenType(C_CharValueCd12) = -1(user�� hidden ó��)
				If iHiddenType(C_CharValueCd12) <> -1 Then Call ggoSpread.SSSetColHidden(C_CharValueCd12, C_CharValueNm12, False)
			End If
		End If
	End With
End Sub

'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If gblnWinEvent Then Exit Function

	gblnWinEvent = True
	
	Select Case pvIntWhere
	Case C_PopClassCd												
		iArrParam(1) = "B_CLASS"						' TABLE ��Ī 
		iArrParam(2) = Trim(frm1.txtConClassCd.value)	' Code Condition
		iArrParam(3) = ""								' Name Cindition
		iArrParam(4) = ""								' Where Condition
		iArrParam(5) = "Ŭ����"						' TextBox �� 
		
		iArrField(0) = "ED15" & PopupParent.gColSep & "CLASS_CD"	' Field�� 
		iArrField(1) = "ED30" & PopupParent.gColSep & "CLASS_NM"
		iArrField(2) = "HH1" & PopupParent.gColSep & "CHAR_CD1"
		iArrField(3) = "HH1" & PopupParent.gColSep & "CHAR_CD2"
    
	    iArrHeader(0) = "Ŭ����"						
	    iArrHeader(1) = "Ŭ������"					

		frm1.txtConClassCd.focus 

	Case C_PopCharValueCd1
		
		If frm1.txtConCharValueCd1.readOnly = True Then
			gblnWinEvent = False
			Exit Function
		End If
														
		iArrParam(1) = "B_CHAR_VALUE"
		iArrParam(2) = Trim(frm1.txtConCharValueCd1.value)
		iArrParam(3) = ""
		iArrParam(4) = "CHAR_CD =  " & FilterVar(lgStrCharCd1 , "''", "S") & ""
		iArrParam(5) = "���"
		
		iArrField(0) = "ED15" & PopupParent.gColSep & "CHAR_VALUE_CD"
		iArrField(1) = "ED30" & PopupParent.gColSep & "CHAR_VALUE_NM"
    
	    iArrHeader(0) = "���"
	    iArrHeader(1) = "��缳��"

		frm1.txtConCharValueCd1.focus 

	Case C_PopCharValueCd2												
	
		If frm1.txtConCharValueCd2.readOnly = True Then
			gblnWinEvent = False
			Exit Function
		End If
		
		iArrParam(1) = "B_CHAR_VALUE"
		iArrParam(2) = Trim(frm1.txtConCharValueCd2.value)
		iArrParam(3) = ""
		iArrParam(4) = "CHAR_CD =  " & FilterVar(lgStrCharCd2 , "''", "S") & ""
		iArrParam(5) = "���"
		
		iArrField(0) = "ED15" & PopupParent.gColSep & "CHAR_VALUE_CD"
		iArrField(1) = "ED30" & PopupParent.gColSep & "CHAR_VALUE_NM"
    
	    iArrHeader(0) = "���"
	    iArrHeader(1) = "��缳��"

		frm1.txtConCharValueCd2.focus 

	Case C_PopPlantCd
		iArrParam(1) = "B_PLANT"
		iArrParam(2) = Trim(frm1.txtConPlantCd.value)
		iArrParam(3) = ""
		iArrParam(4) = ""
		iArrParam(5) = "����"
		
		iArrField(0) = "ED15" & PopupParent.gColSep & "PLANT_CD"
		iArrField(1) = "ED30" & PopupParent.gColSep & "PLANT_NM"
    
	    iArrHeader(0) = "����"
	    iArrHeader(1) = "�����"

		frm1.txtConPlantCd.focus 

	End Select
 
	iArrParam(0) = iArrParam(5)							 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

'========================================================================================================
' Spread button popup
Function OpenSpreadPopup(ByVal pvLngCol, ByVal pvLngRow, ByVal pvStrData)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenSpreadPopup = False
	
	If gblnWinEvent Then Exit Function

	gblnWinEvent = True
	
	Select Case pvLngCol
		Case C_UnitPopup
			iArrParam(1) = "dbo.B_UNIT_OF_MEASURE "			
			iArrParam(2) = pvStrData						
			iArrParam(3) = ""								
			iArrParam(4) = " DIMENSION <> " & FilterVar("TM", "''", "S") & " "			
			iArrParam(5) = "����"						
				
			iArrField(0) = "ED15" & PopupParent.gColSep & "UNIT"
			iArrField(1) = "ED30" & PopupParent.gColSep & "UNIT_NM"
			    
			iArrHeader(0) = "����"
			iArrHeader(1) = "������"
	End Select
 
	iArrParam(0) = iArrParam(5)							 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenSpreadPopup = SetSpreadPopup(iArrRet,pvLngCol, pvLngRow)
	End If	

End Function

'========================================================================================================
' Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		Case C_PopClassCd
			.txtConClassCd.value = pvArrRet(0) 
			.txtConClassNm.value = pvArrRet(1)   
			lgStrCharCd1 = pvArrRet(2)
			lgStrCharCd2 = pvArrRet(3)

			Call SetReqAttr()

		Case C_PopCharValueCd1
			.txtConCharValueCd1.value = pvArrRet(0)
			.txtConCharValueNm1.value = pvArrRet(1)

		Case C_PopCharValueCd2
			.txtConCharValueCd2.value = pvArrRet(0)
			.txtConCharValueNm2.value = pvArrRet(1)

		Case C_PopPlantCd
			.txtConPlantCd.value = pvArrRet(0) 
			.txtConPlantNm.value = pvArrRet(1)   

		End Select
	End With

	SetConPopup = True
End Function

'========================================================================================================
Function SetSpreadPopup(Byval pvArrRet,ByVal pvLngCol, ByVal pvLngRow)
	SetSpreadPopup = False

	With frm1.vspdData
		.Row = pvLngRow
		
		Select Case pvLngCol
			Case C_UnitPopup
				.Col = C_Unit			: .Text = pvArrRet(0)
				' �ܰ� Fetch
				Call GetItemPrice(pvLngRow)
		End Select
	End With

	SetSpreadPopup = True
End Function

'========================================================================================================
' Name : GetClassInfo
' Description : Class ������ �����´�.
Function GetClassInfo()
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrTemp
	Dim iArrRs(3)
	
	GetClassInfo = False
	
	iStrSelectList = " CLASS_CD, CLASS_NM, CHAR_CD1, ISNULL(CHAR_CD2, '') "
	iStrFromList  = " B_CLASS "
	iStrWhereList = " CLASS_CD =  " & FilterVar(frm1.txtConClassCd.value, "''", "S") & ""
		
	Err.Clear
	    
	'�ܰ� Fetch
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))

		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		iArrRs(3) = iArrTemp(4)
			
		GetClassInfo = SetConPopup(iArrRs, C_PopClassCd)
	Else
		If Err.number = 0 Then
			GetClassInfo = OpenConPopup(C_PopClassCd)
		Else
			MsgBox Err.description 
			Err.Clear
			Exit Function
		End If
	End If

End Function
'========================================================================================================
' ���1, ���2 �ʵ� �������ɿ��� ó�� 
Sub SetReqAttr()
	With frm1
		If lgStrCharCd1 <> "" Then
			Call ggoOper.SetReqAttr(.txtConCharValueCd1, "D")
		Else
			Call ggoOper.SetReqAttr(.txtConCharValueCd1, "Q")
		End If
		
		If lgStrCharCd2 <> "" Then
			Call ggoOper.SetReqAttr(.txtConCharValueCd2, "D")
		Else
			Call ggoOper.SetReqAttr(.txtConCharValueCd2, "Q")
		End If
	End With
	
	lgBlnChgClass = False
End Sub

'========================================================================================================
Function GetItemPrice(ByVal pvLngRow)
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim  iStrItemCd, iStrUnit
	Dim iStrRs
	Dim iArrPrice
	Dim iDblOldPrice
	
	GetItemPrice = False
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_ItemCd			:	iStrItemCd = Trim(.Text)
		.Col = C_Unit			:	iStrUnit = Trim(.Text)
		
		iStrSelectList = " dbo.ufn_s_GetItemSalesPrice('" & FilterVar(lgObjCaller.frm1.txtSoldToParty.value, "''", "S")	& "," _
														  & FilterVar(iStrItemCd, "''", "S")			& ", " _
														  & FilterVar(lgObjCaller.frm1.txtHDealType.value, "''", "S")		& ", " _
														  & FilterVar(Trim(lgObjCaller.frm1.txtHPayTermsCd.value), "''", "S")	& ", " _
														  & FilterVar(iStrUnit, "''", "S")	& ", " _
														  & FilterVar(lgObjCaller.frm1.txtCurrency.value, "''", "S")	& ", '" _
														  & UNIConvDate(lgObjCaller.frm1.txtSoDt.value)		& "') "
		iStrFromList  = ""
		iStrWhereList = ""
		
		Err.Clear
	    
		'�ܰ� Fetch
		If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
			iArrPrice = Split(iStrRs, Chr(11))
			
			.Col = C_Price
			iDblOldPrice = UNICDbl(.text)
			.Text = UNIConvNumPCToCompanyByCurrency(iArrPrice(1), lgObjCaller.frm1.txtCurrency.value, PopUpParent.ggUnitCostNo, "X" , "X")
			
			' �ݾ� ���� 
			If iDblOldPrice <> Cdbl(iArrPrice(1)) Then
				Call CalcAmt(pvLngRow, C_Price)
			End If
			
			GetItemPrice = True
			Exit Function
		Else
			If Err.number <> 0 Then
				MsgBox Err.description 
				Err.Clear
				Exit Function
			End If
		End If
	End With

End Function

'========================================================================================================
Sub CalcAmt(ByVal pvLngRow, ByVal pvLngCol)
	Dim iStrCur, iStrNewAmt
	Dim iDblQty, iDblOldQty, iDblPrice, iDblOldAmt, iDblNewAmt
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_Qty	:	iDblQty = UNICDbl(.Text)
		.Col = C_OldQty	:	iDblOldQty = UNICDbl(.Text)
		.Col = C_Price	:	iDblPrice = UNICDbl(.Text)
		
		' ������ ����� ��� 
		If pvLngCol = C_Qty And iDblPrice = 0 Then
			If iDblOldQty <> iDblQty Then Call CalcTotal(pvLngRow, iDblQty - iDblOldQty, 0)

			Exit Sub
		End If

		' �ܰ��� ����� ��� 
		If pvLngCol = C_Price And iDblQty = 0 Then Exit Sub
		
		iStrCur = frm1.txtCurrency.value
		.Col = C_Amt
		iDblOldAmt = UNICDbl(.Text)
		iDblNewAmt = iDblQty * iDblPrice
		
		iStrNewAmt = UNIConvNumPCToCompanyByCurrency(iDblNewAmt,iStrCur,PopUpParent.ggAmtOfMoneyNo, "X" , "X")
		.Text = iStrNewAmt
		iDblNewAmt = UNICDbl(iStrNewAmt)
		
		If (iDblOldAmt <> iDblNewAmt) Or (iDblOldQty <> iDblQty) Then
			Call CalcTotal(pvLngRow, iDblQty - iDblOldQty,iDblNewAmt - iDblOldAmt)
			.Col = C_OldAmt	: .Text = iStrNewAmt
		End If
	End With

End Sub

'========================================================================================================
Sub CalcTotal(ByVal pvLngRow, ByVal pvDblQty, ByVal pvDblAmt)
	With frm1
		.txtTotQty.Text = UNIFormatNumber(UNICDbl(.txtTotQty.Text) + pvDblQty,ggQty.DecPoint,-2,0,ggQty.RndPolicy,ggQty.RndUnit)	
		.txtTotAmt.Text = UNIConvNumPCToCompanyByCurrency(UNICDbl(.txtTotAmt.Text) + pvDblAmt,.txtCurrency.value,PopUpParent.ggAmtOfMoneyNo, "X" , "X")
	End With
	
	Call InsertRowIntoSpread2(pvLngRow)
End Sub

'========================================================================================================
Sub InsertRowIntoSpread2(ByVal pvLngRow)
	Dim iStrPointer
	Dim iStrItemCd, iStrItemNm, iStrSpec, iStrPlantCd, iStrPlantNm
	Dim iStrUnit, iStrPrice, iStrQty, iStrAmt
	Dim iStrHsCd, iStrVatType, iStrVatNm, iStrVatRate
	
	With frm1.vspdData
		.Row = pvLngRow
		.Col = C_Pointer	: iStrPointer = Trim(.Text)
		.Col = C_Unit		: iStrUnit	  = Trim(.Text)
		.Col = C_Price		: iStrPrice	  = Trim(.Text)
		.Col = C_Qty		: iStrQty	  = Trim(.Text)
		.Col = C_Amt		: iStrAmt	  = Trim(.Text)
				
		If iStrPointer <> "" Then
			With frm1.vspdData2
				.Row = CLng(iStrPointer)
				.Col = C_Unit2		: .Text = iStrUnit
				.Col = C_Price2		: .Text = iStrPrice
				.Col = C_Qty2		: .Text = iStrQty
				.Col = C_Amt2		: .Text = iStrAmt
			End With
		Else
			.Col = C_ItemCd		: iStrItemCd  = Trim(.Text)
			.Col = C_ItemNm		: iStrItemNm  = Trim(.Text)
			.Col = C_Spec		: iStrSpec	  = Trim(.Text)
			.Col = C_PlantCd	: iStrPlantCd = Trim(.Text)
			.Col = C_PlantNm	: iStrPlantNm = Trim(.Text)
			.Col = C_HsCd		: iStrHsCd	  = Trim(.Text)
			.Col = C_VatType	: iStrVatType = Trim(.Text)
			.Col = C_VatNm		: iStrVatNm	  = Trim(.Text)
			.Col = C_VatRate	: iStrVatRate = Trim(.Text)
		
			With frm1.vspdData2
				.MaxRows = .MaxRows + 1
				.Row = .MaxRows

				.Col = C_ItemCd2	: .Text = iStrItemCd
				.Col = C_ItemNm2	: .Text = iStrItemNm
				.Col = C_Spec2		: .Text = iStrSpec
				.Col = C_Unit2		: .Text = iStrUnit
				.Col = C_Qty2		: .Text = iStrQty
				.Col = C_Price2		: .Text = iStrPrice
				.Col = C_Amt2		: .Text = iStrAmt
				.Col = C_PlantCd2	: .Text = iStrPlantCd
				.Col = C_PlantNm2	: .Text = iStrPlantNm
				.Col = C_HsCd2		: .Text = iStrHsCd
				.Col = C_VatType2	: .Text = iStrVatType
				.Col = C_VatNm2	: .Text = iStrVatNm
				.Col = C_VatRate2	: .Text = iStrVatRate

				iStrPointer = CStr(.MaxRows)
			End With
			
			' Set the Pointer
			.Col = C_Pointer
			.Text = iStrPointer
		End If
	End With
End Sub

'========================================================================================================
Sub InsertRowsIntoSoDtl()

	Dim iStrItemCd, iStrItemNm, iStrSpec, iStrPlantCd, iStrPlantNm
	Dim iStrUnit, iStrPrice, iStrQty, iStrAmt
	Dim iStrHsCd, iStrVatType, iStrVatNm, iStrVatRate
	Dim iLngRow, iLngInsertedRows
	Dim iIntIndex
	
	iLngInsertedRows = 0
	
	With frm1.vspdData2
		For iLngRow = 1 To .MaxRows
			.Row = iLngRow
			.Col = C_Qty2		: iStrQty = Trim(.Text)
			
			If UNICDbl(iStrQty) > 0 Then
				iLngInsertedRows = iLngInsertedRows + 1
				
				.Col = C_ItemCd2	: iStrItemCd  = Trim(.Text)
				.Col = C_ItemNm2	: iStrItemNm  = Trim(.Text)
				.Col = C_Spec2		: iStrSpec	  = Trim(.Text)
				.Col = C_Unit2		: iStrUnit	  = Trim(.Text)
				.Col = C_Price2		: iStrPrice	  = Trim(.Text)
				.Col = C_Amt2		: iStrAmt	  = Trim(.Text)
				.Col = C_PlantCd2	: iStrPlantCd = Trim(.Text)
				.Col = C_PlantNm2	: iStrPlantNm = Trim(.Text)
				.Col = C_HsCd2		: iStrHsCd	  = Trim(.Text)
				.Col = C_VatType2	: iStrVatType = Trim(.Text)
				.Col = C_VatNm2		: iStrVatNm	  = Trim(.Text)
				.Col = C_VatRate2	: iStrVatRate = Trim(.Text)
				
				With lgObjCaller.frm1.vspdData
					lgObjCaller.FncInsertRow(1)
					.Row = .ActiveRow
					
					.Col = lgObjCaller.C_ItemCd		: .Text = iStrItemCd
					.Col = lgObjCaller.C_ItemName	: .Text = iStrItemNm
					.Col = lgObjCaller.C_ItemSpec	: .Text = iStrSpec
					.Col = lgObjCaller.C_SoUnit		: .Text = iStrUnit
					.Col = lgObjCaller.C_SoQty		: .Text = iStrQty
					.Col = lgObjCaller.C_SoPrice	: .Text = iStrPrice
					.Col = lgObjCaller.C_TotalAmt	: .Text = iStrAmt
					.Col = lgObjCaller.C_PlantCd	: .Text = iStrPlantCd
					.Col = lgObjCaller.C_PlantNm	: .Text = iStrPlantNm
					.Col = lgObjCaller.C_HsNo		: .Text = iStrHsCd
					' ���������� Vat������ ��ϵ��� ���� ��� ǰ�� �Ҵ�� VAT������ �����Ѵ�.
					If Trim(lgObjCaller.frm1.txtHVATType.value) = "" Then
						.Col = lgObjCaller.C_VatType	: .Text = iStrVatType
						.Col = lgObjCaller.C_VatTypeNm	: .Text = iStrVatNm
						.Col = lgObjCaller.C_VatRate	: .Text = iStrVatRate
					End If
					
					' �ΰ��� �ݾ� ��� 
					Call lgObjCaller.TotalAmtChange(.Row)
		
					Call lgObjCaller.SetTrackingNoByItem(.Row)
				End With
			End If
		Next
	End With
	
	If iLngInsertedRows = 0 Then
		'�߰��� ǰ���� �����ϴ�.
		Call DisplayMsgBox("203253", "X", "X", "X")	
	Else
		lgLngTotalInsertedRows = lgLngTotalInsertedRows + iLngInsertedRows
		' ���� ǰ���� �߰��Ͽ����ϴ�.
		Call DisplayMsgBox("203254", "X", CStr(iLngInsertedRows), "X")	
		Call ggoOper.ClearField(Document, "2")								
	End If
	
End Sub

'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029											<%  %>
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")						<%  %>
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
End Sub

'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
	If lgLngTotalInsertedRows > 0 Then
		Self.Returnvalue = DisplayMsgBox("900004", PopupParent.VB_YES_NO, "X", "X")
	Else
		Self.Returnvalue = vbNo
	End If	
End Sub

'========================================================================================================
'   Event Name : txtConClassCd_OnChange
'   Event Desc : Class����� Class ��ȿ���� Check �� �� Fetch
Function txtConClassCd_OnChange()
	lgStrCharCd1 = ""
	lgStrCharCd2 = ""
	
	With frm1
		If Trim(.txtConClassCd.value) <> "" Then
			If Not GetClassInfo() Then
				.txtConClassCd.value = ""
				.txtConClassNm.value = ""
				.txtConClassCd.focus
				Call SetReqAttr()
			End If
			txtConClassCd_OnChange = False
		Else
			.txtConClassNm.value = ""
			Call SetReqAttr()
		End If
	End With
	
End Function

Function txtConClassCd_OnKeyDown()
	lgBlnChgClass = True
End Function

'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		   
		If Row > 0 Then
			Select Case Col
				CASE C_UnitPopup
					.Col = C_Unit
					Call OpenSpreadPopup(Col, Row, .Text)

			End Select
		End If
	End With
	Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")
End Sub

'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
	Dim iStrData
	Dim iDblOldAmt, iDblQty, lDblAmt
	
	With frm1.vspdData
		.Row = Row
		.Col = 0

		.Col = Col	: iStrData = .Text
		
		If iStrData = "" Then Exit Sub
		
		Select Case Col
			Case C_Unit
				Call GetItemPrice(Row)
				
			Case C_Qty
				Call CalcAmt(Row, C_Qty)
				.Col = C_OldQty	: .Text = iStrData
				
			Case C_Price
				Call CalcAmt(Row, C_Price)
				
			Case C_Amt
				.Col = C_OldAmt	: iDblOldAmt = UNICDbl(.Text)
				.Text = iStrData
				Call CalcTotal(Row, 0, UNICDbl(iStrData) - iDblOldAmt)
		End Select
	End With

End Sub

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub    

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    	
	   	If lgPageNo <> "" Then
           Call DBQuery          
    	End If
    End If    
End Sub
	
'========================================================================================================
Sub FncSplitColumn()     
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)     
End Sub

'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	
	Call SetQuerySpreadColor(1)    

End Sub

'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
   On Error Resume Next
   If KeyAscii = 27 Then
	  Call CancelClick()
   End If
End Function

'========================================================================================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False													

	Err.Clear															

	' Ŭ������ ����� ��� ��ȿ�� check
	If lgBlnChgClass Then
		If Not GetClassInfo Then Exit Function
	End If

	Call ggoOper.ClearField(Document, "2")								
	Call InitVariables													

	If Not chkField(Document, "1") Then						
		Exit Function
	End If

	Call DbQuery()													

	FncQuery = True														
End Function

'========================================================================================================
Function DbQuery()
	Dim iStrVal

	DbQuery = False														

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	With frm1
		iStrVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
		iStrVal = iStrVal & "&txtClassCd=" & Trim(.txtConClassCd.value)
		iStrVal = iStrVal & "&txtCharValueCd1=" & Trim(.txtConCharValueCd1.value)
		iStrVal = iStrVal & "&txtCharValueCd2=" & Trim(.txtConCharValueCd2.value)
		iStrVal = iStrVal & "&txtPlantCd=" & Trim(.txtConPlantCd.value)
		iStrVal = iStrVal & "&txtSoldToParty=" & lgObjCaller.frm1.txtSoldToParty.value
		iStrVal = iStrVal & "&txtCurrency=" & lgObjCaller.frm1.txtCurrency.value
		iStrVal = iStrVal & "&txtDealType=" & lgObjCaller.frm1.txtHDealType.value
		iStrVal = iStrVal & "&txtPayMeth=" & lgObjCaller.frm1.txtHPayTermsCd.value
		iStrVal = iStrVal & "&txtSoDt=" & lgObjCaller.frm1.txtSoDt.value
	End With

	Call RunMyBizASP(MyBizASP, iStrVal)									

	DbQuery = True														
End Function

'========================================================================================================
Function DbQueryOk()													
	
	Dim iHiddenType
		
	lgIntFlgMode = PopupParent.OPMD_UMODE											
	Call SetQuerySpreadColor(1)
		
	frm1.vspdData.Focus
End Function

'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")
	
    gMouseClickStatus = "SPC"    
    
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
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

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub  

'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>Ŭ����</TD>
						<TD CLASS="TD6"><INPUT NAME="txtConClassCd" ALT="Ŭ����" TYPE="Text" MAXLENGTH=16 SiZE=16 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConClassCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopClassCd)">&nbsp;<INPUT NAME="txtConClassNm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
						<TD CLASS="TD5" NOWRAP>���1</TD>
						<TD CLASS="TD6"><INPUT NAME="txtConCharValueCd1" ALT="���1" TYPE="Text" MAXLENGTH=16 SiZE=16 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConCharValueCd1" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopCharValueCd1)">&nbsp;<INPUT NAME="txtConCharValueNm1" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>���2</TD>
						<TD CLASS="TD6"><INPUT NAME="txtConCharValueCd2" ALT="���2" TYPE="Text" MAXLENGTH=16 SiZE=16 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConCharValueCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopCharValueCd2)">&nbsp;<INPUT NAME="txtConCharValueNm2" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
						<TD CLASS="TD5" NOWRAP>����</TD>
						<TD CLASS="TD6"><INPUT NAME="txtConPlantCd" ALT="����" TYPE="Text" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopPlantCd)">&nbsp;<INPUT NAME="txtConPlantNm" TYPE="Text" MAXLENGTH="40" SIZE=25 tag="14"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>�Ѽ���</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3112ra9_fpDoubleSingle1_txtTotQty.js'></script>
						<TD CLASS=TD5 NOWRAP>�ѱݾ�</TD>
						<TD CLASS=TD6 NOWRAP>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<INPUT NAME="txtCurrency" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="14" ALT="ȭ��">&nbsp;
									</TD>
									<TD>
										<script language =javascript src='./js/s3112ra9_fpDoubleSingle2_txtTotAmt.js'></script>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_60%>>
				<TR>
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<script language =javascript src='./js/s3112ra9_vaSpread_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="Add"   NAME="Add"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="InsertRowsIntoSoDtl()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="Close" NAME="Close"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
			<script language =javascript src='./js/s3112ra9_OBJECT1_vspdData2.js'></script>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
