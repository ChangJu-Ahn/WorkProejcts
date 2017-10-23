<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Inventory Management
'*  2. Function Name        : 
'*  3. Program ID           : I3112QA1
'*  4. Program Name         : ���������� 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2006/05/25
'*  8. Modified date(Last)  : 2006/05/25
'*  9. Modifier (First)     : KiHong Han
'* 10. Modifier (Last)      : KiHong Han
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit																'��: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "i3112qb1.asp"                 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP1_ID = "i3112ma1"
Const BIZ_PGM_JUMP2_ID = "i3111qa1"

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
Dim CompanyYMFr
Dim CompanyYMTo

CompanyYMTo = UNIMonthClientFormat(UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gAPDateFormat))
CompanyYMFr = UnIDateAdd("m", -11, CompanyYMTo, Parent.gDateFormat)
'--------------- ������ coding part(�������,End)------------------------------------------------------------- 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "I", "NOCOOKIE","QA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()	
	If ReadCookie("txtPlantCd") = "" Then
		If Parent.gPlant <> "" Then
			frm1.txtPlantCd.value = Ucase(Parent.gPlant)
			frm1.txtPlantNm.value = Parent.gPlantNm
		End If
    Else
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If	

	If ReadCookie("txtYYYYMM") = "" Then
		frm1.txtToYYYYMM.Text	= CompanyYMTo
		frm1.txtFrYYYYMM.Text	= CompanyYMFr
	Else
		frm1.txtToYYYYMM.Text = ReadCookie("txtYYYYMM")
		frm1.txtFrYYYYMM.Text = UnIDateAdd("m", -11, frm1.txtToYYYYMM.Text, Parent.gDateFormat)
	End If	
	
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtYYYYMM", ""
End Sub

Sub InitComboBox()
	'ABC FLAG SEARCH B_MINOR 2005-03-18 LSW
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("I1001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboABCFlg, lgF0, lgF0, Chr(11))
End Sub
'==========================================  2.2.6 InitSpreadSheet()  =======================================
'	Name : InitSpreadSheet()
'	Description : 
'========================================================================================================= 
Sub InitSpreadSheet()
	With frm1.vspdData
		.ReDraw = false
		.MaxCols = 0

		.MaxRows = 4
		
		.Col = 0
		.Row = 1
		.Text = "�Ǽ�������"
		.Row = 2
		.Text = "�Ǽ����ݾ�"
		.Row = 3
		.Text = "��⺸��������"
		.Row = 4
		.Text = "��⺸�����ݾ�"
		
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit
		
		.ColWidth(0) = 15
		.ColWidth(-1) = 12
		
		.Row = -1
		.Col = -1
		.TypeHAlign = 1
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		
		.ReDraw = true
	End With
End Sub

'==========================================  2.2.6 InitChartFx()  =======================================
'	Name : InitChartFx()
'	Description : Initialize ChartFx
'========================================================================================================= 
Sub InitChartFx()
	With frm1.ChartFX1
		'Chart Title �� Font ���� 
		'.Title_(2) = "���������̵�(����)"
		.LeftFont.Name = "����"
		.Axis(0).decimals = 4
		
		'���ʹڽ� ���̱� 
		.SerLegBox = True
		
		'Chart Series Legent Font ���� 
		.SerLegBoxObj.Font.Name = "����"
		
		'�׷����� GAP ���� 
		.TopGap = 5			'�׷����� ���� ���� ���� 
		.BottomGap = 20			'�׷����� �Ʒ��� ���� ���� 
		.RightGap = 5
		.LeftGap = 50
	End With	
End Sub

'==========================================  2.2.7 ClearChartFx()  =======================================
'	Name : ClearChartFx()
'	Description : Clear Chart FX Datas
'========================================================================================================= 
Sub ClearChartFx()
	With frm1.ChartFX1
		' X��/Y�� ���� �� ���� �Ⱥ��̰� �� 
		.Axis(2).Visible = False
		.Axis(0).Visible = False
		
		'���� Clear
		.ClearLegend 1
		
		'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
		.OpenDataEx 1, 1, 1
		.CloseData 1 Or &H800		'COD_VALUES Or COD_REMOVE
		
		'�迭�� �Ⱥ��̰� �� 
		.Series(0).Visible = False
		'.Title_(2) = "���������̵�(����)"
	End With
End Sub


'==========================================  ChangingChartDisplay()  =======================================
'	Name : ChangingChartDisplay()
'	Description : Clear Chart FX Datas
'========================================================================================================= 
Sub ChangingChartDisplay(Byval pvTarget)
	Dim iLngCol
	Dim iLngMaxCols
	
	iLngMaxCols = frm1.vspdData.MaxCols
	
	If iLngMaxCols <= 0 Then Exit Sub
	
	With frm1.ChartFX1
		'-------------------------
		' �ʱ�ȭ 
		'-------------------------
		' X��/Y�� ���� �� ���� ���̰� �� 
		.Axis(2).Visible = True
		.Axis(0).Visible = True
		
		'���� Clear
		.ClearLegend 1
		
		'��Ʈ FX���� ������ ä�� �ʱ�ȭ 
		.OpenDataEx 1, 1, 1
		.CloseData 1 Or &H800		'COD_VALUES Or COD_REMOVE
		
		'-------------------------
		' ���� 
		'-------------------------
		If pvTarget = "1" Then
			'���� 
			.Title_(2) = "���������̵�(����)"
			.Axis(0).Decimals = ggQty.DecPoint
			.Axis(0).Format = 1		' AF_NUMBER
			
			' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points " 
			.SerLeg(0) = "�Ǽ�������"
			.SerLeg(1) = "��⺸��������"
			
			.OpenDataEx 1, 2, iLngMaxCols					'��Ʈ FX���� ������ ä�� �����ֱ� 

			For iLngCol = 0 to iLngMaxCols - 1
				frm1.vspdData.Col = iLngCol + 1
				
				'X�� �� 
				frm1.vspdData.Row = 0
				.Axis(2).Label(iLngCol) = frm1.vspdData.Text
				
				'�Ǽ������� 
				frm1.vspdData.Row = 1
				If frm1.vspdData.Text = "" Then
						.Series(0).Yvalue(iLngCol) = 1E+308				'CHART_HIDDEN
				Else
						.Series(0).Yvalue(iLngCol) =  parent.UNICDbl(frm1.vspdData.Text)
				End If
			Next
			
			.Series(0).Visible = True 
			frm1.vspdData.Row = 3
			
			For iLngCol = 0 to iLngMaxCols - 1
				frm1.vspdData.Col = iLngCol + 1
				
				'��⺸�������� 
				frm1.vspdData.Row = 3
				If frm1.vspdData.Text = "" Then
						.Series(1).Yvalue(iLngCol) = 1E+308				'CHART_HIDDEN
				Else
						.Series(1).Yvalue(iLngCol) =  parent.UNICDbl(frm1.vspdData.Text)
				End If
			Next

			.Series(1).Visible = True

			' Close the VALUES channel
			.CloseData 1		'COD_VALUES
			
			' Y�� ������ ���� 
			.RecalcScale

		Else
			'�ݾ� 
			.Title_(2) = "���������̵�(�ݾ�)"
			.Axis(0).Decimals = ggAmtOfMoney.DecPoint
			.Axis(0).Format = 1		' AF_NUMBER
			
			' Open the VALUES channel specifying "nSeries" Series and "nPoints" Points " 
			.SerLeg(0) = "�Ǽ����ݾ�"
			.SerLeg(1) = "��⺸�����ݾ�"
			
			.OpenDataEx 1, 2, iLngMaxCols					'��Ʈ FX���� ������ ä�� �����ֱ� 

			For iLngCol = 0 to iLngMaxCols - 1
				frm1.vspdData.Col = iLngCol + 1
				
				'X�� �� 
				frm1.vspdData.Row = 0
				.Axis(2).Label(iLngCol) = frm1.vspdData.Text
				
				'�Ǽ������� 
				frm1.vspdData.Row = 2
				If frm1.vspdData.Text = "" Then
						.Series(0).Yvalue(iLngCol) = 1E+308				'CHART_HIDDEN
				Else
						.Series(0).Yvalue(iLngCol) =  parent.UNICDbl(frm1.vspdData.Text)
				End If
			Next
			
			.Series(0).Visible = True
			
			'��⺸�������� 
			frm1.vspdData.Row = 4
			For iLngCol = 0 to iLngMaxCols - 1
				frm1.vspdData.Col = iLngCol + 1
				
				If frm1.vspdData.Text = "" Then
						.Series(1).Yvalue(iLngCol) = 1E+308				'CHART_HIDDEN
				Else
						.Series(1).Yvalue(iLngCol) =  parent.UNICDbl(frm1.vspdData.Text)
				End If
			Next

			.Series(1).Visible = True

			' Close the VALUES channel
			.CloseData 1		'COD_VALUES
			
			' Y�� ������ ���� 
			.RecalcScale
		End If
	End With
End Sub
'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "�����˾�"						' �˾� ��Ī 
	arrParam(1) = "B_Plant"								' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Condition
	arrParam(4) = ""
	arrParam(5) = "����"							' TextBox ��Ī 

    arrField(0) = "Plant_Cd"					' Field��(0)
    arrField(1) = "Plant_NM"					' Field��(1)
        
    arrHeader(0) = "����"						' Header��(0)
    arrHeader(1) = "�����"							' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	frm1.txtPlantCd.Focus
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = Trim(arrRet(0))
		frm1.txtPlantNm.Value = Trim(arrRet(1))
	End If	
	
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenQueryTarget()  -------------------------------------------------
'	Name : OpenQueryTarget()
'	Description : OpenQueryTarget PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenQueryTarget()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strQueryTargetClass

	If frm1.rdoQueryTargetClass.rdoQueryTargetClass1.Checked = True Then
		strQueryTargetClass = "1"
	ElseIf frm1.rdoQueryTargetClass.rdoQueryTargetClass2.Checked = True Then
		strQueryTargetClass = "2"
	Else
		strQueryTargetClass = "3"
	End If
	
	Select Case strQueryTargetClass
		Case "1"
			'�����ڵ尡 �ִ� �� üũ 
			If Trim(frm1.txtPlantCd.Value) = "" then 
				Call DisplayMsgBox("220705","X","X","X") 		'���������� �ʿ��մϴ� 
				Exit Function
			End If
	
			arrParam(0) = "ǰ��"													' �˾� ��Ī 
			arrParam(1) = "B_Item_By_Plant,B_Item"									' TABLE ��Ī 
			arrParam(2) = Trim(frm1.txtQueryTargetCd.Value)							' Code Condition
			arrParam(3) = ""														' Name Condition
			arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd"
			arrParam(4) = arrParam(4) & "  And B_Item_By_Plant.Plant_Cd = '" & FilterVar(Trim(frm1.txtPlantCd.Value),"''","SNM") & "'" 			' Where Condition
			arrParam(5) = "ǰ��"													' TextBox ��Ī 
	
			arrField(0) = "B_Item_By_Plant.Item_Cd"		' Field��(0)
			arrField(1) = "B_Item.Item_NM"				' Field��(1)
			arrField(2) = "B_Item.SPEC"					' Field��(2)
				
			arrHeader(0) = "ǰ��"					' Header��(0)
			arrHeader(1) = "ǰ���"						' Header��(1)
			arrHeader(2) = "�԰�"						' Header��(2)
		Case "2"
			arrParam(0) = "ǰ��׷��˾�"	
			arrParam(1) = "B_ITEM_GROUP"				
			arrParam(2) = Trim(frm1.txtQueryTargetCd.Value)
			arrParam(3) = ""
			arrParam(4) = "DEL_FLG = 'N' " 			
			arrParam(5) = "ǰ��׷�"			
	
			arrField(0) = "ITEM_GROUP_CD"	
			arrField(1) = "ITEM_GROUP_NM"	
    
			arrHeader(0) = "ǰ��׷�"		
			arrHeader(1) = "ǰ��׷��"
		Case "3"
			Exit Function
		
	End Select
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	frm1.txtQueryTargetCd.Focus
	If arrRet(0) <> "" Then
		frm1.txtQueryTargetCd.Value = Trim(arrRet(0))
		frm1.txtQueryTargetNm.Value = Trim(arrRet(1))
	End If	
	
	Set gActiveElement = document.activeElement
End Function

'=============================================  2.5.2 JumpToLongtermInvAnal()  ======================================
'=	Event Name : JumpToLongtermInvAnal
'=	Event Desc : ������м����� Jump
'========================================================================================================
Function JumpToLongtermInvAnal()
	With frm1
		'�����ڵ�/��/�м����� 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtYYYYMM", .txtToYYYYMM.Text
	End With
	
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'=============================================  2.5.2 JumpToLongtermInvList()  ======================================
'=	Event Name : JumpToLongtermInvChange
'=	Event Desc : ��������Ȳ�� Jump
'========================================================================================================
Function JumpToLongtermInvList()
	With frm1
		'�����ڵ�/��/�м����� 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtYYYYMM", .txtToYYYYMM.Text
	End With
	
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
		
	Call ggoOper.LockField(Document, "N")
	Call ggoOper.FormatDate(frm1.txtFrYYYYMM, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtToYYYYMM, Parent.gDateFormat, 2)
	
	Call InitVariables	
	Call SetDefaultVal
	
	Call InitComboBox
	'Call InitChartFX
	Call InitSpreadSheet
		
	Call SetToolbar("11000000000011")										'��: ��ư ���� ����	
'--------------- ������ coding part(�������,Start)----------------------------------------------------
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.focus 
    Else
		frm1.txtFrYYYYMM.focus
	End If
'--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'=======================================================================================================
'   Event Name : txtFrYYYYMM_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFrYYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtFrYYYYMM.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtFrYYYYMM.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFrYYYYMM_KeyPress(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtFrYYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtToYYYYMM_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToYYYYMM_DblClick(Button)
    If Button = 1 Then
        frm1.txtToYYYYMM.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtToYYYYMM.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToYYYYMM_KeyPress(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToYYYYMM_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'==========================================================================================
'   Event Name : txtPlantCd_onChange
'   Event Desc : 
'==========================================================================================
Function  txtPlantCd_onChange()
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : txtQueryTargetCd_onChange
'   Event Desc : 
'==========================================================================================
Function  txtQueryTargetCd_onChange()
	If Trim(frm1.txtQueryTargetCd.Value) = "" Then
		frm1.txtQueryTargetNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : rdoQueryTargetClass1_onClick
'   Event Desc : 
'==========================================================================================
Function  rdoQueryTargetClass1_onClick()
	frm1.cboABCFlg.value = ""
	frm1.txtQueryTargetCd.value = ""
	
	Call ggoOper.SetReqAttr(frm1.txtQueryTargetCd, "D")
	Call ggoOper.SetReqAttr(frm1.cboABCFlg, "Q")
End Function

'==========================================================================================
'   Event Name : rdoQueryTargetClass2_onClick
'   Event Desc : 
'==========================================================================================
Function  rdoQueryTargetClass2_onClick()
	frm1.cboABCFlg.value = ""
	frm1.txtQueryTargetCd.value = ""
	
	Call ggoOper.SetReqAttr(frm1.txtQueryTargetCd, "D")
	Call ggoOper.SetReqAttr(frm1.cboABCFlg, "Q")
	
	
End Function

'==========================================================================================
'   Event Name : rdoQueryTargetClass3_onClick
'   Event Desc : 
'==========================================================================================
Function  rdoQueryTargetClass3_onClick()
	frm1.cboABCFlg.value = ""
	frm1.txtQueryTargetCd.value = ""
	
	Call ggoOper.SetReqAttr(frm1.txtQueryTargetCd, "Q")
	Call ggoOper.SetReqAttr(frm1.cboABCFlg, "D")
End Function

'==========================================================================================
'   Event Name : rdoChartTarget1_onClick
'   Event Desc : 
'==========================================================================================
Function  rdoChartTarget1_onClick()
	Call ChangingChartDisplay("1")
End Function

'==========================================================================================
'   Event Name : rdoChartTarget2_onClick
'   Event Desc : 
'==========================================================================================
Function  rdoChartTarget2_onClick()
	Call ChangingChartDisplay("2")
End Function

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 

	FncQuery = False                                                        '��: Processing is NG

	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then	Exit Function
	
	'-----------------------
    '��ȸ�Ⱓ üũ 
    '-----------------------
	If ValidDateCheck(frm1.txtFrYYYYMM, frm1.txtToYYYYMM) = False Then
		Exit Function
	End If
	
	' Clear & Change Contents Area 
	frm1.rdoChartTarget.rdoChartTarget1.Checked = True
	
	frm1.txtLongtermStockCalPeriod.value = ""
	frm1.txtPerniciousStockCalPeriod.value = ""
	
    'Call ClearChartfx
	
	frm1.vspdData.MaxCols = 0
	
	Call InitVariables	
	
	'-----------------------
	'Query function call area
	'----------------------- 
	If DbQuery = False then Exit Function
	
	FncQuery = True
End Function	

'========================================================================================
' Function Name : FncFind
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
    Call parent.FncPrint()                                              
    FncPrint = True
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================
Function FncPrev() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
Function FncNext() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel()
	FncExcel = False 
	Call parent.FncExport(Parent.C_MULTI)
	FncExcel = True 
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	DbQuery = False
	
	Dim strVal
	Dim strQueryTargetClass
    Dim strQueryTargetCd
    
	Call LayerShowHide(1)
	
    With frm1
			
		If .rdoQueryTargetClass.rdoQueryTargetClass1.Checked = True Then
			strQueryTargetClass = "1"
			strQueryTargetCd = Trim(.txtQueryTargetCd.value)
		ElseIf .rdoQueryTargetClass.rdoQueryTargetClass2.Checked = True Then
			strQueryTargetClass = "2"
			strQueryTargetCd = Trim(.txtQueryTargetCd.value)
		Else
			strQueryTargetClass = "3"
			strQueryTargetCd = Trim(.cboABCFlg.value)
			
		End If
							
		strVal = BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value) _
							& "&txtFrYr=" & Left(.txtFrYYYYMM.DateValue, 4) _
							& "&txtFrMnth=" & Mid(.txtFrYYYYMM.DateValue, 5, 2) _
							& "&txtToYr=" & Left(.txtToYYYYMM.DateValue, 4) _
							& "&txtToMnth=" & Mid(.txtToYYYYMM.DateValue, 5, 2) _
							& "&txtQueryTargetClass=" & strQueryTargetClass _
							& "&txtQueryTargetCd=" & strQueryTargetCd
    End With

    Call RunMyBizASP(MyBizASP, strVal)	
    
    DbQuery = True
                                                          					'��: Processing is NG
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
	'-----------------------
    'Reset variables area
    '-----------------------
    Call DrawEBChart
    Call SetToolbar("11000000000111")
	Set gActiveElement = document.activeElement
End Function

'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================

Function SetPrintCond(StrEbrFile, strUrl, intChartNo)

	SetPrintCond = False

	StrEbrFile	= "I3112QA1"

	StrUrl = ""

	SetPrintCond = True

End Function


Function DrawEBChart()
	Dim StrUrl, StrEbrFile, ObjName

	If Not SetPrintCond(StrEbrFile, strUrl, 1) Then
		Exit Function
	End If

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	EBActionA.menu.value = 0
    Call FncEBR5RC2(ObjName, "view", StrUrl,EBActionA,"EBR")
End Function 

Function MyBizASP1_onReadyStateChange()
		If LCase(MyBizASP1.Document.ReadyState) = "complete" Then
			Call LayerShowHide(0)
		End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<!-- SPACE AREA-->
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<!-- TAB, REFERENCE AREA -->
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    	</TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<!-- CONDITION, CONTENT AREA -->
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<!-- SAPCE AREA -->
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<!-- CONDITION AREA -->
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
        							<TD CLASS="TD6" NOWRAP>
        								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="����" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH=40 tag="14"></TD>								
        							<TD CLASS="TD5" NOWRAP>�Ⱓ</TD>
									<TD CLASS="TD6">
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtFrYYYYMM name=txtFrYYYYMM CLASS=FPDTYYYYMM title=FPDATETIME ALT="�Ⱓ(FROM)" tag="12"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=txtToYYYYMM name=txtToYYYYMM CLASS=FPDTYYYYMM title=FPDATETIME ALT="�Ⱓ(TO)" tag="12"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��ȸ���з�</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryTargetClass" TAG="1X" ID="rdoQueryTargetClass1" CHECKED><LABEL FOR="rdoQueryTargetClass1">ǰ��</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryTargetClass" TAG="1X" ID="rdoQueryTargetClass2"><LABEL FOR="rdoQueryTargetClass2">ǰ��׷�</LABEL>
										<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoQueryTargetClass" TAG="1X" ID="rdoQueryTargetClass3"><LABEL FOR="rdoQueryTargetClass3">ABC����</LABEL>
									</TD>
									<TD CLASS="TD5" NOWRAP>��ȸ���</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtQueryTargetCd" SIZE=15 MAXLENGTH=18 ALT="��ȸ���" tag="11XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnQueryTarget align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenQueryTarget()">
										<INPUT TYPE=TEXT NAME="txtQueryTargetNm" SIZE=20 MAXLENGTH=20 tag="14">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ��ABC����</TD>
									<TD CLASS=TD6 NOWRAP>
										<SELECT NAME="cboABCFlg" ALT="ǰ��ABC����" STYLE="Width: 98px;" tag="14"><option VALUE></option></SELECT>
									</TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP></TD>		
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>   
				</TR>      
				<!-- SAPCE AREA -->
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<!-- CONTENT AREA(MULTI) -->
				<TR>      
					<!--<TD WIDTH=100% HEIGHT=44% valign=top>      -->
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>  
							    <TD WIDTH=100% HEIGHT=5 valign=top>  
									<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>   
										<TR>
											<TD CLASS="TD5" NOWRAP>íƮ���</TD>
											<TD CLASS="TD6" NOWRAP>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoChartTarget" TAG="2X" ID="rdoChartTarget1" CHECKED><LABEL FOR="rdoChartTarget1">����</LABEL>
												<INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoChartTarget" TAG="2X" ID="rdoChartTarget2"><LABEL FOR="rdoChartTarget2">�ݾ�</LABEL>
											</TD>
											<TD CLASS="TD5" NOWRAP></TD>
											<TD CLASS="TD6" NOWRAP></TD>
										</TR>
										<TR>
											<TD CLASS="TD5" NOWRAP>��⺸�������رⰣ</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLongtermStockCalPeriod" SIZE="5" ALT="��⺸�������رⰣ" tag="24" STYLE="Text-Align: Center">&nbsp;���� �ʰ�</TD>
											<TD CLASS="TD5" NOWRAP>�Ǽ������رⰣ</TD>
											<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPerniciousStockCalPeriod" SIZE="5" ALT="�Ǽ������رⰣ" tag="24" STYLE="Text-Align: Center">&nbsp;���� �ʰ�</TD>
										</TR>      
									</TABLE>      
								</TD>      
							</TR>
							<TR>      
								<TD WIDTH=100% HEIGHT=75% valign=top>      
									<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>      
										<TR>      
											<TD HEIGHT="100%">      
												<IFRAME NAME="MyBizASP1"  WIDTH=100% HEIGHT=100% FRAMEBORDER=0 SCROLLING=AUTO framespacing=0 marginwidth=0 marginheight=0 ></IFRAME>      
											</TD>      
										</TR>      
									</TABLE>      
								</TD>      
							</TR>      
							<TR>      
								<TD HEIGHT=5 WIDTH=100%></TD>      
							</TR>      
							<TR>      
								<TD WIDTH=100% HEIGHT=* valign=top>      
									<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>      
										<TR>      
											<TD HEIGHT="100%">      
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="24" TITLE="SPREAD" id=OBJECT1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>      
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
	<!-- SPACE AREA -->
	<TR>
		<TD <%=HEIGHT_TYPE_01%>> </TD>
	</TR>
	<!-- BATCH,JUMP AREA -->
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpToLongtermInvAnal">������м�</A>&nbsp;|&nbsp;<A href="vbscript:JumpToLongtermInvList">��������Ȳ</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<!-- IFRAME AREA -->
	<TR>      
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  tabindex=-1 SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>      
		</TD>      
	</TR>      
</TABLE>      
</FORM>
<FORM NAME="EBActionA" ID="EBActionA" TARGET="MyBizASP1" METHOD="POST"  scroll=yes> 
	<input TYPE="HIDDEN" NAME="menu" value=0 > 
	<input TYPE="HIDDEN" NAME="id" > 
	<input TYPE="HIDDEN" NAME="pw" >
	<input TYPE="HIDDEN" NAME="doc" > 
	<input TYPE="HIDDEN" NAME="form" > 
	<input TYPE="HIDDEN" NAME="runvar" >
</FORM>      
<DIV ID="MousePT" NAME="MousePT">      
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>      
</DIV>      
</BODY>      
</HTML>

