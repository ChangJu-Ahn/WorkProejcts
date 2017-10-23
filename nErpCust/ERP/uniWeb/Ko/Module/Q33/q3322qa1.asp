<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q3322QA1
'*  4. Program Name         : �ҷ�����������ȸ 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2004/08/06
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Lee Seung Wook
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit												'��: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgIsOpenPop                                             '��: Popupȭ���� ���� ���庯��               

'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID		= "Q3322qb1.asp"                 '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_JUMP_ID	= "Q2214qa1"                     '��: Cookie���� ����� ��� 
Const C_SHEETMAXROWS_D  = 100                            '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 2
'--------------- ������ coding part(��������,End)-------------------------------------------------------------

'--------------- ������ coding part(�������,Start)-----------------------------------------------------------
Dim CompanyYM
CompanyYM = UNIMonthClientFormat(UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gAPDateFormat))
'--------------- ������ coding part(�������,End)------------------------------------------------------------- 
                                                                  ' 1. Program id
'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	lgStrPrevKey     = ""
	lgBlnFlgChgValue = False
End Sub 

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtDtFr.Text = CompanyYM
	frm1.txtDtTo.Text = CompanyYM
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","QA") %>
End Sub

'===========================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : PopZAdoConfigGrid Reference Popup
'===========================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp", Array(ggoSpread.GetXMLData("A"), gMethodText), _
	         "dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		Exit Function
	Else
		Call ggoSpread.SaveXMLData("A", arrRet(0), arrRet(1))
		Call InitVariables
		Call InitSpreadSheet													'��: Initializes Spread Sheet 1
    End If
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = "����"						' �˾� ��Ī 
	arrParam(1) = "B_Plant"	' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		' Code Condition
	arrParam(3) = ""							' Name Condition
	arrParam(4) = ""
	arrParam(5) = "����"						' TextBox ��Ī 

    arrField(0) = "B_Plant.Plant_Cd"				' Field��(0)
    arrField(1) = "B_Plant.Plant_NM"				' Field��(1)
        
    arrHeader(0) = "�����ڵ�"						' Header��(0)
    arrHeader(1) = "�����"						' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	frm1.txtPlantCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
		frm1.txtPlantCd.Focus		
	End If	
	
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705","X","X","X") 		'���������� �ʿ��մϴ� 
		Exit Function
	End If
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
	
	arrParam(0) = "ǰ��"																	' �˾� ��Ī 
	arrParam(1) = "B_Item_By_Plant,B_Item"												' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtItemCd.Value)							' Code Condition
	arrParam(3) = ""												' Name Condition
	arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd"
	arrParam(4) = arrParam(4) & "  And B_Item_By_Plant.Plant_Cd = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & "" 			' Where Condition
	arrParam(5) = "ǰ��"																	' TextBox ��Ī 
	
	arrField(0) = "B_Item_By_Plant.Item_Cd"					' Field��(0)
	arrField(1) = "B_Item.Item_NM"				' Field��(1)
	arrField(2) = "B_Item.SPEC"					' Field��(2)
		
	arrHeader(0) = "ǰ���ڵ�"						' Header��(0)
	arrHeader(1) = "ǰ���"					' Header��(1)
	arrHeader(2) = "�԰�"						' Header��(2)
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	frm1.txtItemCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.Focus
	End If	
	
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenRoutNo()  -------------------------------------------------
'	Name : OpenRoutNo()
'	Description : RoutNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenRoutNo()

	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		lgIsOpenPop = False 
		Exit Function
	End If

	arrParam(0) = "����� �˾�"					' �˾� ��Ī 
	arrParam(1) = "P_ROUTING_HEADER"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtRoutNo.value)		' Code Condition
	arrParam(3) = ""								' Name Cindition
	If Trim(frm1.txtItemCd.value) <> "" Then
		arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
					" AND ITEM_CD = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S")
	Else
		arrParam(4) = "P_ROUTING_HEADER.PLANT_CD = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S")
	End if		
	arrParam(5) = "�����"			
	
    arrField(0) = "ED10" & parent.gcolsep & "ROUT_NO"							
    arrField(1) = "DESCRIPTION"
    arrField(2) = "ITEM_CD"													
    arrField(3) = "ED10" & parent.gcolsep & "BOM_NO"							
    arrField(4) = "ED10" & parent.gcolsep & "MAJOR_FLG"						
   
    arrHeader(0) = "�����"						
    arrHeader(1) = "����ø�"
    arrHeader(2) = "ǰ��"											
    arrHeader(3) = "BOM Type"					
    arrHeader(4) = "�ֶ����"				        
    
    arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=640px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
    lgIsOpenPop = False
    
    frm1.txtRoutNo.focus
	If arrRet(0) <> "" Then
		frm1.txtRoutNo.Value		= arrRet(0)		
		frm1.txtRoutNoDesc.Value	= arrRet(1)
	Else
		Exit Function
	End If		
	Set gActiveElement = document.activeElement
End Function


'------------------------------------------  OpenOprNo()  -------------------------------------------------
'	Name : OpenOprNo()
'	Description : OprNo PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenOprNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function    

	lgIsOpenPop = True
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		lgIsOpenPop = False 
		Exit Function
	End If

	arrParam(0) = "�����˾�"	
	arrParam(1) = "P_ROUTING_DETAIL A inner join P_WORK_CENTER B on A.wc_cd = B.wc_cd and A.plant_cd = B.plant_cd " & _
				  " left outer join B_MINOR C on A.job_cd = C.minor_cd and C.major_cd = " & FilterVar("P1006", "''", "S") & "" & _
				  " and A.rout_order in (" & FilterVar("F", "''", "S") & " ," & FilterVar("I", "''", "S") & " ) "				
	arrParam(2) = UCase(Trim(frm1.txtOprNo.Value))
	arrParam(3) = ""
	If (Trim(frm1.txtItemCd.value) <> "" AND Trim(frm1.txtRoutNo.value) <> "") THEN
		arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
					  " and	A.item_cd = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S") & _
					  " and	A.rout_no = " & FilterVar(UCase(frm1.txtRoutNo.value), "''", "S")
	ElseIf (Trim(frm1.txtItemCd.value) = "" AND Trim(frm1.txtRoutNo.value) <> "") THEN
		arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
					  " and	A.rout_no = " & FilterVar(UCase(frm1.txtRoutNo.value), "''", "S")
	ElseIf (Trim(frm1.txtItemCd.value) <> "" AND Trim(frm1.txtRoutNo.value) = "") THEN
		arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") & _
					  " and	A.item_cd = " & FilterVar(UCase(frm1.txtItemCd.value), "''", "S")
	Else 		
		arrParam(4) = "A.plant_cd = " & FilterVar(UCase(frm1.txtPlantCd.value), "''", "S") 
	End If	
	
	arrParam(5) = "����"			
	
	arrField(0) = "ED10" & parent.gcolsep & "A.OPR_NO"	
	arrField(1) = "ED15" & parent.gcolsep & "C.MINOR_NM"
	arrField(2) = "ED10" & parent.gcolsep & "A.ROUT_NO"
	arrField(3) = "A.ITEM_CD"
	arrField(4) = "ED10" & parent.gcolsep & "A.WC_CD"
	arrField(5) = "ED10" & parent.gcolsep & "A.INSIDE_FLG"
	arrField(6) = "ED10" & parent.gcolsep & "A.INSP_FLG"
	
	arrHeader(0) = "����"
	arrHeader(1) = "�����۾���"
	arrHeader(2) = "�����"
	arrHeader(3) = "ǰ��"		
	arrHeader(4) = "�۾���"	
	arrHeader(5) = "�系����"
	arrHeader(6) = "�˻翩��"	
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=640px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	frm1.txtOprNo.focus
	If arrRet(0) <> "" Then
		frm1.txtOprNo.Value	= arrRet(0)
		frm1.txtOprNoDesc.Value	= arrRet(1)
	Else
		Exit Function
	End If		
	Set gActiveElement = document.activeElement
	
End Function
    	
'------------------------------------------  OpenInspItem()  -------------------------------------------------
'	Name : OpenInspItem()
'	Description : Inspection Item By Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenInspItem()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705","X","X","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	'ǰ���ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtItemCd.Value) = "" then 
		Call DisplayMsgBox("229916","X","X","X")
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "ǰ�� �˻��׸�"						' �˾� ��Ī 
	arrParam(1) = "Q_Inspection_Standard_By_Item, Q_Inspection_Item"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtInspItemCd.Value)		' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "Q_Inspection_Standard_By_Item.Insp_Item_Cd = Q_Inspection_Item.Insp_Item_Cd"
	arrParam(4) = arrParam(4) & "  And Q_Inspection_Standard_By_Item.Plant_Cd = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & ""
	arrParam(4) = arrParam(4) & "  And Q_Inspection_Standard_By_Item.Item_Cd = " & FilterVar(frm1.txtItemCd.Value, "''", "S") & ""			' Where Condition
	arrParam(4) = arrParam(4) & "  And Q_Inspection_Standard_By_Item.insp_class_cd=" & FilterVar("P", "''", "S") & "   "
	arrParam(5) = "�˻��׸�"						' TextBox ��Ī 
	
	arrField(0) = "Q_Inspection_Standard_By_Item.INSP_ITEM_CD"							' Field��(0)
	arrField(1) = "Q_Inspection_Item.INSP_ITEM_NM"							' Field��(1)
	
	arrHeader(0) = "�˻��׸��ڵ�"						' Header��(0)
	arrHeader(1) = "�˻��׸��"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	frm1.txtInspItemCd.Focus
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtInspItemCd.Value = arrRet(0)
		frm1.txtInspItemNm.Value = arrRet(1)
		frm1.txtInspItemCd.Focus
	End If
	
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenDefectType()  -------------------------------------------------
'	Name : OpenDefectType()
'	Description : DefectType PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenDefectType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	'�����ڵ尡 �ִ� �� üũ 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
		
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "�ҷ������˾�"						' �˾� ��Ī 
	arrParam(1) = "Q_Defect_Type"					' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtDefectTypeCd.Value)		' Code Condition
	arrParam(3) = ""							' Name Cindition
	arrParam(4) = "Plant_Cd = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND Insp_Class_Cd = " & FilterVar("P", "''", "S") & " "
	arrParam(5) = "�ҷ�����"						' TextBox ��Ī 
	
	arrField(0) = "DEFECT_TYPE_CD"							' Field��(0)
	arrField(1) = "DEFECT_TYPE_NM"							' Field��(1)
	
	arrHeader(0) = "�ҷ������ڵ�"						' Header��(0)
	arrHeader(1) = "�ҷ�������"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	frm1.txtDefectTypeCd.Focus
	If Trim(arrRet(0)) = "" Then
		Exit Function
	Else
		frm1.txtDefectTypeCd.Value = Trim(arrRet(0))
		frm1.txtDefectTypeNm.Value = Trim(arrRet(1))
		frm1.txtDefectTypeCd.Focus
	End If
	
	Set gActiveElement = document.activeElement
End Function

'========================================= 2.6 InitSpreadSheet() =========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'==========================================================================================================
Sub InitSpreadSheet()
    Call AppendNumberPlace("6", "10", "2")
    Call SetZAdoSpreadSheet("Q3322QA1", "G", "A", "V20040806", parent.C_GROUP_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
    Call SetSpreadLock 
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
	Call ggoOper.FormatDate(frm1.txtDtFr, Parent.gDateFormat, 2)
	Call ggoOper.FormatDate(frm1.txtDtTo, Parent.gDateFormat, 2)

	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call SetToolbar("11000000000011")										'��: ��ư ���� ����	

	If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.value = UCase(Parent.gPlant)
	   	frm1.txtPlantNm.value = Parent.gPlantNm
	End IF
	frm1.txtPlantCd.focus
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode ) 
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"
	Set gActiveSpdSheet = frm1.vspdData
	Call SetPopupMenuItemInf("00000000001")
	
	If frm1.vspdData.MaxRows = 0 Then
 		Exit Sub
 	End If

 	If Row <= 0 Then
 		ggoSpread.Source = frm1.vspdData
 		ggoSpread.SSSort Col					'Sort in Ascending
 		lgSortKey = 2
	Else
 		'------ Developer Coding part (Start)
	 	'------ Developer Coding part (End)
 	End If
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
' Function Desc : 
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
    End If
End Sub 

'========================================================================================
' Function Name : vspdData_DblClick
' Function Desc : �׸��� �ش� ����Ŭ���� ���� ���� 
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
 	Dim iColumnName
        
 	If Row <= 0 Then
		Exit Sub
 	End If
    
  	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
 	End If
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)
End Sub

'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub

'========================================================================================
' Function Name : vspdData_ColWidthChange
' Function Desc : �׸��� ������ 
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
' Function Desc : �׸��� ��ġ ���� 
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
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
Sub PopRestoreSpreadColumnInf()	'###�׸��� ������ ���Ǻκ�###
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet
    Call InitComboBox
    Call ggoSpread.ReOrderingSpreadData
 	'------ Developer Coding part (Start)
 	'------ Developer Coding part (End)	
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	 
	'----------  Coding part  -----------------------------
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
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
	End If    
End Sub

'==========================================================================================
'   Event Name : txtDtFr
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtFr_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtFr.Action = 7
        frm1.txtDtFr.Focus 
	End If
End Sub

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtDtTo_DblClick(Button)
	If Button = 1 Then
		frm1.txtDtTo.Action = 7
        frm1.txtDtTo.Focus 
	End If
End Sub

'==========================================================================================
'   Event Name : txtDtFr
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtDtFr_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function txtDtTo_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtPlantCd_onChange()
	If Trim(frm1.txtPlantCd.Value) = "" Then
		frm1.txtPlantNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : txtDtTo
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtItemCd_onChange()
	If Trim(frm1.txtItemCd.Value) = "" Then
		frm1.txtItemNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : txtWcCd
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtWcCd_onChange()
	If Trim(frm1.txtWcCd.Value) = "" Then
		frm1.txtWcNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : txtInspItemCd
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtInspItemCd_onChange()
	If Trim(frm1.txtInspItemCd.Value) = "" Then
		frm1.txtInspItemNm.Value = ""
	End If
End Function

'==========================================================================================
'   Event Name : txtDefectTypeCd
'   Event Desc : Date OCX Double Click
'==========================================================================================
Function  txtDefectTypeCd_onChange()
	If Trim(frm1.txtDefectTypeCd.Value) = "" Then
		frm1.txtDefectTypeNm.Value = ""
	End If
End Function

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 
    Dim IntRetCD
    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
   
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
    
    If ValidDateCheck(frm1.txtDtFr, frm1.txtDtTo) = False Then
		Exit Function
    End If
    
    '�˻��׸��ڵ尡 ���� ���, ǰ���ڵ尡 �ִ� �� üũ 
    If Trim(frm1.txtInspItemCd.Value) <> "" Then
	    If Trim(frm1.txtItemCd.Value) = "" then 
			Call DisplayMsgBox("229916","X","X","X")
			Exit Function
		End If
	End If
   
	If Name_check("A") = False Then
		Set gActiveElement = document.activeElement
		Exit Function
	End If
    '-----------------------
    'Query function call area
    '-----------------------

	If DbQuery = False then
		Exit Function
	End If															'��: Query db data

    FncQuery = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
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
	Dim strDtFr
	Dim strDtTo
	Dim Yr
	Dim Mnth
	
	Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
	
	Yr = Left(frm1.txtDtFr.DateValue,4)
	Mnth = Mid(frm1.txtDtFr.DateValue,5, 2)
	strDtFr = Yr & Mnth
	
	Yr = Left(frm1.txtDtTo.DateValue,4)
	Mnth = Mid(frm1.txtDtTo.DateValue,5, 2)
	strDtTo = Yr & Mnth
	
	With frm1
		'--------------- ������ coding part(�������,Start)----------------------------------------------------------------------------
		strVal	= BIZ_PGM_ID & "?txtPlantCd=" & Trim(.txtPlantCd.value) _
				& "&txtDtFr=" & strDtFr _
				& "&txtDtTo=" & strDtTo _
				& "&txtItemCd=" & Trim(.txtItemCd.value) _
				& "&txtRoutNo=" & Trim(.txtRoutNo.value) _
				& "&txtOprNo=" & Trim(.txtOprNo.value) _
				& "&txtInspItemCd=" & Trim(.txtInspItemCd.value) _
				& "&txtDefectTypeCd=" & Trim(.txtDefectTypeCd.value)
		'--------------- ������ coding part(�������,End)------------------------------------------------------------------------------
		
		strVal	= strVal & "&lgStrPrevKey="   & lgStrPrevKey _
				& "&lgMaxCount="     & CStr(C_SHEETMAXROWS_D) _
				& "&lgSelectListDT=" & GetSQLSelectListDataType("A") _
				& "&lgTailList="     & MakeSQLGroupOrderByList("A") _
				& "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		Call RunMyBizASP(MyBizASP, strVal)
	    
	End With
	
	DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()
    '-----------------------
    'Reset variables area
    '-----------------------
    Call SetToolbar("11000000000111")							'��: ��ư ���� ���� 
	lgBlnFlgChgValue = False
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : Name_Check
'========================================================================================
Function Name_Check(ByVal Check)
	Name_Check = False

	With frm1
		If Check = "A" Then
			'-----------------------
			'Check Rout_No	 
			'-----------------------
			If Trim(.txtRoutNo.Value) <> "" Then
				If 	CommonQueryRs(" DESCRIPTION "," P_ROUTING_HEADER ", " ROUT_NO = " & FilterVar(.txtRoutNo.Value, "''", "S") & " AND PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
							
					lgF0 = Split(lgF0, Chr(11))
					.txtRoutNoDesc.Value = lgF0(0)
				Else
					.txtRoutNoDesc.Value = ""
					Call DisplayMsgBox("181300","X","X","X")
					.txtRoutNo.focus 
					Exit Function
				End If
			End If
			'-----------------------
			'Check Opr_No	 
			'-----------------------
			If Trim(.txtOprNo.Value) <> "" Then
				If 	CommonQueryRs(" B.MINOR_NM "," P_ROUTING_DETAIL A,B_MINOR B ", " A.JOB_CD = B.MINOR_CD AND B.MAJOR_CD = " & FilterVar("P1006", "''", "S") & "" & " AND A.OPR_NO = " & FilterVar(.txtOprNo.Value, "''", "S") & _
								"AND A.PLANT_CD = " & FilterVar(.txtPlantCd.Value, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					lgF0 = Split(lgF0, Chr(11))
					.txtOprNoDesc.Value = lgF0(0)
				End If
			End If
		End If
	End With
	
	Name_Check = True

End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL="NO" TABINDEX="-1">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�ҷ�����������ȸ</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<!--
					<TD WIDTH="*" align=right><button name="btnAutoSel" class="clsmbtn" ONCLICK="PopZAdoConfigGrid()">���ļ���</button></TD>
					-->
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>����</TD>
        									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="����" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
										<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE="20" MAXLENGTH=40 tag="14" ></TD>								
        									<TD CLASS="TD5" NOWRAP>�Ⱓ</TD>
									<TD CLASS="TD6">
										<script language =javascript src='./js/q3322qa1_fpDateTime5_txtDtFr.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/q3322qa1_fpDateTime6_txtDtTo.js'></script>										
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="ǰ��" tag="11XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnItemCd align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenItem()">
															<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=20 MAXLENGTH=20 tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRoutNo" SIZE=12 MAXLENGTH=20 tag="11XXXU" ALT="�����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenRoutNo()">&nbsp;<input TYPE=TEXT NAME="txtRoutNoDesc" SIZE="30" tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="11XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprNo()">&nbsp;<input TYPE=TEXT NAME="txtOprNoDesc" SIZE="30" tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>�˻��׸�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtInspItemCd" SIZE="10" MAXLENGTH="5" ALT="�˻��׸�" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnInspItem" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenInspItem()">
										<INPUT TYPE=TEXT NAME="txtInspItemNm" SIZE=20 MAXLENGTH="40" tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ҷ�����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDefectTypeCd" SIZE="10" MAXLENGTH="3" ALT="�ҷ�����" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDefectType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDefectType()">
										<INPUT TYPE=TEXT NAME="txtDefectTypeNm" SIZE=20 MAXLENGTH="40" tag="14" ></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>	
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
							<TD CLASS="TD5" NOWRAP>LOT���հݷ�����</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtLotRejUnit" SIZE="10" MAXLENGTH="10" ALT="LOT���հݷ�����" tag="24" STYLE="Text-Align: Center"></TD>
							<TD CLASS="TD5" NOWRAP>�ҷ�������</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDefectRatioUnit" SIZE="10" MAXLENGTH="10" ALT="�ҷ�������" tag="24" STYLE="Text-Align: Center"></TD>
						</TR>
						<TR>
								<TD HEIGHT=100% WIDTH=100% Colspan=4>
									<script language =javascript src='./js/q3322qa1_I199119478_vspdData.js'></script>
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
					<!--<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(1)">�����˻�ҷ�������ȸ</a></TD>-->
					<TD WIDTH="*" ALIGN="RIGHT">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" tabindex=-1 SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>

