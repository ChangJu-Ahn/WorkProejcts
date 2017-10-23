
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           :  p3112ma2.asp
'*  4. Program Name         :  Production Configuration
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/11/28
'*  8. Modified date(Last)  : 2003/06/18
'*  9. Modifier (First)     : Jung Yu Kyung
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<Script LANGUAGE = "VBScript">

Option Explicit

Const BIZ_PGM_QRY_ID = "p3112mb3.asp"            '��: �����Ͻ� ���� ASP�� 
Const BIZ_PGM_SAVE_ID = "p3112mb4.asp"            '��: �����Ͻ� ���� ASP�� 
 
<!-- #Include file="../../inc/lgVariables.inc" -->

Dim lgBlnFlgConChg		'��: Condition ���� Flag
Dim lgPlantCd			'��: PlantCd���� ����Ǿ����� ���ϴ� ���� 

Dim lgRdoOldVal1
Dim lgRdoOldVal2
Dim lgRdoOldVal3
Dim lgRdoOldVal4
Dim lgRdoOldVal5
Dim lgRdoOldVal6
Dim lgRdoOldVal7
Dim lgRdoOldVal8
Dim lgRdoOldVal9
Dim lgRdoOldVal10
Dim lgRdoOldVal11
Dim lgRdoOldVal12
Dim lgRdoOldVal13

'Add 2005-03-07/ eng_bom_flag
Dim lgRdoOldVal14 
'Add 2005-09-20/ 
Dim lgRdoOldVal15
'Add 2006-04-20/ backlog_flg
Dim lgRdoOldVal16

'Add 2006-07-20/ prod_child_flg
Dim lgRdoOldVal17
Dim lgRdoOldVal18
 
Dim IsOpenPop

'=========================================================================================================
' Name : InitVariables()
' Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
	lgIntFlgMode = parent.OPMD_CMODE                                        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
	lgIntGrpCount = 0                                                       '��: Initializes Group View Size
	'----------  Coding part  -------------------------------------------------------------
	IsOpenPop = False              '��: ����� ���� �ʱ�ȭ 
End Sub

'=========================================================================================================
' Name : SetDefaultVal()
' Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
 Sub SetDefaultVal()
	With frm1
		.rdoRlsInvChkFlg1.checked = True
		lgRdoOldVal1 = 1
		.rdoAutoRcptFlg1.checked = True
		lgRdoOldVal2 = 1
		.rdoPreOprChkFlg1.checked = True
		lgRdoOldVal3 = 1
		'.rdoBkupClsOrdFlg1.checked = True
		'lgRdoOldVal4 = 1
		.rdoProdEtcMthd2.checked = True
		lgRdoOldVal5 = 2  
		.rdoOrdCloseMthd2.checked = True
		lgRdoOldVal6 = 2
		.rdoExssRcptFlg2.checked = True 
		lgRdoOldVal7 = 2
		.rdoProdMthd3.checked = True
		lgRdoOldVal8 = 3
		.rdoProdFlg2.checked = True
		lgRdoOldVal9 = 2
		.rdoMPSMETHOD1.checked = True
		lgRdoOldVal10 = 1
		.rdoDELIVERYORDERFLG2.checked = True
		lgRdoOldVal11 = 2
		.rdoBOMHISTORYFLG2.checked = True
		lgRdoOldVal12 = 2
		.rdoRoutingLTFlg2.checked = True
		lgRdoOldVal13 = 2
		'Add 2005-03-07/ eng_bom_flag
		.rdoENGBOMFLG2.checked = True
		lgRdoOldVal14 = 2
		'Add 2005-09-17/ opr_cost_flag
		.rdoOprCostFlg2.checked = True
		lgRdoOldVal15 = 2
		'Add 2006-09417/ backlog_flag
		.rdoBacklogFlg2.checked = True
		lgRdoOldVal16 = 2
		'Add 2006-07-17/ prod_child_flag
		.rdoProdChildMthd1.checked = True
		lgRdoOldVal17 = 1
		.rdoProdRscMthd1.checked = True
		lgRdoOldVal18 = 1
	End With
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
' Name : OpenPlant()
' Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd1.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�" 
	arrParam(1) = "B_PLANT"    
	arrParam(2) = Trim(UCase(frm1.txtPlantCd1.Value))
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
 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If 
 
End Function

'++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPlant()  --------------------------------------------------
' Name : SetPlant()
' Description : Plant Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlant(byval arrRet)
 frm1.txtPlantCd1.Value    = arrRet(0)  
 frm1.txtPlantNm1.Value    = arrRet(1)
 frm1.txtPlantCd1.focus   
End Function

'==========================================  3.1.1 Form_Load()  ======================================
' Name : Form_Load()
' Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
   
	Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
	    
	'----------  Coding part  -------------------------------------------------------------
	Call SetToolBar("110010000000111")
	Call ggoOper.LockField(Document, "N")           '��: Lock  Suitable  Field
 
	If parent.gPlant <> "" Then
		frm1.txtPlantCd1.value = parent.gPlant
		frm1.txtPlantNm1.value = parent.gPlantNm
  
		Call InitVariables
		Call MainQuery
		frm1.txtPlantCd1.focus
	Else
		frm1.txtPlantCd1.value = ""
		frm1.txtPlantNm1.value = ""
		frm1.txtPlantCd1.focus
		Call SetDefaultVal()
		Call InitVariables()
	End If
 
	 
	Set gActiveElement = document.activeElement 
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

Sub rdoAutoRcptFlg1_OnClick()
	If lgRdoOldVal2 = 1 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal2 = 1
End Sub 

Sub rdoAutoRcptFlg2_OnClick()
	If lgRdoOldVal2 = 2 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal2 = 2
End Sub 

Sub rdoPreOprChkFlg1_OnClick()
	If lgRdoOldVal3 = 1 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal3 = 1
End Sub 

Sub rdoPreOprChkFlg2_OnClick()
	If lgRdoOldVal3 = 2 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal3 = 2
End Sub 

Sub rdoProdEtcMthd1_OnClick()
	If lgRdoOldVal5 = 1 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal5 = 1
End Sub 

Sub rdoProdEtcMthd2_OnClick()
	If lgRdoOldVal5 = 2 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal5 = 2
End Sub 

Sub rdoProdFlg1_OnClick()
	If lgRdoOldVal9 = 1 Then Exit Sub
		 
	lgBlnFlgChgValue = True
	lgRdoOldVal9 = 1
End Sub 

Sub rdoProdFlg2_OnClick()
	If lgRdoOldVal9 = 2 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal9 = 2
End Sub 

Sub rdoRlsInvChkFlg1_OnClick()
	If lgRdoOldVal1 = 1 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal1 = 1
End Sub 

Sub rdoRlsInvChkFlg2_OnClick()
	If lgRdoOldVal1 = 2 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal1 = 2
End Sub

Sub rdoRlsInvChkFlg3_OnClick()
	If lgRdoOldVal1 = 3 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal1 = 3
End Sub  

Sub rdoOrdCloseMthd1_OnClick()
	If lgRdoOldVal6 = 1 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal6 = 1
End Sub 

Sub rdoOrdCloseMthd2_OnClick()
	If lgRdoOldVal6 = 2 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal6 = 2
End Sub 
    
Sub rdoExssRcptFlg1_OnClick()
	If lgRdoOldVal7 = 1 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal7 = 1
End Sub 

Sub rdoExssRcptFlg2_OnClick()
	If lgRdoOldVal7 = 2 Then Exit Sub
	 
	lgBlnFlgChgValue = True
	lgRdoOldVal7 = 2
End Sub 

Sub rdoProdMthd1_OnClick()
	If lgRdoOldVal8 = 1 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal8 = 1
End Sub 

Sub rdoProdMthd2_OnClick()
	If lgRdoOldVal8 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal8 = 2
End Sub 

Sub rdoProdMthd3_OnClick()
	If lgRdoOldVal8 = 3 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal8 = 3
End Sub 

Sub rdoProdMthd4_OnClick()
	If lgRdoOldVal8 = 4 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal8 = 4
End Sub 

Sub rdoProdMthd5_OnClick()
	If lgRdoOldVal8 = 5 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal8 = 5
End Sub 

Sub rdoMPSMETHOD1_OnClick()
	If lgRdoOldVal10 = 1 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal10 = 1
End Sub 

Sub rdoMPSMETHOD2_OnClick()
	If lgRdoOldVal10 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal10 = 2
End Sub 

Sub rdoMPSMETHOD3_OnClick()
	If lgRdoOldVal10 = 3 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal10 = 3
End Sub 

Sub rdoDELIVERYORDERFLG1_OnClick()
	If lgRdoOldVal11 = 1 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal11 = 1
End Sub 

Sub rdoDELIVERYORDERFLG2_OnClick()
	If lgRdoOldVal11 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal11 = 2
End Sub 

Sub rdoBOMHISTORYFLG1_OnClick()
	If lgRdoOldVal12 = 1 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal12 = 1
End Sub 

Sub rdoBOMHISTORYFLG2_OnClick()
	If lgRdoOldVal12 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal12 = 2
End Sub 

Sub rdoRoutingLTFlg1_OnClick()
	If lgRdoOldVal13 = 1 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal13 = 1
End Sub 

Sub rdoRoutingLTFlg2_OnClick()
	If lgRdoOldVal13 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal13 = 2
End Sub 

'Add 2005-03-07/ eng_bom_flag
Sub rdoENGBOMFLG1_OnClick()
	If lgRdoOldVal14 = 1 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal14 = 1
End Sub 

'Add 2005-03-07/ eng_bom_flag
Sub rdoENGBOMFLG2_OnClick()
	If lgRdoOldVal14 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal14 = 2
End Sub 

'Add 2005-09-17/ opr_cost_flag
Sub  rdoOprCostFlg1_OnClick()
	If lgRdoOldVal15 = 1 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal15 = 1
End Sub 

'Add 2005-09-17/ opr_cost_flag
Sub rdoOprCostFlg2_OnClick()
	If lgRdoOldVal15 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal15 = 2
End Sub 

'Add 2006-04-17/ backlog_flg
Sub rdoBacklogFlg1_OnClick()
	If lgRdoOldVal16 = 1 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal16 = 1
End Sub 


'Add 2006-04-17/ backlog_flg
Sub rdoBacklogFlg2_OnClick()
	If lgRdoOldVal16 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal16 = 2
End Sub 


'Add 2006-07-18
Sub rdoProdChildMthd1_OnClick()
	If lgRdoOldVal17 = 1 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal17 = 1
End Sub 

Sub rdoProdChildMthd2_OnClick()
	If lgRdoOldVal17 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal17 = 2
End Sub 

Sub rdoProdChildMthd3_OnClick()
	If lgRdoOldVal17 = 3 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal17 = 3
End Sub 

Sub rdoProdChildMthd4_OnClick()
	If lgRdoOldVal17 = 4 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal17 = 4
End Sub 

Sub rdoProdChildMthd5_OnClick()
	If lgRdoOldVal8 = 5 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal8 = 5
End Sub 

Sub rdoBacklogFlg2_OnClick()
	If lgRdoOldVal16 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal16 = 2
End Sub 


'Add 2006-07-18
Sub rdoProdRscMthd1_OnClick()
	If lgRdoOldVal18 = 1 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal18 = 1
End Sub 

Sub rdoProdRscMthd2_OnClick()
	If lgRdoOldVal18 = 2 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal18 = 2
End Sub 

Sub rdoProdRscMthd3_OnClick()
	If lgRdoOldVal18 = 3 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal18 = 3
End Sub 

Sub rdoProdRscMthd4_OnClick()
	If lgRdoOldVal18 = 4 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal18 = 4
End Sub 

Sub rdoProdRscMthd5_OnClick()
	If lgRdoOldVal8 = 5 Then Exit Sub
	lgBlnFlgChgValue = True
	lgRdoOldVal8 = 5
End Sub 




'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False
    Err.Clear                 '��: Protect system from crashing

    '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")     '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
      
	If frm1.txtPlantCd1.value = "" Then
		frm1.txtPlantNm1.value = ""
	End If
    
    Call ggoOper.ClearField(Document, "2")          '��: Clear Contents  Field
    'Call SetDefaultVal
    Call InitVariables               '��: Initializes local global variables

	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then         '��: This function check indispensable field
       Exit Function
    End If
    
	'-----------------------
    'Query function call area
    '----------------------- 
    If DbQuery = False Then   
		Exit Function           
    End If                   '��: Query db data
       
    FncQuery = True                '��: Processing is OK
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
End Function

'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
 Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
	'-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '��: No data changed!!
        Exit Function
    End If
    
    if lgIntFlgMode = parent.OPMD_CMODE then 
        IntRetCD = DisplayMsgBox("900002", "X", "X", "X")  
        Exit Function
    End If
    
	'-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2") Then                             '��: Check contents area
       Exit Function
    End If
    
	'-----------------------
    'Save function call area
    '-----------------------
    If lgPlantCd <> frm1.txtPlantCd1.value then
		Call DisplayMsgBox("900002", "X", "X", "X")          
	Else
		If DbSave = False Then   
			Exit Function           
		End If 
	End If
    
    FncSave = True                                                          '��: Processing is OK
    
End Function

'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================
Function FncCancel() 
End Function

'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================
Function FncInsertRow() 
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================
Function FncDeleteRow() 
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
    Dim strVal
    Dim IntRetCD 
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
 
	'-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")     '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '------------------------------------
    'Data Sheet �ʱ�ȭ 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")          '��: Clear Contents  Field
    
    'Call SetDefaultVal
    Call InitVariables               '��: Initializes local global variables
 
    Err.Clear                                                               '��: Protect system from crashing
    
    LayerShowHide(1)
  
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001      '��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd1=" & Trim(UCase(frm1.txtPlantCd1.value))  '��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&PrevNextFlg=" & "P"         '��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)          '��: �����Ͻ� ASP �� ���� 

End Function

'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================
 Function FncNext() 
    Dim strVal
    Dim IntRetCD 

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002","X","X","X")                                '��: �ؿ� �޼����� ID�� ó���ؾ� �� 
        Exit Function
    End If
    
    '-----------------------
    'Check previous data area
    '----------------------- 
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")     '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    '------------------------------------
    'Data Sheet �ʱ�ȭ 
    '------------------------------------
    Call ggoOper.ClearField(Document, "2")          '��: Clear Contents  Field
    
    'Call SetDefaultVal
    Call InitVariables               '��: Initializes local global variables


    Err.Clear                                                               '��: Protect system from crashing
    
    LayerShowHide(1)
  
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001      '��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd1=" & Trim(UCase(frm1.txtPlantCd1.value))  '��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&PrevNextFlg=" & "N"         '��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)          '��: �����Ͻ� ASP �� ���� 

End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)           '��: ȭ�� ���� 
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                   '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")    '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
Function DbDelete() 
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()              '��: ���� ������ ���� ���� 
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
 Function DbQuery() 
    
    DbQuery = False                                                         '��: Processing is NG
    
    Dim strVal
    
    LayerShowHide(1)
       
    strVal = BIZ_PGM_QRY_ID & "?txtMode=" & parent.UID_M0001       '��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtPlantCd1=" & Trim(UCase(frm1.txtPlantCd1.value))  '��: ��ȸ ���� ����Ÿ 
    strVal = strVal & "&PrevNextFlg=" & ""
    
	Call RunMyBizASP(MyBizASP, strVal)          '��: �����Ͻ� ASP �� ���� 
 
    DbQuery = True                                                          '��: Processing is NG

End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()              '��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE            '��: Indicates that current mode is Update mode
    lgBlnFlgChgValue = false
    
    lgPlantCd = frm1.txtPlantCd1.value 
    Call ggoOper.LockField(Document, "Q")         '��: This function lock the suitable field
    
	Call SetToolbar("11001000110111")
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave() 

	DbSave = False               '��: Processing is NG

	LayerShowHide(1)
	  
	With frm1
		.txtMode.value = parent.UID_M0002           '��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
		.txtUpdtUserId.value = parent.gUsrID
		Call ExecMyBizASP(frm1, BIZ_PGM_SAVE_ID)          
	End With
	 
	DbSave = True                                                           '��: Processing is NG
    
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()               '��: ���� ������ ���� ���� 

    Call InitVariables
    
    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
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
								<td NOWRAP background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ȯ�漳��</font></td>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>�� ��</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtPlantCd1" SIZE=10 MAXLENGTH=4 tag="12XXXU" ALT="����" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT ID="txtPlantNm1" NAME="txtPlantNm1" SIZE=30 tag="14X" ALT="�����"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<!--<TABLE <%=LR_SPACE_TYPE_60%>>-->
						<TABLE CLASS="TB2" CELLSPACING=0>
							<TR>
								<TD WIDTH=50% valign=top>
									<FIELDSET>
										<LEGEND>�������</LEGEND>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>������ذ��� ����</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoENGBOMFLG ID=rdoENGBOMFLG1 tag="21" VALUE="Y" ><LABEL FOR=rdoENGBOMFLG1>��</LABEL>&nbsp;&nbsp;&nbsp;
																	 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoENGBOMFLG ID=rdoENGBOMFLG2 tag="21" VALUE="N" CHECKED ><LABEL FOR=rdoENGBOMFLG2>�ƴϿ�</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>BOM �̷°��� ����</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoBOMHISTORYFLG ID=rdoBOMHISTORYFLG1 tag="21" VALUE="Y" ><LABEL FOR=rdoBOMHISTORYFLG1>��</LABEL>&nbsp;&nbsp;&nbsp;
																	 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoBOMHISTORYFLG ID=rdoBOMHISTORYFLG2 tag="21" VALUE="N" CHECKED ><LABEL FOR=rdoBOMHISTORYFLG2>�ƴϿ�</LABEL></TD>
											</TR>																															
											<TR>
												<TD CLASS=TD5 NOWRAP>����� ����L/T ���뿩��</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoRoutingLTFlg ID=rdoRoutingLTFlg1 tag="21" VALUE="Y" ><LABEL FOR=rdoRoutingLTFlg1>��</LABEL>&nbsp;&nbsp;&nbsp;
																	 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoRoutingLTFlg ID=rdoRoutingLTFlg2 tag="21" VALUE="N" CHECKED ><LABEL FOR=rdoRoutingLTFlg2>�ƴϿ�</LABEL></TD>
											</TR>
											<TR>
												<TD CLASS=TD5 NOWRAP>������ ���� ���뿩��</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoOprCostFlg ID=rdoOprCostFlg1 tag="21" VALUE="Y" ><LABEL FOR=rdoOprCostFlg1>��</LABEL>&nbsp;&nbsp;&nbsp;
																	 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoOprCostFlg ID=rdoOprCostFlg2 tag="21" VALUE="N" CHECKED ><LABEL FOR=rdoOprCostFlg2>�ƴϿ�</LABEL></TD>
											</TR>																															
										</TABLE>
									</FIELDSET>

									<FIELDSET>
										<LEGEND>MPS/MRP</LEGEND>
										<TABLE CLASS="BasicTB" CELLSPACING=0>
											<TR>
												<TD CLASS=TD5 NOWRAP>MPS �ۼ��� �����</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdEtcMthd ID=rdoProdEtcMthd1 tag="21" VALUE="Y" ><LABEL FOR=rdoProdEtcMthd1>����</LABEL>&nbsp;
																	 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdEtcMthd ID=rdoProdEtcMthd2 tag="21" VALUE="N" CHECKED ><LABEL FOR=rdoProdEtcMthd2>����</LABEL></TD>
											</TR>	
											<TR>
												<TD CLASS=TD5 NOWRAP>Mixedǰ MPS Consume���</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoMPSMETHOD ID=rdoMPSMETHOD1 tag="21" VALUE="1" CHECKED><LABEL FOR=rdoMPSMETHOD1>SP vs. SO MAX</LABEL>&nbsp;
																	 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoMPSMETHOD ID=rdoMPSMETHOD2 tag="21" VALUE="2" ><LABEL FOR=rdoMPSMETHOD2>Sales Plan �߽� Consume</LABEL>&nbsp;
																	 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoMPSMETHOD ID=rdoMPSMETHOD3 tag="21" VALUE="3" ><LABEL FOR=rdoMPSMETHOD3>Sales Order �߽� Consume</LABEL></TD>
											</TR>	
											<TR>
												<TD CLASS=TD5 NOWRAP>MRP ��ȯ�� ��ü �ڵ�����</TD>
												<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdFlg ID=rdoProdFlg1 tag="21" VALUE="Y" ><LABEL FOR=rdoProdFlg1>��</LABEL>&nbsp;&nbsp;&nbsp;
																	 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdFlg ID=rdoProdFlg2 tag="21" VALUE="N" CHECKED ><LABEL FOR=rdoProdFlg2>�ƴϿ�</LABEL></TD>
											</TR>
										</TABLE>
									</FIELDSET>

									<FIELDSET>
										<LEGEND>��������</LEGEND>
										<FIELDSET>
											<LEGEND>��������ȹ����</LEGEND>
											<TABLE CLASS="BasicTB" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>���ְ��� ��ǰ�� ����</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdMthd ID=rdoProdMthd1 tag="21" VALUE="1" ><LABEL FOR=rdoProdMthd1>����</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdMthd ID=rdoProdMthd2 tag="21" VALUE="2" ><LABEL FOR=rdoProdMthd2>�ʰ���</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdMthd ID=rdoProdMthd3 tag="21" VALUE="3" CHECKED ><LABEL FOR=rdoProdMthd3>������</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdMthd ID=rdoProdMthd4 tag="21" VALUE="4" ><LABEL FOR=rdoProdMthd4>�İ���</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdMthd ID=rdoProdMthd5 tag="21" VALUE="5" ><LABEL FOR=rdoProdMthd5>������</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>���������� ��ǰ�� ����</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdChildMthd ID=rdoProdChildMthd1 tag="21" VALUE="1" CHECKED><LABEL FOR=rdoProdChildMthd1>����</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdChildMthd ID=rdoProdChildMthd2 tag="21" VALUE="2" ><LABEL FOR=rdoProdChildMthd2>�ʰ���</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdChildMthd ID=rdoProdChildMthd3 tag="21" VALUE="3" ><LABEL FOR=rdoProdChildMthd3>������</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdChildMthd ID=rdoProdChildMthd4 tag="21" VALUE="4" ><LABEL FOR=rdoProdChildMthd4>�İ���</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdChildMthd ID=rdoProdChildMthd5 tag="21" VALUE="5" ><LABEL FOR=rdoProdChildMthd5>������</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>���������� �ڿ� ����</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdRscMthd ID=rdoProdRscMthd1 tag="21" VALUE="1" CHECKED><LABEL FOR=rdoProdRscMthd1>����</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdRscMthd ID=rdoProdRscMthd2 tag="21" VALUE="2" ><LABEL FOR=rdoProdRscMthd2>�ʰ���</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdRscMthd ID=rdoProdRscMthd3 tag="21" VALUE="3" ><LABEL FOR=rdoProdRscMthd3>������</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdRscMthd ID=rdoProdRscMthd4 tag="21" VALUE="4" ><LABEL FOR=rdoProdRscMthd4>�İ���</LABEL>&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoProdRscMthd ID=rdoProdRscMthd5 tag="21" VALUE="5" ><LABEL FOR=rdoProdRscMthd5>������</LABEL></TD>
												</TR>
											</TABLE>		
										</FIELDSET>
										<FIELDSET>
											<LEGEND>��������Ȯ��</LEGEND>
											<TABLE CLASS="BasicTB" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>�������� Ȯ���� ��� Check</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoRlsInvChkFlg ID=rdoRlsInvChkFlg1 tag="21" VALUE="1" CHECKED><LABEL FOR=rdoRlsInvChkFlg1>���Check ����</LABEL>
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoRlsInvChkFlg ID=rdoRlsInvChkFlg2 tag="21" VALUE="2"><LABEL FOR=rdoRlsInvChkFlg2>�����Check</LABEL>
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoRlsInvChkFlg ID=rdoRlsInvChkFlg3 tag="21" VALUE="3"><LABEL FOR=rdoRlsInvChkFlg3>�������Check</LABEL></TD>
												</TR>		
												<TR>
													<TD CLASS=TD5 NOWRAP>�������� ��뿩��</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoDELIVERYORDERFLG ID=rdoDELIVERYORDERFLG1 tag="21" VALUE="Y" ><LABEL FOR=rdoDELIVERYORDERFLG1>��</LABEL>&nbsp;&nbsp;&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoDELIVERYORDERFLG ID=rdoDELIVERYORDERFLG2 tag="21" VALUE="N" CHECKED ><LABEL FOR=rdoDELIVERYORDERFLG2>�ƴϿ�</LABEL></TD>
												</TR>	
											</TABLE>		
										</FIELDSET>
										<FIELDSET>
											<LEGEND>�������</LEGEND>
											<TABLE CLASS="BasicTB" CELLSPACING=0>
												<TR>
													<TD CLASS=TD5 NOWRAP>�����Է½� ���������� Check</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoPreOprChkFlg ID=rdoPreOprChkFlg1 tag="21" VALUE="Y" CHECKED><LABEL FOR=rdoPreOprChkFlg1>��</LABEL>&nbsp;&nbsp;&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoPreOprChkFlg ID=rdoPreOprChkFlg2 tag="21"VALUE="N"><LABEL FOR=rdoPreOprChkFlg2>�ƴϿ�</LABEL></TD>
												</TR>	
												<TR>
													<TD CLASS=TD5 NOWRAP>������ ��뿩��</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoExssRcptFlg ID=rdoExssRcptFlg1 tag="21" VALUE="Y" ><LABEL FOR=rdoExssRcptFlg1>��</LABEL>&nbsp;&nbsp;&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoExssRcptFlg ID=rdoExssRcptFlg2 tag="21" VALUE="N" CHECKED ><LABEL FOR=rdoExssRcptFlg2>�ƴϿ�</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>�������� �ڵ��԰�</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAutoRcptFlg ID=rdoAutoRcptFlg1 tag="21" VALUE="Y" CHECKED><LABEL FOR=rdoAutoRcptFlg1>��</LABEL>&nbsp;&nbsp;&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAutoRcptFlg ID=rdoAutoRcptFlg2 tag="21" VALUE="N" ><LABEL FOR=rdoAutoRcptFlg2>�ƴϿ�</LABEL></TD>
												</TR>	
												<TR>
													<TD CLASS=TD5 NOWRAP>�԰�� �ڵ�����</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoOrdCloseMthd ID=rdoOrdCloseMthd1 tag="21" VALUE="Y" ><LABEL FOR=rdoOrdCloseMthd1>��</LABEL>&nbsp;&nbsp;&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoOrdCloseMthd ID=rdoOrdCloseMthd2 tag="21" VALUE="N" CHECKED ><LABEL FOR=rdoOrdCloseMthd2>�ƴϿ�</LABEL></TD>
												</TR>
												<TR>
													<TD CLASS=TD5 NOWRAP>Backlog ��뿩��</TD>
													<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoBacklogFlg ID=rdoBacklogFlg1 tag="21" VALUE="Y" ><LABEL FOR=rdoBacklogFlg1>��</LABEL>&nbsp;&nbsp;&nbsp;
																		 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoBacklogFlg ID=rdoBacklogFlg2 tag="21" VALUE="N" CHECKED ><LABEL FOR=rdoBacklogFlg2>�ƴϿ�</LABEL></TD>
												</TR>	
											</TABLE>
										</FIELDSET>
									</FIELDSET>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtProdEtcFlg" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
