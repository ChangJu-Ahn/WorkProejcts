<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111qa1
'*  4. Program Name         : ����������ȸ 
'*  5. Program Desc         : ����������ȸ 
'*  6. Component List       : 
'*  7. Modified date(First) : 2001/01/08
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Min, Hak-jun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="vbscript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit								
'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                    
Dim IscookieSplit 
Dim lgSaveRow                      

Dim StartDate
Dim EndDate

EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  

Const BIZ_PGM_ID 		= "m3111qb1_KO441.asp"         
Const BIZ_PGM_JUMP_ID 	= "m3111qa2"             
Const C_MaxKey          = 28					

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()
    lgStrPrevKey     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
	lgIntFlgMode	= Parent.OPMD_CMODE
    lgPageNo         = ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtPoFrDt.Text	= StartDate
	frm1.txtPoToDt.Text	= EndDate
	' Tracker No.9743 �����ڵ� ���� - 2005.07.22 =========================================
	frm1.txtPlantCd.value=parent.gPlant
	frm1.txtPlantNm.value=parent.gPlantNm
	' Tracker No.9743 �����ڵ� ���� - 2005.07.22 =========================================
	frm1.txtPlantCd.focus
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
  	frm1.txtPlantCd.value = lgPLCd
	End If
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd, "Q") 
  	frm1.txtPurGrpCd.value = lgPGCd
	End If
	Set gActiveElement = document.activeElement
End Sub

'======================================================================================
' Function Name : InitComboBox()
' Function Desc : Initialize ComboBox
'========================================================================================
Sub InitComboBox()
	Call SetCombo(frm1.cboPrcFlg, "T", "���ܰ�")
	Call SetCombo(frm1.cboPrcFlg, "F", "���ܰ�")
End Sub

'======================================================================================
' Function Name : ConvertComboValue()
' Function Desc : code value -> code name display 
'========================================================================================
Sub ConvertComboValue()

	Dim intCnt
	Dim strFlg, strPriceVal
	
	if frm1.vspdData.MaxRows < 1 then exit sub
	
	strPriceVal = GetKeyPos("A",14) 
	For intCnt = 1 to frm1.vspdData.MaxRows
		frm1.vspddata.Row = intCnt
		frm1.vspddata.col = strPriceVal '�ܰ� flg 2003-04-10 ������ 
		strFlg			  = Trim(frm1.vspddata.text)			
		If strFlg = "T" or strFlg = "���ܰ�" then			
			frm1.vspddata.text = "���ܰ�"
		Else 
			frm1.vspddata.text = "���ܰ�"
		End if
	Next
End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA")%>
End Sub

'======================= 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("M3111QA1","G","A","V20051201", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock("A") 
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(ByVal pOpt)
    If pOpt = "A" Then
      ggoSpread.Source = frm1.vspdData
      ggoSpread.SpreadLockWithOddEvenRowColor()
	Else
	End If
End Sub

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)
	
	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ���				
    			
    iCalledAspName = AskPRAspName("B1B11PA3")
    
    If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
	End If	
End Function

'------------------------------------------  OpenPlantCd()  -------------------------------------------------
'	Name : OpenPlantCd()
'	Description : Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	If frm1.txtPlantCd.className = "protected" Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����"						<%' �˾� ��Ī %>
	arrParam(1) = "B_PLANT"      					<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		<%' Code Condition%>
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)		<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "����"						<%' TextBox ��Ī %>
	
    arrField(0) = "PLANT_CD"						<%' Field��(0)%>
    arrField(1) = "PLANT_NM"						<%' Field��(1)%>
    
    arrHeader(0) = "����"						<%' Header��(0)%>
    arrHeader(1) = "�����"						<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd.value=""
	frm1.txtItemNm.value=""
End Function


'------------------------------------------  OpenBpCd()  -------------------------------------------------
'	Name : OpenBpCd()
'	Description : Supplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "����ó"						<%' �˾� ��Ī %>
	arrParam(1) = "B_Biz_Partner"					<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtBpCd.Value)			<%' Code Condition%>
	'arrParam(3) = Trim(frm1.txtBpNm.Value)			<%' Name Cindition%>
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					<%' Where Condition%>
	arrParam(5) = "����ó"						<%' TextBox ��Ī %>
	
    arrField(0) = "BP_CD"							<%' Field��(0)%>
    arrField(1) = "BP_NM"							<%' Field��(1)%>
    
    arrHeader(0) = "����ó"						<%' Header��(0)%>
    arrHeader(1) = "����ó��"					<%' Header��(1)%>
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd.focus
		Exit Function
	Else
		frm1.txtBpCd.Value = arrRet(0)
		frm1.txtBpNm.Value = arrRet(1)
		frm1.txtBpCd.focus
	End If	
End Function

 '------------------------------------------  OpenPoType()  -------------------------------------------------
'	Name : OpenPoType()
'	Description : PoType PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoType()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "��������"						<%' �˾� ��Ī %>
	arrParam(1) = "M_CONFIG_PROCESS"					<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtPoType.Value)			<%' Code Condition%>
'	arrParam(3) = Trim(frm1.txtPoTypeNm.Value)			<%' Name Condition%>
	arrParam(4) = ""									<%' Where Condition%>
	arrParam(5) = "��������"						<%' TextBox ��Ī %>
	
    arrField(0) = "PO_TYPE_CD"							<%' Field��(0)%>
    arrField(1) = "PO_TYPE_NM"							<%' Field��(1)%>
        
    arrHeader(0) = "��������"						<%' Header��(0)%>
    arrHeader(1) = "�������¸�"						<%' Header��(1)%>
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPoType.focus
		Exit Function
	Else
		frm1.txtPoType.Value = arrRet(0)
		frm1.txtPoTypeNm.Value = arrRet(1)
		frm1.txtPoType.focus
	End If	
End Function

'------------------------------------------  OpenPurGrpCd()  -------------------------------------------------
'	Name : OpenPurGrpCd()
'	Description : PurGrp PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPurGrpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
	If frm1.txtPurGrpCd.className = "protected" Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "���ű׷�"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "���ű׷�"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "���ű׷�"		
    arrHeader(1) = "���ű׷��"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd.focus
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)
		frm1.txtPurGrpCd.focus
	End If	

End Function 

'------------------------------------------  OpenTrackNo()  -------------------------------------------------
'	Name : OpenTrackNo()
'	Description : TrackNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackNo()
	Dim arrRet
	Dim arrParam(5)
	Dim IntRetCD
	Dim iCalledAspName
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True

	arrParam(0) = ""	'�ֹ�ó 
	arrParam(1) = ""	'�����׷� 
    arrParam(2) = Trim(frm1.txtPlantCd.value)	'���� 
    arrParam(3) = ""	'��ǰ�� 
    arrParam(4) = ""	'���ֹ�ȣ 
    arrParam(5) = ""	'�߰� Where�� 
    
 	iCalledAspName = AskPRAspName("S3135PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "S3135PA1", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
   
	lgIsOpenPop = False

	If arrRet = "" Then
		frm1.txtTrackNo.focus
		Exit Function
	Else
		frm1.txtTrackNo.Value = Trim(arrRet)
		frm1.txtTrackNo.focus
	End If	
End Function

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
'	Name : PopZAdoConfigGrid()
'	Description : Group Condition PopUp
'---------------------------------------------------------------------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupPopup("A")
End Sub


'------------------------------------  OpenGroupPopup()  ----------------------------------------------
'	Name : OpenGroupPopup()
'	Description : Group Condition PopUp
'--------------------------------------------------------------------------------------------------------
Function OpenGroupPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOGroupPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.GROUPW_WIDTH & "px; dialogHeight=" & parent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function


'==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== 
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						<% 'Cookie Split String : CookiePage Function Use%>

	If Kubun = 1 Then								<% 'Jump�� ȭ���� �̵��� ��� %>
	
		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)				
		WriteCookie CookieSplit , IsCookieSplit		<% 'Jump�� ȭ���� �̵��Ҷ� �ʿ��� Cookie �������� %>
	
		If Len(Trim(frm1.txtPlantCd.value)) Then	<% 'Jump�� ȭ���� �̵��Ҷ� �ʿ��� ���� Cookie �������� %>
			WriteCookie "PlantCd",Trim(frm1.txtPlantCd.value) 
		Else
			WriteCookie "PlantCd",""
		End If
		
		If Len(Trim(frm1.txtPurGrpCd.value)) Then
			WriteCookie "PurGrpCd",Trim(frm1.txtPurGrpCd.value) 
		Else
			WriteCookie "PurGrpCd",""
		End If
		
		If Len(Trim(frm1.txtBpCd.value)) Then
			WriteCookie "BpCd",Trim(frm1.txtBpCd.value) 
		Else
			WriteCookie "BpCd",""
		End If
		
		If Len(Trim(frm1.txtPoFrDt.text)) Then
			WriteCookie "PoFrDt",Trim(frm1.txtPoFrDt.text) 
		Else
			WriteCookie "PoFrDt",""
		End If
		
		If Len(Trim(frm1.txtPoToDt.text)) Then
			WriteCookie "PoToDt",Trim(frm1.txtPoToDt.text) 
		Else
			WriteCookie "PoToDt",""
		End If
		
		If Len(Trim(frm1.txtItemCd.value)) Then
			WriteCookie "ItemCd",Trim(frm1.txtItemCd.value) 
		Else
			WriteCookie "ItemCd",""
		End If
		
		If Len(Trim(frm1.txtTrackNo.value)) Then
			WriteCookie "TrackNo",Trim(frm1.txtTrackNo.value) 
		Else
			WriteCookie "TrackNo",""
		End If
		
		If Len(Trim(frm1.txtPoType.value)) Then
			WriteCookie "PoType",Trim(frm1.txtPoType.value) 
		Else
			WriteCookie "PoType",""
		End If
		
		If Len(Trim(frm1.cboPrcFlg.value)) Then
			WriteCookie "PrcFlg",Trim(frm1.cboPrcFlg.value) 
		Else
			WriteCookie "PrcFlg",""
		End If
		
				
		Call PgmJump(BIZ_PGM_JUMP_ID)		

	ElseIf Kubun = 0 Then			

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, Parent.gRowSep)

		Dim iniSep

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF

End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call InitVariables														'��: Initializes local global variables
	Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")										'��: ��ư ���� ����	
    Call InitComboBox()
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
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Function FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Function

'==========================================================================================
'   Event Name : txtPoFrDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtPoFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoFrDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : txtPoToDt
'   Event Desc : Date OCX Double Click
'==========================================================================================
Sub txtPoToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPoToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoToDt.Focus
	End If
End Sub
'==========================================================================================
'   Event Name : OCX_KeyDown()
'   Event Desc : 
'==========================================================================================
Sub txtPoFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()		
End Sub

Sub txtPoToDt_Keydown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
	
'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Const C_L_BpCd			= 1
	Const C_L_BpNm			= 2
	Const C_L_PlantCd		= 3
	Const C_L_PlantNm		= 4
	Const C_L_PurGrp		= 5
	Const C_L_PurGrpNm		= 6
	Const C_L_PoDt			= 7
	Const C_L_ItemCd		= 8
	Const C_L_ItemNm		= 9
	Const C_L_Spec			= 10
	Const C_L_TrackingNo	= 11
	Const C_L_PoTypeCd		= 12
	Const C_L_PoTypeNm		= 13
	Const C_L_PoPrcFlg		= 14
	Const C_L_PoUnit		= 15
	Const C_L_PoCur			= 16

	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	Call SetPopupMenuItemInf("00000000001")
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
	
	IsCookieSplit=""
    With frm1.vspddata
		.Row = Row
		
		.Col = GetKeyPos("A", C_L_BpCd)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_BpNm)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_PlantCd)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_PlantNm)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_PurGrp)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_PurGrpNm)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_PoDt)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_ItemCd)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_ItemNm)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_TrackingNo)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_PoTypeCd)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_PoTypeNm)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep

		.Col = GetKeyPos("A", C_L_PoPrcFlg)
		IsCookieSplit =  IsCookieSplit & Trim(.Text) & parent.gRowSep
	End With	
End Sub

'======================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : Ư�� column�� click�Ҷ� 
'======================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    
	
'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================

Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '��: ������ üũ	
		If lgPageNo <> "" Then		                                                    '���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
 			If CheckRunningBizProcess = True Then
				Exit Sub
			End If			
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub

'*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************
Function FncQuery() 
    FncQuery = False                                                        '��: Processing is NG
   
    Err.Clear                                                               '��: Protect system from crashing
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

	with frm1
		if (UniConvDateToYYYYMMDD(.txtPoFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtPoToDt.text,Parent.gDateFormat,"")) And Trim(.txtPoFrDt.text) <> "" And Trim(.txtPoToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","������", "X")			
			Exit Function
		End if   
	End with
	
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function

	Set gActiveElement = document.activeElement
    FncQuery = True															'��: Processing is OK

End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)                                     <%'��:ȭ�� ����, Tab ���� %>
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
	Set gActiveElement = document.activeElement
End Function

'*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'*********************************************************************************************************

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
	Dim strVal
	Dim strClsFlg
	
    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	If LayerShowHide(1) = False Then Exit Function

    With frm1
		If .rdoClsFlg(0).checked Then
			strClsFlg = ""
		ElseIf .rdoClsFlg(1).checked Then
			strClsFlg = "Y"
		Else
			strClsFlg = "N"
		End If

		If lgIntFlgMode = Parent.OPMD_UMODE Then	
			strVal = BIZ_PGM_ID	& "?txtPlantCd=" & Trim(.hdnPlantCd.value)
			strVal = strVal	& "&txtPurGrpCd=" &	Trim(.hdnPurGrpCd.value)
			strVal = strVal	& "&txtBpCd="     &	Trim(.hdnBpCd.value)
			strVal = strVal	& "&txtPoFrDt="	  & Trim(.hdnPoFrDt.value)
			strVal = strVal	& "&txtPoToDt="	  & Trim(.hdnPoToDt.value)
			strVal = strVal	& "&txtItemCd="	  & Trim(.hdnItemCd.value)		
			strVal = strVal	& "&txtTrackNo=" & Trim(.hdnTrackNo.value)
			strVal = strVal	& "&txtPoType="	  & Trim(.hdnPoType.value)
			strVal = strVal	& "&txtPrcFlg="	  & Trim(.hdncboPrcFlg.value)
			strVal = strVal	& "&rdoClsFlg="	  & Trim(strClsFlg)
			strVal = strVal & "&lgPageNo="   & lgPageNo         
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		Else
			strVal = BIZ_PGM_ID	& "?txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal	& "&txtPurGrpCd=" &	Trim(.txtPurGrpCd.value)
			strVal = strVal	& "&txtBpCd="     &	Trim(.txtBpCd.value)
			strVal = strVal	& "&txtPoFrDt="	  & Trim(.txtPoFrDt.Text)
			strVal = strVal	& "&txtPoToDt="	  & Trim(.txtPoToDt.Text)
			strVal = strVal	& "&txtItemCd="	  & Trim(.txtItemCd.value)		
			strVal = strVal	& "&txtTrackNo=" & Trim(.txtTrackNo.value)
			strVal = strVal	& "&txtPoType="	  & Trim(.txtPoType.value)
			strVal = strVal	& "&txtPrcFlg="	  & Trim(.cboPrcFlg.value)
			strVal = strVal	& "&rdoClsFlg="	  & Trim(strClsFlg)
			strVal = strVal & "&lgPageNo="   & lgPageNo         
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		End If

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  

        Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With
    
    DbQuery = True
    Call SetToolbar("1100000000011111")										'��: ��ư ���� ����	

End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 
    '-----------------------
    'Reset variables area
    '-----------------------
	lgBlnFlgChgValue = False
    lgSaveRow        = 1
	lgIntFlgMode = Parent.OPMD_UMODE
	Call ConvertComboValue() 
    If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.focus
	Else
		frm1.txtPlantCd.focus
	End If
	Set gActiveElement = document.activeElement
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'	���: Tag�κ� ���� 
	' �Է� �ʵ��� ��� MaxLength=? �� ��� 
	' CLASS="required" required  : �ش� Element�� Style �� Default Attribute 
		' Normal Field�϶��� ������� ���� 
		' Required Field�϶��� required�� �߰��Ͻʽÿ�.
		' Protected Field�϶��� protected�� �߰��Ͻʽÿ�.
			' Protected Field�ϰ�� ReadOnly �� TabIndex=-1 �� ǥ���� 
	' Select Type�� ��쿡�� className�� ralargeCB�� ���� width="153", rqmiddleCB�� ���� width="90"
	' Text-Transform : uppercase  : ǥ�Ⱑ �빮�ڷ� �� �ؽ�Ʈ 
	' ���� �ʵ��� ��� 3���� Attribute ( DDecPoint DPointer DDataFormat ) �� ��� 
'######################################################################################################### -->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right>&nbsp;</td>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����"  NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd() ">
														   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="���ű׷�" NAME="txtPurGrpCd" SIZE=10 MAXLENGTH=4  tag="11xXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrpCd()">
														   <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 tag="14"></TD>					   
								</TR>					   
								<TR>						   
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����ó" NAME="txtBpCd" SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBpCd()">
														   <INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 tag="14"></TD>			
									<TD CLASS="TD5" NOWRAP>������</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtPoFrDt CLASSID=<%=gCLSIDFPDT%> tag="11X1" ALT="������"></OBJECT>');</SCRIPT>
												</td>
												<td>~</td>
												<td>
													<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtPoToDt CLASSID=<%=gCLSIDFPDT%> ALT="������" tag="11X1"></OBJECT>');</SCRIPT>
												</td>
											<tr>
										</table>
							         </TD>
	                            </TR>	
	                            <TR>
									<TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="ǰ��" NAME="txtItemCd" SIZE=34 MAXLENGTH=18  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd()"></TD>
									<TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="Tracking No." NAME="txtTrackNo" SIZE=34 MAXLENGTH=25  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackNo()"></TD>
								</TR>
	                            <TR>
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="ǰ��" NAME="txtItemNm" SIZE=34 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="��������"  NAME="txtPoType" SIZE=10 LANG="ko" MAXLENGTH=5 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPoType() ">
														   <INPUT TYPE=TEXT NAME="txtPoTypeNm" SIZE=20 tag="14"></TD>
								</TR>
	                            <TR>				   		   
									<TD CLASS="TD5" NOWRAP>�ܰ�����</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboPrcFlg" tag="11"  CLASS=cboNormal><OPTION VALUE="" selected></OPTION></SELECT></TD>
									<TD CLASS="TD5" NOWRAP>��������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="��ü"   NAME="rdoClsFlg" ID="rdoClsFlg0" CLASS="RADIO" value = "A" tag="11" checked><label for="rdoClsFlg0">&nbsp;��ü&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="����"   NAME="rdoClsFlg" ID="rdoClsFlg1" CLASS="RADIO" value = "Y" tag="11"><label for="rdoClsFlg1">&nbsp;����&nbsp;</label>
														   <INPUT TYPE=radio AlT="�̸���" NAME="rdoClsFlg" ID="rdoClsFlg2" CLASS="RADIO" value = "N" tag="11"><label for="rdoClsFlg2">&nbsp;�̸���&nbsp;&nbsp;</label></TD>
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
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>			 
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:CookiePage(1)">���ֻ���ȸ</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPurGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnTrackNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdncboPrcFlg" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
