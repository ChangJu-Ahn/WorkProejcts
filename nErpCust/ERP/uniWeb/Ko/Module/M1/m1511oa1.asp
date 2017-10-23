<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1511OA1
'*  4. Program Name         : 구매품기준정보출력 
'*  5. Program Desc         : 구매품기준정보출력 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/06/25
'*  8. Modified date(Last)  : 2003/06/25
'*  9. Modifier (First)     : Kang Su Hwan
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
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
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit						
<!-- #Include file="../../inc/lgvariables.inc" -->
<!--'==========================================  1.2.1 Global 상수 선언  ======================================
'==========================================================================================================!-->
Dim StartDate, EndDate
	StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
	StartDate = UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat) 
	EndDate   = UniConvDateAToB("<%=GetSvrDate%>"  ,Parent.gServerDateFormat,Parent.gDateFormat)
	
<!-- '==========================================  1.2.3 Global Variable값 정의  ===============================
'========================================================================================================= !-->
Dim lblnWinEvent
Dim IsOpenPop 

<!-- '==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= !-->
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

<!-- '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= !-->
Sub SetDefaultVal()
	frm1.txtPlantCd.value=parent.gPlant
	frm1.txtPlantNm.value=parent.gPlantNm
End Sub

<!--'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== !-->
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","QA") %> 
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub

<!-- '------------------------------------------  OpenPlantCd()  -------------------------------------------------
'	Name : OpenPlantCd()
'	Description : OpenPlantCd PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"						
	arrParam(1) = "B_PLANT"      					
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)		
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
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd1.value=""
	frm1.txtItemNm1.value=""
	frm1.txtItemCd2.value=""
	frm1.txtItemNm2.value=""
End Function

<!-- '------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : OpenItemCd PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenItemCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtItemCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	IsOpenPop = True

	arrParam(0) = "품목"						
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"	
	arrParam(2) = Trim(frm1.txtItemCd1.Value)		
'	arrParam(3) = Trim(frm1.txtItemNm.Value)		
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
	if Trim(frm1.txtPlantCd.Value)<>"" then
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " "    
	End if 
	arrParam(5) = "품목"						

    arrField(0) = "B_Item.Item_Cd"					
    arrField(1) = "B_Item.Item_NM"					
    arrField(2) = "B_Plant.Plant_Cd"				
    arrField(3) = "B_Plant.Plant_NM"				
    
    arrHeader(0) = "품목"						
    arrHeader(1) = "품목명"						
    arrHeader(2) = "공장"						
    arrHeader(3) = "공장명"						
    
    iCalledAspName = AskPRAspName("m1111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m1111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam,arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd1.focus
		Exit Function
	Else
		frm1.txtItemCd1.Value = arrRet(0)
		frm1.txtItemNm1.Value = arrRet(1)
		frm1.txtItemCd1.focus
	End If	
End Function

<!-- '------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : OpenItemCd PopUp
'--------------------------------------------------------------------------------------------------------- !-->
Function OpenItemCd2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Or UCase(frm1.txtItemCd2.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "공장", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if
	
	IsOpenPop = True

	arrParam(0) = "품목"						
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"	
	arrParam(2) = Trim(frm1.txtItemCd2.Value)		
'	arrParam(3) = Trim(frm1.txtItemNm.Value)		
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
	if Trim(frm1.txtPlantCd.Value)<>"" then
		arrParam(4) = arrParam(4) & "And B_Plant.Plant_Cd= " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " "    
	End if 
	arrParam(5) = "품목"						

    arrField(0) = "B_Item.Item_Cd"					
    arrField(1) = "B_Item.Item_NM"					
    arrField(2) = "B_Plant.Plant_Cd"				
    arrField(3) = "B_Plant.Plant_NM"				
    
    arrHeader(0) = "품목"						
    arrHeader(1) = "품목명"						
    arrHeader(2) = "공장"						
    arrHeader(3) = "공장명"						
    
    iCalledAspName = AskPRAspName("m1111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m1111pa1", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam,arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd2.focus
		Exit Function
	Else
		frm1.txtItemCd2.Value = arrRet(0)
		frm1.txtItemNm2.Value = arrRet(1)
		frm1.txtItemCd2.focus
	End If	
End Function

 '------------------------------------------  OpenPurOrg()  -------------------------------------------------
'	Name : OpenPurOrg()	구매조직 
'	Description : PurOrg PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPurOrg()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPurOrg.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매조직팝업"	
	arrParam(1) = "B_PUR_ORG"				
	arrParam(2) = Trim(frm1.txtPurOrg.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "구매조직"
	
    arrField(0) = "PUR_ORG"	
    arrField(1) = "PUR_ORG_NM"	
    
    arrHeader(0) = "구매조직"		
    arrHeader(1) = "구매조직명"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetPurOrg(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtPurOrg.focus
	
End Function

'========================================================================================
Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6

	'-----------------------------------------------------------------------------------------------------
	' List Minor code for Item Account
	'-----------------------------------------------------------------------------------------------------
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1001' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboAccount, lgF0, lgF1, Chr(11))

	'-----------------------------------------------------------------------------------------------------
	' List Minor code for Procurement Type(자재Type)
	'-----------------------------------------------------------------------------------------------------
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'P1008' ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboMatType, Chr(11)&lgF0, Chr(11)&lgF1, Chr(11))
End Sub

<!-- '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= !-->
 Sub Form_Load()

	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 
    Call InitComboBox
    
    frm1.txtPlantCd.focus 
	Set gActiveElement = document.activeElement
    
End Sub

'==========================================  2.2.6 ChkKeyField()  =======================================
'	Name : ChkKeyField()
'	Description : 
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	
	ChkKeyField = true
	
	'Plant
	strWhere = " PLANT_CD = '" & FilterVar(frm1.txtPlantCd.value, "","SNM") & "' "
	Call CommonQueryRs(" PLANT_NM "," B_PLANT ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("17a003","X","공장","X")
		frm1.txtPlantCd.value = ""
		frm1.txtPlantNm.value = ""
		frm1.txtPlantCd.focus
		ChkKeyField = False
		Exit Function
	Else
		frm1.txtPlantNm.value = split(lgF0,chr(11))(0)
	End If
	
	'Item Account
	strWhere = " MAJOR_CD = 'P1001' AND MINOR_CD = '" & FilterVar(frm1.cboAccount.value, "","SNM") & "' "
	Call CommonQueryRs(" MINOR_NM "," B_MINOR ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("17a003","X","품목계정","X")
		frm1.cboAccount.focus
		ChkKeyField = False
		Exit Function
	End If

	'Pur.Org.
	If Trim(frm1.txtPurOrg.value)<> "" Then
		strWhere = " PUR_ORG = '" & FilterVar(frm1.txtPurOrg.value, "","SNM") & "' "
		Call CommonQueryRs(" PUR_ORG_NM "," B_PUR_ORG ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","구매조직","X")
			frm1.txtPurOrg.value = ""
			frm1.txtPurOrgNm.value = ""
			frm1.txtPurOrg.focus
			ChkKeyField = False
			Exit Function
		Else
			frm1.txtPurOrgNm.value = split(lgF0,chr(11))(0)
		End If
	End If
	
	'Material Type - P1008
	If Trim(frm1.cboMatType.value) <> "" Then
		strWhere = " MAJOR_CD = 'P1008' AND MINOR_CD = '" & FilterVar(frm1.cboMatType.value, "","SNM") & "' "
		Call CommonQueryRs(" MINOR_NM "," B_MINOR ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","자재Type","X")
			frm1.cboMatType.focus
			ChkKeyField = False
			Exit Function
		End If
	End IF
	
	'품목1
	If Trim(frm1.txtItemCd1.value) <> "" Then
		strWhere = " ITEM_CD = '" & FilterVar(frm1.txtItemCd1.value, "","SNM") & "' "
		Call CommonQueryRs(" ITEM_NM "," B_ITEM ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","자재Type","X")
			frm1.txtItemCd1.value = ""
			frm1.txtItemNm1.value = ""
			frm1.txtItemCd1.focus
			ChkKeyField = False
			Exit Function
		Else
			frm1.txtItemNm1.value = split(lgF0,chr(11))(0)
		End If
	End If

	'품목2
	If Trim(frm1.txtItemCd2.value) <> "" Then
		strWhere = " ITEM_CD = '" & FilterVar(frm1.txtItemCd2.value, "","SNM") & "' "
		Call CommonQueryRs(" ITEM_NM "," B_ITEM ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","자재Type","X")
			frm1.txtItemCd2.value = ""
			frm1.txtItemNm2.value = ""
			frm1.txtItemCd2.focus
			ChkKeyField = False
			Exit Function
		Else
			frm1.txtItemNm2.value = split(lgF0,chr(11))(0)
		End If
	End If
End Function

<!--
'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
-->
 Function FncPrint() 
	Call parent.FncPrint()
End Function
<!--
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
-->
 Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function
<!--'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================!-->
 Function FncBtnPrint() 
	Dim StrUrl
	dim var1,var2,var3,var4,var5,var6,var7,var8
    	
    If Not chkField(Document, "1") Then									
        Exit Function
    End If
	
	IF ChkKeyField() = False Then 
		Exit Function
    End if
    
	with frm1
		if (frm1.txtItemCd1.value <> "") AND  (frm1.txtItemCd2.value <> "") then
			if  UCase(frm1.txtItemCd1.value) > UCase(frm1.txtItemCd2.value)  then	
				Call DisplayMsgBox("17a003","X","품목","X")
				frm1.txtItemCd1.focus 
				Exit Function
			End if  
		End if 
	END WITH
	
	On Error Resume Next                                                    '☜: Protect system from crashing
	
	var1 = frm1.txtPlantCd.value
	var2 = frm1.cboAccount.value
	If Trim(frm1.txtPurOrg.value) <> "" Then
		var3 = Trim(frm1.txtPurOrg.value)
		var4 = Trim(frm1.txtPurOrg.value)
	Else
		var3 = ""
		var4 = "ZZZZ"
	End If
	If Trim(frm1.cboMatType.value) <> "" Then
		var5 = Trim(frm1.cboMatType.value)
		var6 = Trim(frm1.cboMatType.value)
	Else
		var5 = ""
		var6 = "ZZ"
	End If
	
	If Trim(frm1.txtItemCd1.value) <> "" Then
		var7 = Trim(frm1.txtItemCd1.value)
	Else
		var7 = "ZZZZZZZZZZZZZZZZZZ"
	End If

	If Trim(frm1.txtItemCd2.value) <> "" Then
		var8 = Trim(frm1.txtItemCd2.value)
	Else
		var8 = "ZZZZZZZZZZZZZZZZZZ"
	End If

	strUrl = strUrl & "plant|" & var1
	strUrl = strUrl & "|account|" & var2
	strUrl = strUrl & "|fr_purorg|" & var3
	strUrl = strUrl & "|to_purorg|" & var4
	strUrl = strUrl & "|fr_mattype|" & var5
	strUrl = strUrl & "|to_mattype|" & var6
	strUrl = strUrl & "|fr_item|" & var7
	strUrl = strUrl & "|to_item|" & var8
	
'----------------------------------------------------------------
' Print 함수에서 호출 
'----------------------------------------------------------------
	ObjName = AskEBDocumentName("m1511oa1","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)
'----------------------------------------------------------------
End Function

<!--
'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
-->
Function BtnPreview() 
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
    
    IF ChkKeyField() = False Then 
		frm1.txtPlantCd.focus
		Exit Function
    End if
    
   With frm1
		If (frm1.txtItemCd1.value <> "") AND  (frm1.txtItemCd2.value <> "") Then
			If  UCase(frm1.txtItemCd1.value) > UCase(frm1.txtItemCd2.value)  Then	
				Call DisplayMsgBox("17a003","X","품목","X")
				frm1.txtItemCd1.focus 
				Exit Function
			End If  
		End If 
	END WITH
	
	dim var1,var2,var3,var4,var5,var6,var7,var8
	dim strUrl
	dim arrParam, arrField, arrHeader
		
	var1 = frm1.txtPlantCd.value
	var2 = frm1.cboAccount.value
	If Trim(frm1.txtPurOrg.value) <> "" Then
		var3 = Trim(frm1.txtPurOrg.value)
		var4 = Trim(frm1.txtPurOrg.value)
	Else
		var3 = ""
		var4 = "ZZZZ"
	End If
	If Trim(frm1.cboMatType.value) <> "" Then
		var5 = Trim(frm1.cboMatType.value)
		var6 = Trim(frm1.cboMatType.value)
	Else
		var5 = ""
		var6 = "ZZ"
	End If
	
	If Trim(frm1.txtItemCd1.value) <> "" Then
		var7 = Trim(frm1.txtItemCd1.value)
	Else
		var7 = ""
	End If

	If Trim(frm1.txtItemCd2.value) <> "" Then
		var8 = Trim(frm1.txtItemCd2.value)
	Else
		var8 = "ZZZZZZZZZZZZZZZZZZ"
	End If

	strUrl = strUrl & "plant|" & var1
	strUrl = strUrl & "|account|" & var2
	strUrl = strUrl & "|fr_purorg|" & var3
	strUrl = strUrl & "|to_purorg|" & var4
	strUrl = strUrl & "|fr_mattype|" & var5
	strUrl = strUrl & "|to_mattype|" & var6
	strUrl = strUrl & "|fr_item|" & var7
	strUrl = strUrl & "|to_item|" & var8

	ObjName = AskEBDocumentName("m1511oa1","ebr")

	Call FncEBRPreview(ObjName, strUrl)
End Function

<!--
'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
-->
Function FncExit()
    FncExit = True
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
    //alert(value)
    //alert(escape(value))
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST"> 
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>발주품목 기준정보 출력</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
	    		<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장"  NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">
													   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>													   
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>품목계정</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="cboAccount" ALT="품목계정" STYLE="Width: 98px;" tag="22"></SELECT></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>구매조직</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="구매조직"  NAME="txtPurOrg" SIZE=10 LANG="ko" MAXLENGTH=4 tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurOrg" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurOrg()">
													   <INPUT TYPE=TEXT NAME="txtPurOrgNm" SIZE=20 tag="14"></TD>													   
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>자재Type</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="cboMatType" ALT="자재Type" STYLE="Width: 98px;" tag="11"></SELECT></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" MAXLENGTH=18  SIZE=10 MAXLENGTH=10 ALT="품목" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd1()">
													   <INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=20 ALT="품목" tag="14"></TD>
								<TD CLASS="TD6" NOWRAP>~ <INPUT TYPE=TEXT NAME="txtItemCd2" MAXLENGTH=18  SIZE=10 MAXLENGTH=10 ALT="품목" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd2()">
													   <INPUT TYPE=TEXT NAME="txtItemNm2" SIZE=20 MAXLENGTH=18 ALT="품목" tag="14"></TD>					   
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex = -1></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
</BODY>
</HTML>
