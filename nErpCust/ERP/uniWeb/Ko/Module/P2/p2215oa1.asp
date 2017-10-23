<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : p2215oa1.asp
'*  4. Program Name         : MPS��û��� 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001-05-16
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jung Yu Kyung	
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment              :
'* 12. History              : Tracking No 9�ڸ����� 25�ڸ��� ����(2003.03.03)
'**********************************************************************************************-->
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop

Dim StartDate
Dim LastDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
LastDate =  UNIDateAdd("m",1,StartDate,parent.gDateFormat)

'=========================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE 
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    IsOpenPop = False     
End Sub

'=========================================================================================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtStartDt.Text = StartDate
	frm1.txtEndDt.Text = LastDate
End Sub

'=========================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	
    Call SetCombo(frm1.cboStatus , "RQ", "��û")
	Call SetCombo(frm1.cboStatus, "AC", "�ݿ�")
	
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6, i
	Dim iItemAcctArr, iItemAcctNmArr
	
	
    On Error Resume Next
    Err.Clear

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " AND MINOR_CD < " & FilterVar("30", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iItemAcctArr = Split(lgF0, Chr(11))
    iItemAcctNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iItemAcctArr) - 1
		Call SetCombo(frm1.cboAccount, UCase(iItemAcctArr(i)), iItemAcctNmArr(i))
	Next
	
End Sub

'=========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "OA") %>
End Sub

'=========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029
    
	Call ggoOper.FormatField(Document, "X",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")
    Call InitComboBox	
	Call SetDefaultVal
    Call InitVariables
    	
    Call SetToolbar("10000000000011")
    
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.cboStatus.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
	Set gActiveElement = document.activeElement   
End Sub

'=========================================================================================================
' Function Name : FncFind
' Function Desc : 
'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)
End Function

'=========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'=========================================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'=========================================================================================================
Function BtnPrint()
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
	Dim var7
	Dim var8
	Dim var9
	Dim var11
	Dim var13
	Dim var14
	
	Dim strUrl, strEbrFile, objName

	If frm1.txtFromItemCd.value = "" Then
		frm1.txtFromItemNm.value = "" 
	End If	
	
	If frm1.txtToItemCd.value = "" Then
		frm1.txtToItemNm.value = "" 
	End If	
	
    Call BtnDisabled(1)
	
	If Not chkfield(Document, "X") Then
		Call BtnDisabled(0)	
       Exit Function
    End If
    
	If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If

	var1 = UCase(Trim(frm1.txtPlantCd.value))	
	
	If frm1.txtTrackingNo1.value = "" Then
		var2 = "!"
	Else
		var2 = Trim(frm1.txtTrackingNo1.value)
	End If
	
	If frm1.txtTrackingNo2.value = "" Then
		var9 = "zzzzzzzzz"
	Else
		var9 = Trim(frm1.txtTrackingNo2.value)
	End If
	
	If frm1.cboStatus.value = "" Then
		var3 = "0"	
		var4 = "zz"
	Else
		var3 = frm1.cboStatus.value
		var4 = frm1.cboStatus.value	
	End If
	
	
	If frm1.txtFromItemCd.value = "" Then
		var5 = "0"
	Else
		var5 = frm1.txtFromItemCd.value  
	End If
	
	If frm1.txtToItemCd.value = "" Then
		var6 = "zzzzzzzzzzzzzzzzzz"
	Else
		var6 = frm1.txtToItemCd.value
	End If
	
	
	var7 = UNIConvDate(frm1.txtStartDt.Text)
	var8 = UNIConvDate(frm1.txtEndDt.Text)
	
	
	If frm1.txtItemGroupCd1.value = "" then 
		var11 = ""
	else
		var11 = "and b_item.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & _
				FilterVar(UCase(frm1.txtItemGroupCd1.value),"''","S") & "))"
	End If
	
	If frm1.cboAccount.value = "" then
		var13 = "0"
		var14 = "zz"
	else	
		var13 = Trim(frm1.cboAccount.value)
		var14 = Trim(frm1.cboAccount.value)
	End If
	
	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|fr_tracking_no|" & var2 
	strUrl = strUrl & "|to_tracking_no|" & var9
	strUrl = strUrl & "|fr_status|" & var3 
	strUrl = strUrl & "|to_status|" & var4 
	strUrl = strUrl & "|fr_item_cd|" & var5 
	strUrl = strUrl & "|to_item_cd|" & var6 
	strUrl = strUrl & "|fr_req_dt|" & var7 
	strUrl = strUrl & "|to_req_dt|" & var8
	strUrl = strUrl & "|cond|" & var11
	strUrl = strUrl & "|fr_item_acct|" & var13
	strUrl = strUrl & "|to_item_acct|" & var14
	
'----------------------------------------------------------------
' Print �Լ����� �߰��Ǵ� �κ� 
'----------------------------------------------------------------
	strEbrFile = "p2215oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")

	call FncEBRprint(EBAction, objName, strUrl)
'----------------------------------------------------------------
	
	Call BtnDisabled(0)	
	
End Function

'=========================================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'=========================================================================================================
Function BtnPreview() 
    
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
	Dim var7
	Dim var8
	Dim var9
	
	Dim var11
	Dim var13
	Dim var14
	
	Dim strUrl, strEbrFile, objName
	
	Call BtnDisabled(1)
	
	If Not chkfield(Document, "X") Then
		Call BtnDisabled(0)	
       Exit Function
    End If
    
	If ValidDateCheck(frm1.txtStartDt, frm1.txtEndDt) = False Then
		Call BtnDisabled(0)	
		Exit Function	
	End IF
	
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtFromItemCd.value = "" Then
		frm1.txtFromItemNm.value = "" 
	End If	
	
	If frm1.txtToItemCd.value = "" Then
		frm1.txtToItemNm.value = "" 
	End If	
	
	var1 = Trim(frm1.txtPlantCd.value)
		
	If frm1.txtTrackingNo1.value = "" Then
		var2 = "!"
	Else
		var2 = Trim(frm1.txtTrackingNo1.value)
	End If
	
	If frm1.txtTrackingNo2.value = "" Then
		var9 = "zzzzzzzzz"
	Else
		var9 = Trim(frm1.txtTrackingNo2.value)
	End If
	
	If frm1.cboStatus.value = "" Then
		var3 = "0"	
		var4 = "zz"
	Else
		var3 = frm1.cboStatus.value
		var4 = frm1.cboStatus.value	
	End If
	
	
	If frm1.txtFromItemCd.value = "" Then
		var5 = "0"
	Else
		var5 = frm1.txtFromItemCd.value  
	End If
	
	If frm1.txtToItemCd.value = "" Then
		var6 = "zzzzzzzzzzzzzzzzzz"
	Else
		var6 = frm1.txtToItemCd.value
	End If
	
	
	var7 = UNIConvDate(frm1.txtStartDt.Text)
	var8 = UNIConvDate(frm1.txtEndDt.Text)
	
	If frm1.txtItemGroupCd1.value = "" then 
		var11 = ""
	else
		var11 = "and b_item.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & _
				FilterVar(UCase(frm1.txtItemGroupCd1.value),"''","S") & "))"
	End If
	
	If frm1.cboAccount.value = "" then
		var13 = "0"
		var14 = "zz"
	else	
		var13 = Trim(frm1.cboAccount.value)
		var14 = Trim(frm1.cboAccount.value)
	End If
	
	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|fr_tracking_no|" & var2 
	strUrl = strUrl & "|to_tracking_no|" & var9
	strUrl = strUrl & "|fr_status|" & var3 
	strUrl = strUrl & "|to_status|" & var4 
	strUrl = strUrl & "|fr_item_cd|" & var5 
	strUrl = strUrl & "|to_item_cd|" & var6 
	strUrl = strUrl & "|fr_req_dt|" & var7 
	strUrl = strUrl & "|to_req_dt|" & var8
	strUrl = strUrl & "|cond|" & var11
	strUrl = strUrl & "|fr_item_acct|" & var13
	strUrl = strUrl & "|to_item_acct|" & var14
	
	strEbrFile = "p2215oa1"
	objName = AskEBDocumentName(strEbrFile,"ebr")

	call FncEBRPreview(objName, strUrl)
	
	Call BtnDisabled(0)	
	
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function



Function OpenPlantCd()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "�����˾�"		' �˾� ��Ī 
	arrParam(1) = "B_PLANT"				' TABLE ��Ī 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""					' Name Cindition
	arrParam(4) = ""					' Where Condition
	arrParam(5) = "����"			' TextBox ��Ī 
	
    arrField(0) = "PLANT_CD"			' Field��(0)
    arrField(1) = "PLANT_NM"			' Field��(1)
    
    arrHeader(0) = "����"			' Header��(0)
    arrHeader(1) = "�����"			' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlantCd(arrRet)
	End If	
	
End Function

'-------------------------------------  OpenItemGroup()  -------------------------------------------------
'	Name : OpenItemGroup()
'	Description : ItemGroup PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ǰ��׷��˾�"	
	arrParam(1) = "B_ITEM_GROUP"				
	arrParam(2) = Trim(frm1.txtItemGroupCd1.Value)
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & "  " 			
	arrParam(5) = "ǰ��׷�"			
	
    arrField(0) = "ITEM_GROUP_CD"	
    arrField(1) = "ITEM_GROUP_NM"	
    
    arrHeader(0) = "ǰ��׷�"		
    arrHeader(1) = "ǰ��׷��"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemGroup(arrRet)
	End If	
	
End Function

'-----------------------------------------  OpenTrackingInfo()  ------------------------------------------
'	Name : OpenTrackingInfo()
'	Description : OpenTracking Info PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenTrackingInfo1()
	Dim iCalledAspName, IntRetCD

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo1.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo1.value)
'	arrParam(2) = Trim(frm1.txtItemCd.value)
'	arrParam(3) = UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01")'frm1.txtPlanStartDt.Text
'	arrParam(4) = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31")'frm1.txtPlanEndDt.Text
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtTrackingNo1.Value = arrRet(0)
	End If
	
End Function


Function OpenTrackingInfo2()
	Dim iCalledAspName, IntRetCD

	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "����","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo2.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo2.value)
'	arrParam(2) = Trim(frm1.txtItemCd.value)
'	arrParam(3) = UniConvYYYYMMDDToDate(parent.gDateFormat, "1900", "01", "01")'frm1.txtPlanStartDt.Text
'	arrParam(4) = UniConvYYYYMMDDToDate(parent.gDateFormat, "2999", "12", "31")'frm1.txtPlanEndDt.Text
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtTrackingNo2.Value = arrRet(0)
	End If
	
End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenFromItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If		

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtfromItemCd.Value		' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetFromItemCd(arrRet)
	End If	

End Function

'------------------------------------------  OpenItemInfo()  -------------------------------------------------
'	Name : OpenToItemCd()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenToItemCd()
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD

	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("189220", "x", "x", "x")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If		

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtToItemCd.value		' Item Code
	arrParam(2) = "12!MO"						' Combo Set Data:"1020!MP" -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field��(0) :"ITEM_CD"
    arrField(1) = 2 							' Field��(1) :"ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
    
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetToItemCd(arrRet)
	End If	

End Function


Function SetPlantCd(ByRef arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)
	frm1.txtPlantCd.focus
	Set gActiveElement = document.activeElement
End Function

Function SetFromItemCd(ByRef arrRet)
	frm1.txtFromItemCd.value = arrRet(0)
	frm1.txtFromItemNm.value = arrRet(1)
	frm1.txtFromItemCd.focus
	Set gActiveElement = document.activeElement
End Function

Function SetToItemCd(ByRef arrRet)
	frm1.txtToItemCd.value = arrRet(0)
	frm1.txtToItemNm.value = arrRet(1)
	frm1.txtToItemCd.focus
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  SetItemGroup()  --------------------------------------------------
'	Name : SetItemGroup()
'	Description : ItemGroup Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetItemGroup(ByRef arrRet)
	frm1.txtItemGroupCd1.Value    = arrRet(0)
	frm1.txtItemGroupCd1.focus
	Set gActiveElement = document.activeElement	
End Function

'=======================================================================================================
'   Event Name : txtValidFromDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtStartDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtStartDt.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtStartDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtEndDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtEndDt.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtEndDt.Focus
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	

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
		<TD HEIGHT=5 colspan="2">&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100% colspan="2">
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>MPS��û���</font></td>
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
		<TD WIDTH=100% CLASS="Tab11" colspan="2">
			<TABLE CLASS="BasicTB" CELLSPACING=0 >	
	    		<TR>
					<TD HEIGHT=10 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="X2XXXU" ALT="����"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="X4" ALT="�����">&nbsp;
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��û��</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtStartDt CLASS=FPDTYYYYMMDD title=FPDATETIME SIZE="10" MAXLENGTH="10" ALT="������" tag="X2" VIEWASTEXT > </OBJECT>');</SCRIPT>								
										&nbsp;~&nbsp; 
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtEndDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="X2" ALT="������" MAXLENGTH="10" SIZE="10" VIEWASTEXT > </OBJECT>');</SCRIPT>								
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>��û����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboStatus" ALT="��û����" STYLE="Width: 160px;" tag="X1"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>	
								<TR>
								    <TD CLASS="TD5" NOWRAP>ǰ��׷�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtItemGroupCd1" SIZE=19 MAXLENGTH=10 tag="X1XXXU" ALT="ǰ��׷�"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()"> </TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>ǰ�����</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboAccount" ALT="ǰ�����" STYLE="Width: 160px;" tag="X1"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>ǰ��</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtFromItemCd" SIZE=18 MAXLENGTH=18 tag="X1XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtFromItemNm" SIZE=40 MAXLENGTH=40 tag="X4" ALT="ǰ���">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtToItemCd" SIZE=18 MAXLENGTH=18 tag="X1XXXU" ALT="ǰ��"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToItemCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtToItemNm" SIZE=40 MAXLENGTH=40 tag="X4" ALT="ǰ���">&nbsp;
									</TD>
								</TR>
								<TR>
								    <TD CLASS="TD5" NOWRAP>Tracking No.</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtTrackingNo1" SIZE=25 MAXLENGTH=25 tag="X1XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingInfo1">&nbsp;~&nbsp;
									<INPUT TYPE=TEXT NAME="txtTrackingNo2" SIZE=25 MAXLENGTH=25 tag="X1XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnProdOrder" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingInfo2">
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
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
				     <TD WIDTH = 10 > &nbsp; </TD>
				     <TD>
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>�μ�</BUTTON>
                     </TD> 		
 		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>

