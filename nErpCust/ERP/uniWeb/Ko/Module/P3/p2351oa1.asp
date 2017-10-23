<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : 
'*  2. Function Name        : 
'*  3. Program ID           : p2351oa1.asp
'*  4. Program Name         : MRP예시결과출력 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002-01-07
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Jung Yu Kyung	
'* 10. Modifier (Last)      : Jung Yu Kyung
'* 11. Comment              :
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

'========================================================================================================= 
<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim IsOpenPop

'=========================================================================================================
Dim EndDate
Dim ToDate

EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
ToDate =  UNIDateAdd("m",1,EndDate,parent.gDateFormat)

'=========================================================================================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    IsOpenPop = False     
End Sub

'=========================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtFromDt.Text = EndDate
	frm1.txtToDt.Text = ToDate
	
End Sub

'=========================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()

	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6, i
	Dim iProdMgrArr, iProdMgrNmArr
	
    On Error Resume Next
    Err.Clear

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1015", "''", "S") & " ORDER BY MINOR_CD ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
    iProdMgrArr = Split(lgF0, Chr(11))
    iProdMgrNmArr = Split(lgF1, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.Description 
		Err.Clear 
		Exit Sub
	End If

	For i = 0 to UBound(iProdMgrArr) - 1
		Call SetCombo(frm1.cboProdMgr, UCase(iProdMgrArr(i)), iProdMgrNmArr(i))
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
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
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
		frm1.cboProdMgr.focus 
		Set gActiveElement = document.activeElement
	Else
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
	End If
    
End Sub

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

Sub rdoPrintType1_OnClick()
	frm1.txtPurOrg.value = ""
	Call ggoOper.SetReqAttr(frm1.cboProdMgr,"D")
	Call ggoOper.SetReqAttr(frm1.txtPurOrg ,"Q")
End Sub

Sub rdoPrintType2_OnClick()
	frm1.cboProdMgr.value = ""
	Call ggoOper.SetReqAttr(frm1.cboProdMgr,"Q")
	Call ggoOper.SetReqAttr(frm1.txtPurOrg ,"D")
End Sub

Sub rdoPrintType3_OnClick()
	frm1.txtPurOrg.value = ""
	frm1.cboProdMgr.value = ""
	Call ggoOper.SetReqAttr(frm1.cboProdMgr,"Q")
	Call ggoOper.SetReqAttr(frm1.txtPurOrg ,"Q")
	
End Sub

'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function BtnPrint()
	Dim var1
	Dim var2
	Dim var4
	Dim var5
	Dim var6
	Dim var7
	Dim var8
	
	Dim strUrl, strEbrFile, objName
	
	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt)  = False Then		
		frm1.txtFromDt.focus 
		Exit Function
	End If
	
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
    
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If

	var1 = UCase(Trim(frm1.txtPlantCd.value))
	
	If frm1.rdoPrintType1.checked = True Then
		If frm1.cboProdMgr.value = "" Then
			var2 = "0"
			strEbrFile = "p2351oa1"
			var8 = "zz"
		Else
			var2 = Trim(frm1.cboProdMgr.value)
			strEbrFile = "p2351oa1"
			var8 = Trim(frm1.cboProdMgr.value)
		End If
		
	ElseIf frm1.rdoPrintType2.checked = True Then
		If frm1.txtPurOrg.value = "" Then
			var2 = "0"
			strEbrFile = "p2351oa2"
			var8 = "zzzz"
		Else
			var2 = Trim(frm1.txtPurOrg.value)
			strEbrFile = "p2351oa2"
			var8 = Trim(frm1.txtPurOrg.value)
		End If
	Else
		strEbrFile = "p2351oa3"
	End If
	
	If frm1.txtFromItemCd.value = "" Then
		var4 = "0"
	Else
		var4 = frm1.txtFromItemCd.value  
	End If
	
	If frm1.txtToItemCd.value = "" Then
		var5 = "zzzzzzzzzzzzzzzzzz"
	Else
		var5 = frm1.txtToItemCd.value
	End If
	
	var6 = UniConvDateAToB(frm1.txtFromDt.text, parent.gDateFormat, parent.gServerDateFormat)
	var7 = UniConvDateAToB(frm1.txtToDt.text, parent.gDateFormat, parent.gServerDateFormat)
	
	strUrl = "plant_cd|" & var1 	
	strUrl = strUrl & "|factor1|" & var2
	strUrl = strUrl & "|factor2|" & var8
	strUrl = strUrl & "|fr_item_cd|" & var4 
	strUrl = strUrl & "|to_item_cd|" & var5
	strUrl = strUrl & "|from_dt|" & var6 
	strUrl = strUrl & "|to_dt|" & var7 
	
'----------------------------------------------------------------
' Print 함수에서 추가되는 부분 
'----------------------------------------------------------------
	objName = AskEBDocumentName(strEbrFile,"ebr")

	call FncEBRprint(EBAction, objName, strUrl)
'----------------------------------------------------------------
	
	Call BtnDisabled(0)	
	
End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview()
    
	Dim var1
	Dim var2
	Dim var4
	Dim var5
	dim var6
	Dim var7
	Dim var8
	
	Dim strUrl, strEbrFile, objName

	If Not chkfield(Document, "X") Then
		Call BtnDisabled(0)	
       Exit Function
    End If
    	
    If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt)  = False Then		
		frm1.txtFromDt.focus 
		Exit Function
	End If
	
	Call BtnDisabled(1)
	
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


	If frm1.rdoPrintType1.checked = True Then
		If frm1.cboProdMgr.value = "" Then
			var2 = "0"
			strEbrFile = "p2351oa1"
			var8 = "zz"
		Else
			var2 = Trim(frm1.cboProdMgr.value)
			strEbrFile = "p2351oa1"
			var8 = Trim(frm1.cboProdMgr.value)
		End If
		
	ElseIf frm1.rdoPrintType2.checked = True Then
		If frm1.txtPurOrg.value = "" Then
			var2 = "0"
			strEbrFile = "p2351oa2"
			var8 = "zzzz"
		Else
			var2 = Trim(frm1.txtPurOrg.value)
			strEbrFile = "p2351oa2"
			var8 = Trim(frm1.txtPurOrg.value)
		End If
	Else
		strEbrFile = "p2351oa3"
	End If
			
	
	If frm1.txtFromItemCd.value = "" Then
		var4 = "0"
	Else
		var4 = frm1.txtFromItemCd.value  
	End If
	
	If frm1.txtToItemCd.value = "" Then
		var5 = "zzzzzzzzzzzzzzzzzz"
	Else
		var5 = frm1.txtToItemCd.value
	End If
	
	var6 = UniConvDateAToB(frm1.txtFromDt.text, parent.gDateFormat, parent.gServerDateFormat)
	var7 = UniConvDateAToB(frm1.txtToDt.text, parent.gDateFormat, parent.gServerDateFormat)
	
	strUrl = "plant_cd|" & var1 
	strUrl = strUrl & "|factor1|" & var2
	strUrl = strUrl & "|factor2|" & var8
	strUrl = strUrl & "|fr_item_cd|" & var4 
	strUrl = strUrl & "|to_item_cd|" & var5  
	strUrl = strUrl & "|from_dt|" & var6 
	strUrl = strUrl & "|to_dt|" & var7 
	
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

	arrParam(0) = "공장팝업"		' 팝업 명칭 
	arrParam(1) = "B_PLANT"				' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""					' Name Cindition
	arrParam(4) = ""					' Where Condition
	arrParam(5) = "공장"			' TextBox 명칭 
	
    arrField(0) = "PLANT_CD"			' Field명(0)
    arrField(1) = "PLANT_NM"			' Field명(1)
    
    arrHeader(0) = "공장"			' Header명(0)
    arrHeader(1) = "공장명"			' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlantCd(arrRet)
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
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If		

	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtfromItemCd.Value		' Item Code
	arrParam(2) = ""						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) :"ITEM_CD"
    arrField(1) = 2 							' Field명(1) :"ITEM_NM"
    
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

	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = frm1.txtToItemCd.value		' Item Code
	arrParam(2) = ""						' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""							' Default Value
	
    arrField(0) = 1 							' Field명(0) :"ITEM_CD"
    arrField(1) = 2 							' Field명(1) :"ITEM_NM"
    
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
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPurOrg(arrRet)
	End If	
	
End Function

'------------------------------------------  SetPurOrg()  --------------------------------------------------
'	Name : SetPurOrg()
'	Description : PurOrg Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPurOrg(ByRef arrRet)
	frm1.txtPurOrg.Value    = arrRet(0)
	frm1.txtPurOrgNm.Value    = arrRet(1)
	frm1.txtPurOrg.focus
    Set gActiveElement = document.activeElement	
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

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtToDt.Focus
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>MRP예시결과출력</font></td>
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
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=2>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="X2XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=40 MAXLENGTH=40 tag="X4" ALT="공장명">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>시작일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtFromDt CLASS=FPDTYYYYMMDD title=FPDATETIME SIZE="10" MAXLENGTH="10" ALT="시작일" tag="X2" VIEWASTEXT id=OBJECT1> </OBJECT>');</SCRIPT>								
										&nbsp;~&nbsp; 
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtToDt CLASS=FPDTYYYYMMDD title=FPDATETIME tag="X2" ALT="종료일" MAXLENGTH="10" SIZE="10" VIEWASTEXT id=OBJECT2> </OBJECT>');</SCRIPT>								
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>생산담당자</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboProdMgr" ALT="생산담당" STYLE="Width: 98px;" tag="X1"><OPTION VALUE=""></OPTION></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>구매조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurOrg" SIZE=15 MAXLENGTH=4 tag="X4XXXU" ALT="구매조직"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurOrg" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPurOrg()">&nbsp;<INPUT TYPE=HIDDEN NAME="txtPurOrgNm" SIZE=40 tag="X4"></TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtFromItemCd" SIZE=18 MAXLENGTH=18 tag="X1XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromItemCd()">&nbsp;
									<INPUT TYPE=TEXT NAME="txtFromItemNm" SIZE=40 MAXLENGTH=40 tag="X4" ALT="품목명">&nbsp;~&nbsp;
									</TD>
								</TR>
								<TR>	
								    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>
									<INPUT TYPE=TEXT NAME="txtToItemCd" SIZE=18 MAXLENGTH=18 tag="X1XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToItemCd()">&nbsp;
									<INPUT TYPE=TEXT NAME="txtToItemNm" SIZE=40 MAXLENGTH=40 tag="X4" ALT="품목명">&nbsp;
									</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=10 WIDTH=100%>
					    <FIELDSET CLASS="CLSFLD">
					        <TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=2>
								<TR>
									<TD CLASS=TD5 NOWRAP>출력방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType1" CLASS="RADIO" CHECKED tag="X1"><LABEL FOR="rdoPrintType1">생산</LABEL>
													     <INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType2" CLASS="RADIO" tag="X1"><LABEL FOR="rdoPrintType2">구매</LABEL>
														 <INPUT TYPE="RADIO" NAME="rdoPrintType" ID="rdoPrintType3" CLASS="RADIO" tag="X1"><LABEL FOR="rdoPrintType3">전체</LABEL></TD>
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
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON>
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
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
</BODY>
</HTML>
