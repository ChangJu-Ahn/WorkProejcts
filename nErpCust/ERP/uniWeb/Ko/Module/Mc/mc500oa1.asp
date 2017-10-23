<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : MC500OA1
'*  4. Program Name         : 납입지시서출력 
'*  5. Program Desc         : 납입지시서출력 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/27
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Lee Woo Guen
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
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit	

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
Dim IsOpenPop

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE 
    lgBlnFlgChgValue = False  
    lgIntGrpCount = 0         
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False            
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim CurDoDate
	CurDoDate   = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
	frm1.txtDoDate.Text	= CurDoDate
	Set gActiveElement = document.activeElement
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'=========================================================================================================
Sub InitComboBox()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    	
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("M2110", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboDoTime, lgF0, lgF1, Chr(11))
   
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","OA") %>
End Sub

'------------------------------------------  OpenPlantCd()  -------------------------------------------------
'	Name : OpenPlantCd()
'	Description : Condition Plant PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"				' 팝업 명칭 
	arrParam(1) = "B_PLANT"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)	' Code Condition
	arrParam(3) = ""'Trim(frm1.txtPlantNm.Value)	' Name Cindition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "공장"					' TextBox 명칭 
	
	arrField(0) = "PLANT_CD"					' Field명(0)
	arrField(1) = "PLANT_NM"					' Field명(1)
	
	arrHeader(0) = "공장"					' Header명(0)
	arrHeader(1) = "공장명"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)
		frm1.txtPlantCd.focus    	
		Set gActiveElement = document.activeElement	
	End If
End Function

'------------------------------------------  OpenFromBpCd()  -------------------------------------------------
'	Name : OpenFromBpCd()
'	Description : Biz Partner PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenFromBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtFromBpCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtFromBpCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "	
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "공급처"				
    arrHeader(1) = "공급처명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtFromBpCd.focus
		Exit Function
	Else
		frm1.txtFromBpCd.Value    = arrRet(0)		
		frm1.txtFromBpNm.Value    = arrRet(1)	
		frm1.txtFromBpCd.focus		
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenToBpCd()  -------------------------------------------------
'	Name : OpenToBpCd()
'	Description : Biz Partner PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenToBpCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtToBpCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtToBpCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "	
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "공급처"				
    arrHeader(1) = "공급처명"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtToBpCd.focus		
		Exit Function
	Else
		frm1.txtToBpCd.Value    = arrRet(0)		
		frm1.txtToBpNm.Value    = arrRet(1)	
		frm1.txtToBpCd.focus		
		Set gActiveElement = document.activeElement
	End If	
	
End Function

'------------------------------------------  OpenFromItemCd()  -------------------------------------------------
'	Name : OpenFromItemCd()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenFromItemCd()
	Dim iCalledAspName
	Dim arrParam(5), arrField(2)
	
	Dim IntRetCD
	Dim arrRet

	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")    '공장정보가 필요합니다 
		frm1.txtPlantCd.focus
		Exit Function
	End If

    '공장 체크 함수 호출 
	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	'------------------------------------------------------
	
	If IsOpenPop = True Then Exit Function	

	IsOpenPop = True

	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA3","X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtFromItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtFromItemCd.focus
		Exit Function
	Else
		frm1.txtFromItemCd.value = arrRet(0)
		frm1.txtFromItemNm.value = arrRet(1) 
		frm1.txtFromItemCd.focus
		Set gActiveElement = document.activeElement	End If
End Function


'------------------------------------------  OpenToItemCd()  -------------------------------------------------
'	Name : OpenToItemCd()
'	Description : Item By Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenToItemCd()
	Dim iCalledAspName
	Dim arrParam(5), arrField(2)
	
	Dim IntRetCD
	Dim arrRet

	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("169901","X", "X", "X")    '공장정보가 필요합니다 
		frm1.txtPlantCd.focus
		Exit Function
	End If

    '공장 체크 함수 호출 
	If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Exit function
	End If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)
	'------------------------------------------------------
	
	If IsOpenPop = True Then Exit Function	

	IsOpenPop = True

	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION,"B1B11PA3","X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtToItemCd.value)		' Item Code
	arrParam(2) = "!"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec

	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtToItemCd.focus
		Exit Function
	Else
		frm1.txtToItemCd.value = arrRet(0)
		frm1.txtToItemNm.value = arrRet(1) 
		frm1.txtToItemCd.focus
		Set gActiveElement = document.activeElement
	End If
End Function

'******************************************  3.1 Window 처리  *********************************************
'	Window에 발생 하는 모든 Even 처리	
'*********************************************************************************************************
 Sub txtDoDate_DblClick(Button)
	if Button = 1 then
		frm1.txtDoDate.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtDoDate.Focus
	End if
End Sub

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'=========================================================================================================
 Sub Form_Load()
    Call LoadInfTB19029                                                     

    Call ggoOper.FormatField(Document, "X",ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")   
    
	Call SetDefaultVal
	Call InitVariables														'⊙: Initializes local global variables
 	Call InitComboBox
    
    Call SetToolbar("1000000000001111")										
    
	 'Plant Code, Plant Name Setting 
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
	
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement 
	Else
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 	
	End If
End Sub

'========================================================================================
' Function Name : BtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function BtnPrint()
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	Dim var6
	Dim var7
	Dim var8
	
	Dim strUrl, strEbrFile, objName
	
	Call BtnDisabled(1)
	
	If Not chkfield(Document, "X") Then							'⊙: This function check indispensable field
		Call BtnDisabled(0)	
		Exit Function
    End If
	    
	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If

	If frm1.txtFromBpCd.value = "" Then
		frm1.txtFromBpNm.value = "" 
	End If	
	
	If frm1.txtToBpCd.value = "" Then
		frm1.txtToBpNm.value = "" 
	End If		
	
	If frm1.txtFromItemCd.value = "" Then
		frm1.txtFromItemNm.value = "" 
	End If	
	
	If frm1.txtToItemCd.value = "" Then
		frm1.txtToItemNm.value = "" 
	End If	
	
    Call BtnDisabled(1)
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UNIConvDate(frm1.txtDoDate.text)
	var3 = Trim(frm1.cboDoTime.value)
	
	If frm1.txtFromBpCd.value = "" Then
		var4 = "0"
	Else
		var4 = frm1.txtFromBpCd.value  
	End If
	
	If frm1.txtToBpCd.value = "" Then
		var5 = "zzzzzzzzzzzzzzzzzz"
	Else
		var5 = frm1.txtToBpCd.value
	End If
	
	If frm1.txtFromItemCd.value = "" Then
		var6 = "0"
	Else
		var6 = frm1.txtFromItemCd.value  
	End If
	
	If frm1.txtToItemCd.value = "" Then
		var7 = "zzzzzzzzzzzzzzzzzz"
	Else
		var7 = frm1.txtToItemCd.value
	End If

	strUrl = "plant_cd|" & var1 	
	strUrl = strUrl & "|do_date|" & var2
	strUrl = strUrl & "|do_time|" & var3
	strUrl = strUrl & "|fr_bp_cd|" & var4 
	strUrl = strUrl & "|to_bp_cd|" & var5
	strUrl = strUrl & "|fr_item_cd|" & var6 
	strUrl = strUrl & "|to_item_cd|" & var7
	
	strEbrFile = "mc500oa1"	
	objName = AskEBDocumentName(strEbrFile,"ebr")

	call FncEBRprint(EBAction, objName, strUrl)
	
	Call BtnDisabled(0)	
	
End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview() 
	Dim var1
	Dim var2
	Dim var3
	Dim var4
	Dim var5
	dim var6
	Dim var7
	Dim var8
	
	Dim strUrl, strEbrFile, objName

	Call BtnDisabled(1)
	
	If Not chkfield(Document, "X") Then									'⊙: This function check indispensable field
		Call BtnDisabled(0)	
		Exit Function
    End If

	If frm1.txtPlantCd.value= "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtFromBpCd.value = "" Then
		frm1.txtFromBpNm.value = "" 
	End If	
	
	If frm1.txtToBpCd.value = "" Then
		frm1.txtToBpNm.value = "" 
	End If		
	
	If frm1.txtFromItemCd.value = "" Then
		frm1.txtFromItemNm.value = "" 
	End If	
	
	If frm1.txtToItemCd.value = "" Then
		frm1.txtToItemNm.value = "" 
	End If	
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UNIConvDate(frm1.txtDoDate.text)
	var3 = Trim(frm1.cboDoTime.value)
	
	If frm1.txtFromBpCd.value = "" Then
		var4 = "0"
	Else
		var4 = frm1.txtFromBpCd.value  
	End If
	
	If frm1.txtToBpCd.value = "" Then
		var5 = "zzzzzzzzzzzzzzzzzz"
	Else
		var5 = frm1.txtToBpCd.value
	End If
	
	If frm1.txtFromItemCd.value = "" Then
		var6 = "0"
	Else
		var6 = frm1.txtFromItemCd.value  
	End If
	
	If frm1.txtToItemCd.value = "" Then
		var7 = "zzzzzzzzzzzzzzzzzz"
	Else
		var7 = frm1.txtToItemCd.value
	End If

	strUrl = "plant_cd|" & var1 	
	strUrl = strUrl & "|do_date|" & var2
	strUrl = strUrl & "|do_time|" & var3
	strUrl = strUrl & "|fr_bp_cd|" & var4 
	strUrl = strUrl & "|to_bp_cd|" & var5
	strUrl = strUrl & "|fr_item_cd|" & var6 
	strUrl = strUrl & "|to_item_cd|" & var7
	
	strEbrFile = "mc500oa1"	
	objName = AskEBDocumentName(strEbrFile,"ebr")

	call FncEBRPreview(objName, strUrl)
	Call BtnDisabled(0)	
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
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)  
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncExit = True

End Function

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
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../SChared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>납입지시서출력</font></td>
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
		<TD WIDTH=100% CLASS="Tab11" HEIGHT=* colspan="2">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>	
							    <TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=15 MAXLENGTH=4 tag="X2XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=40 tag="X4" ALT="공장명">
								</TD>
							</TR>								   
							<TR>
								<TD CLASS="TD5" NOWRAP>납입지시일</TD>
								<TD CLASS="TD6">
									<script language =javascript src='./js/mc500oa1_I924395608_txtDoDate.js'></script>&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>납입지시시간</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="cboDoTime" ALT="납입지시시간" STYLE="Width: 100px;" tag="X2XXXU"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFromBpCd" SIZE=12 MAXLENGTH=10 tag="X1XXXU" ALT="공급처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromBpCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromBpCd()">
													   <INPUT TYPE=TEXT NAME="txtFromBpNm" SIZE=30 tag="X4" ALT="공급처명">&nbsp;~&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToBpCd" SIZE=12 MAXLENGTH=10 tag="X1XXXU" ALT="공급처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToBpCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToBpCd()">
													   <INPUT TYPE=TEXT NAME="txtToBpNm" SIZE=30 tag="X4" ALT="공급처명">
							    </TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtFromItemCd" SIZE=20 MAXLENGTH=18 tag="X1XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFromItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenFromItemCd()">
													   <INPUT TYPE=TEXT NAME="txtFromItemNm" SIZE=30 tag="X4" ALT="품목명">&nbsp;~&nbsp;
								</TD>
							</TR>
							<TR>	
							    <TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtToItemCd" SIZE=20 MAXLENGTH=18 tag="X1XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnToItemCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenToItemCd()">
													   <INPUT TYPE=TEXT NAME="txtToItemNm" SIZE=30 tag="X4" ALT="품목명">&nbsp;
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
				     <TD WIDTH = 10 > &nbsp; </TD>
				     <TD>
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
		               <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON>
                     </TD> 
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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
