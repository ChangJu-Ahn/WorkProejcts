<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Quality Management
'*  2. Function Name        : 
'*  3. Program ID           : Q1611OA1
'*  4. Program Name         : 검사기준서 출력 
'*  5. Program Desc         : Quality Configuration
'*  6. Component List       : PQBG010,PD6G020
'*  7. Modified date(First) : 2002/05/14
'*  8. Modified date(Last)  : 2004/07/06
'*  9. Modifier (First)     : Koh Jae Woo
'* 10. Modifier (Last)      : Hong Chang Ho
'* 11. Comment
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit																	'☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop         

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    IsOpenPop = False     
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.cboInspClassCd.value = "R"
	Call ggoOper.SetReqAttr(frm1.txtRoutingNo, "Q")
	Call ggoOper.SetReqAttr(frm1.txtOprNo, "Q")
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "Q", "NOCOOKIE","OA") %>
End Sub

'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
    Dim strCboCd 
    Dim strCboNm 

	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = 'Q0001' ORDER BY MINOR_CD", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboInspClassCd , lgF0, lgF1, Chr(11))		
End Sub

'------------------------------------------  OpenPlant() -------------------------------------------------
'	Name : OpenPlant()
'	Description :Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlant()
	OpenPlant =false
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			

    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	

    arrHeader(0) = "공장코드"		
    arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam,arrField, arrHeader), _
		"dialogWIDTH=420px; dialogHEIGHT=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)		
	End If	
	
	frm1.txtPlantCd.Focus
	Set gActiveElement = document.activeElement
	OpenPlant = true
End Function


 '------------------------------------------  OpenItem()  -------------------------------------------------
'	Name : OpenItem()
'	Description : Item PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItem()
	OpenItem = false
	
	Dim arrRet
	Dim arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrParam6
	Dim arrField(6)
	Dim iCalledAspName, IntRetCD

	'공장코드가 있는 지 체크 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("220705", "X", "X", "X") 		'공장정보가 필요합니다 
		frm1.txtPlantCd.Focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	
	arrParam1 = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam2 = Trim(frm1.txtPlantNm.Value)	' Plant Name
	arrParam3 = Trim(frm1.txtItemCd.Value)	' Item Code
	arrParam4 = ""	'Trim(frm1.txtItemNm.Value)	' Item Name
	arrParam5 = Trim(frm1.cboInspClassCd.Value)
	arrParam6 = "HAVE_STANDARD"
	
	iCalledAspName = AskPRAspName("q1211pa2")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "q1211pa2", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.Parent, arrParam1, arrParam2, arrParam3, arrParam4, arrParam5, arrParam6, arrField), _
	              "dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		  
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtItemCd.Value    = arrRet(0)		
		frm1.txtItemNm.Value    = arrRet(1)		
	End If	

	frm1.txtItemCd.Focus
	Set gActiveElement = document.activeElement
	OpenItem = true
End Function

'------------------------------------------  OpenRoutingNo()  -------------------------------------------------
'	Name : OpenRoutingNo()
'	Description : Routing No PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenRoutingNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtRoutingNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "품목", "X")
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True

	arrParam(0) = "라우팅 팝업"	
	arrParam(1) = "P_ROUTING_HEADER"				
	arrParam(2) = Trim(frm1.txtRoutingNo.Value)
	arrParam(3) = ""
	arrParam(4) =  "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S") & " And ITEM_CD = " & FilterVar(frm1.txtItemCd.value, "''", "S")
	arrParam(5) = "라우팅"			

    arrField(0) = "ROUT_NO"	
    arrField(1) = "DESCRIPTION"	
    arrField(2) = "BOM_NO"
    arrField(3) = "MAJOR_FLG"

    arrHeader(0) = "라우팅"		
    arrHeader(1) = "라우팅명"		
    arrHeader(2) = "BOM Type"
    arrHeader(3) = "주라우팅"
    
    	
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetRoutingNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtRoutingNo.focus
	
End Function


'------------------------------------------  OpenOprNo()  -------------------------------------------------
'	Name : OpenOprNo()
'	Description : Opr No. Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenOprNo()
	Dim arrRet
	Dim arrParam(6), arrField(6), arrHeader(6)
	Dim str
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "공장", "X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If frm1.txtItemCd.value = "" Then
		Call DisplayMsgBox("971012", "X", "품목", "X")
		frm1.txtItemCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If frm1.txtRoutingNo.value = "" Then
		Call DisplayMsgBox("971012", "X", "라우팅", "X")
		frm1.txtRoutingNo.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	If IsOpenPop = True Then 
		IsOpenPop = False
		Exit Function
	End If

	IsOpenPop = True

	str = frm1.txtOprNo.value
	
	arrParam(0) = "공정팝업"	
	arrParam(1) = "P_ROUTING_DETAIL"				
	arrParam(2) = str
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD =" & FilterVar(frm1.txtPlantCd.value, "''", "S") & _
				  " AND ITEM_CD = " & FilterVar(frm1.txtItemCd.value, "''", "S") & _
				  " AND ROUT_NO = " & FilterVar(frm1.txtRoutingNo.value, "''", "S")
	arrParam(5) = "공정"			
	
    arrField(0) = "OPR_NO"
    arrField(1) = "JOB_CD"
    arrField(2) = "MILESTONE_FLG"
    
    arrHeader(0) = "공정"
    arrHeader(1) = "작업코드"
    arrHeader(2) = "Milestone"
    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		Call SetOprNo(arrRet)
	End If	
	
	Call SetFocusToDocument("M")
	frm1.txtOprNo.focus
	
End Function

'------------------------------------------  SetRouting()  --------------------------------------------------
'	Name : SetRoutingNo()
'	Description : Routing No Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetRoutingNo(byval arrRet)
	frm1.txtRoutingNo.Value    = arrRet(0)
	frm1.txtRoutingNm.Value    = arrRet(1)
End Function

'------------------------------------------  SetWcCd()  --------------------------------------------------
'	Name : SetOprNo()
'	Description : Work Center Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetOprNo(Byval arrRet)
	frm1.txtOprNo.value	= arrRet(0)
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")       
	Call InitVariables                                                      '⊙: Initializes local global variables
	
	 '----------  Coding part  -------------------------------------------------------------
	Call InitComboBox
	Call SetDefaultVal
	
	Call SetToolbar("10000000000011")										'⊙: 버튼 툴바 제어 
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Ucase(Parent.gPlant)
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtItemCd.focus 
    Else
		frm1.txtPlantCd.focus 
    End If
End Sub

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
	FncFind = False
    Call parent.FncFind(Parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
    FncFind = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	FncPrint = False
	Call Parent.FncPrint()
	FncPrint = True
End Function

'========================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function FncBtnPrint() 
	dim var1, var2, var3, var4, var5
	dim strUrl
	Dim strEbrFile, objName
	
	FncBtnPrint = false
	
	If Not chkField(Document, "1") Then	Exit Function

	If Plant_Item_Check = False Then Exit Function	
	
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UCase(Trim(frm1.cboInspClassCd.value))
	var3 = UCase(Trim(frm1.txtItemCd.value))
	var4 = UCase(Trim(frm1.txtRoutingNo.value))
	var5 = UCase(Trim(frm1.txtOprNo.value))
	
	' 연보의 조회타입 선택(품목별,공급처별,품목/공급처별) 
	
	strUrl = strUrl & "PlantCd|" & var1
	strUrl = strUrl & "|InspClassCd|" & var2 
	strUrl = strUrl & "|ItemCd|" & var3

	If var2 = "R" Then
		strEbrFile = "Q1611OA11"
	ElseIf var2 = "P" Then
		strEbrFile = "Q1611OA12"
		strUrl = strUrl & "|RoutNo|" & var4 
		strUrl = strUrl & "|OprNo|" & var5
	ElseIf var2 = "F" Then
		strEbrFile = "Q1611OA13"
	Else
		strEbrFile = "Q1611OA14"
	End if

	
	objName = AskEBDocumentName(strEbrFile, "EBR")

	call FncEBRprint(EBAction, objName, strUrl)
	
	FncBtnPrint = true	
End Function

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
Function FncQuery()
	FncQuery = true
	If FncBtnPreview = False Then Exit function
	FncQuery = true
End Function

'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function FncBtnPreview() 
	Dim var1, var2, var3, var4, var5
	Dim strUrl
	Dim strEbrFile, objName
	
	FncBtnPreview = false
			
	If Not chkField(Document, "1") Then Exit Function

	If Plant_Item_Check = False Then Exit Function
		
	var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = UCase(Trim(frm1.cboInspClassCd.value))
	var3 = UCase(Trim(frm1.txtItemCd.value))
	var4 = UCase(Trim(frm1.txtRoutingNo.value))
	var5 = UCase(Trim(frm1.txtOprNo.value))
	
	' 연보의 조회타입 선택(품목별,공급처별,품목/공급처별) 
	
	strUrl = strUrl & "PlantCd|" & var1
	strUrl = strUrl & "|InspClassCd|" & var2 
	strUrl = strUrl & "|ItemCd|" & var3

	If var2 = "R" Then
		strEbrFile = "Q1611OA11"
	ElseIf var2 = "P" Then
		strEbrFile = "Q1611OA12"
		strUrl = strUrl & "|RoutNo|" & var4 
		strUrl = strUrl & "|OprNo|" & var5
	ElseIf var2 = "F" Then
		strEbrFile = "Q1611OA13"
	Else
		strEbrFile = "Q1611OA14"
	End if
	
	objName = AskEBDocumentName(strEbrFile, "EBR")

	call FncEBRPreview(objName, strUrl)
	
	FncBtnPreview = true			
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : Plant_Item_Check
'========================================================================================
Function Plant_Item_Check()

	Plant_Item_Check = False
	
	With frm1
	
 		If  CommonQueryRs(" B.ITEM_NM, C.PLANT_NM "," B_ITEM_BY_PLANT A, B_ITEM B, B_PLANT C ", " A.ITEM_CD = B.ITEM_CD AND A.PLANT_CD = C.PLANT_CD AND A.PLANT_CD = " & FilterVar(.txtPlantCd.Value,"","S") & " AND A.ITEM_CD = " & FilterVar(.txtItemCd.Value,"","S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then

			If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(.txtPlantCd.Value,"","S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
				Call DisplayMsgBox("125000","X","X","X")
				.txtPlantNm.Value = ""
				.txtPlantCd.focus 
				Set gActiveElement = document.activeElement
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			.txtPlantNm.Value = lgF0(0)


			If  CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(.txtItemCd.Value,"","S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
   
				Call DisplayMsgBox("122600","X","X","X")
				.txtItemNm.Value = ""
				.txtItemCd.Focus
				Set gActiveElement = document.activeElement
				Exit function
			Else
				lgF0 = Split(lgF0, Chr(11))
				.txtItemNm.Value = lgF0(0)
				Call DisplayMsgBox("122700","X","X","X")
				.txtItemCd.Focus
				Set gActiveElement = document.activeElement
				Exit function
			End If
		End If

 		lgF0 = Split(lgF0, Chr(11))
 		lgF1 = Split(lgF1, Chr(11))
		.txtPlantNm.Value = lgF1(0)
		.txtItemNm.Value = lgF0(0)
	End With
	Plant_Item_Check = True
End Function

Function cboInspClassCd_OnChange()
	If UCase(Trim(frm1.cboInspClassCd.value)) = "P" Then
		Call ggoOper.SetReqAttr(frm1.txtRoutingNo, "N")
		Call ggoOper.SetReqAttr(frm1.txtOprNo, "N")
	Else
		Call ggoOper.SetReqAttr(frm1.txtRoutingNo, "Q")
		Call ggoOper.SetReqAttr(frm1.txtOprNo, "Q")
	End If
End Function



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>검사기준서</font></td>
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
			<TABLE CLASS="BasicTB" CELLSPACING=0 STYLE="HEIGHT: 100%">	
	    		<TR>
					<TD WIDTH=100%>
						<FIELDSET CLASS="CLSFLD" STYLE="HEIGHT: 100%">
							<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0 STYLE="HEIGHT: 100%">
								<TR>
									<TD CLASS="TD5" NOWRAP>공장</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="공장" TAG="12XXXU" ><IMG ALIGN=top HEIGHT=20 NAME=btnPlantCd ONCLICK=vbscript:OpenPlant() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtPlantNm" TAG="14X"></TD>
                                </TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>검사분류</TD>
									<TD CLASS="TD6" NOWRAP><SELECT NAME="cboInspClassCd" ALT="검사분류" STYLE="WIDTH: 150px" tag="13"></SELECT></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 ALT="품목" tag="13XXXU"><IMG src="../../../CShared/image/btnPopup.gif" name=btnItemCd align=top  TYPE="BUTTON" width=16 height=20 onclick="vbscript:OpenItem()">
												<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=30 MAXLENGTH=20 tag="14" ></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>라우팅</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRoutingNo" SIZE=15 MAXLENGTH=7 tag="13XXXU" ALT="라우팅"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoutingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenRoutingNo()">&nbsp;<INPUT TYPE=TEXT NAME="txtRoutingNm" SIZE=20 MAXLENGTH=40 tag="14" ALT="라우팅명"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>공정</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtOprNo" SIZE=10 MAXLENGTH=3 tag="13XXXU" ALT="공정"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOprNo" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenOprNo()"></TD></TD>	
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<BUTTON NAME="btnPreview" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>               
					</TD>                  
				</TR>
			</TABLE>
		</TD>                
	</TR>               
	<TR>               
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " tabindex=-1 WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>               
		</TD>               
	</TR>               
</TABLE>               
</FORM>               
<DIV ID="MousePT" NAME="MousePT">               
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>               
</DIV>               
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST"> 
    <input type="hidden" name="uname" tabindex=-1>
    <input type="hidden" name="dbname" tabindex=-1>
    <input type="hidden" name="filename" tabindex=-1>
    <input type="hidden" name="condvar" tabindex=-1>
	<input type="hidden" name="date" tabindex=-1>
</FORM>                
</BODY>               
</HTML>

