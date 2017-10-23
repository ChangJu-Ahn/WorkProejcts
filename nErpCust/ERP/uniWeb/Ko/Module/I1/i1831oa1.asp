<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : I1831OA1.asp
'*  4. Program Name         : 재고관리일보출력(품목별)
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/30
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   ******************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->			
<!--'==========================================  1.1.1 Style Sheet  =======================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">	

<!--'==========================================  1.1.2 공통 Include   =======================================-->
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"> </SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs"></SCRIPT>

<SCRIPT LANGUAGE = VBSCRIPT>
Option Explicit            

'==========================================  1.2.1 Global 상수 선언  ======================================
<%
StartDate   = GetSvrDate
%>

'==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim hPosSts
Dim IsOpenPop  

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : 변수 초기화(Global 변수, 초기화가 필요한 변수 또는 Flag들을 Setting한다.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE 
    lgBlnFlgChgValue = False         
    lgIntGrpCount = 0                
    
    IsOpenPop = False		
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("Q", "I", "NOCOOKIE", "OA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
    frm1.txtDt.Text = UniConvDateAToB("<%=StartDate%>",parent.gServerDateFormat,parent.gDateFormat)
End Sub

'--------------------------------------------------------------------------------------------------------- 
'	Name : OpenPlant()
'	Description : Plant Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPlantCd()

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
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		Call SetPlantCd(arrRet)
	End If	
	
End Function

'------------------------------------------  OpenSL()  --------------------------------------------------
'	Name : OpenSL()
'	Description : SL Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	
	If Trim(frm1.txtPlantCd.value)  = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End If

	'-----------------------
	'Check Plant CODE	
	'-----------------------
	'If	CommonQueryRs("	PLANT_NM "," B_PLANT ",	" PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value,"","S"), _
	'	lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	= False	Then
		
	'	Call DisplayMsgBox("125000","X","X","X")
	'	frm1.txtPlantNm.Value = ""
	'	frm1.txtPlantCd.focus
	'	Exit function
	'End	If
	'lgF0 = Split(lgF0, Chr(11))
	'frm1.txtPlantNm.Value = lgF0(0)

	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = ""
	if frm1.txtPlantCd.value <> "" then
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.value, "''", "S")	
	else
	arrParam(4) = ""
	end if
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSLCd.focus
		Exit Function
	Else
		Call SetSL(arrRet)
	End If	
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
'	Name : OpenItemCd()
'	Description : Item1 PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenItemCd(ByVal strCode, ByVal iPos)
	Dim iCalledAspName
	Dim IntRetCD
	Dim arrRet
	Dim arrParam(5), arrField(6)
	
	If IsOpenPop = True Then Exit Function
		
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If	
	
	'-----------------------
	'Check Plant CODE	
	'-----------------------
	'If	CommonQueryRs("	PLANT_NM "," B_PLANT ",	" PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value,"","S"), _
	'	lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	= False	Then
		
	'	Call DisplayMsgBox("125000","X","X","X")
	'	frm1.txtPlantNm.Value = ""
	'	frm1.txtPlantCd.focus
	'	Exit function
	'End	If
	'lgF0 = Split(lgF0, Chr(11))
	'frm1.txtPlantNm.Value = lgF0(0)

	iCalledAspName = AskPRAspName("b1b11pa3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", Parent.VB_INFORMATION, "b1b11pa3", "X")
		IsOpenPop = False
		Exit Function
    End If
    
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.Value)	
	arrParam(1) = strCode	
	arrParam(2) = ""		
	arrParam(3) = ""		
	
    arrField(0) = 1 		
    arrField(1) = 2 		
    arrField(2) = 3			
    
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent,arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtItemCd1.focus
		Exit Function
	Else
		Call SetItemCd(arrRet, iPos)
	End If	
End Function

'------------------------------------------  OpenTrackingNo()  --------------------------------------------------
' Name : OpenTrackingNo()
' Description : Item Account Popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)= UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "Tracking No."	
	arrParam(1) = "S_SO_TRACKING"				
	arrParam(2) = Trim(frm1.txtTrackingNo.value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "Tracking No."			
	
    arrField(0) = "Tracking_No"	
    arrField(1) = "Item_Cd"	
    
    arrHeader(0) = "Tracking_No"		
    arrHeader(1) = "품목"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value = arrRet(0)
		frm1.txtTrackingNo.focus
	End If	
End Function

'------------------------------------------  SetSL()  --------------------------------------------------
'	Name : SetSL()
'	Description : SL Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSL(byval arrRet)
	frm1.txtSLCd.Value    = arrRet(0)		
	frm1.txtSLNm.Value    = arrRet(1)
	frm1.txtSLCd.focus
	lgBlnFlgChgValue	  = True
End Function

'------------------------------------------  SetPlantCd()  --------------------------------------------------
'	Name : SetPlantCd()
'	Description : Plant  Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPlantCd(ByVal arrRet)
	frm1.txtPlantCd.value = arrRet(0)
	frm1.txtPlantNm.value = arrRet(1)
	frm1.txtPlantCd.focus
	lgBlnFlgChgValue	  = True  
End Function

'------------------------------------------  SetItemCd()  --------------------------------------------------
'	Name : SetItemCd()
'	Description : SetItemCd Popup에서 return된 값 
'---------------------------------------------------------------------------------------------------------
Function SetItemCd(ByVal arrRet, ByVal iPos)	
	If iPos = 0 Then
		frm1.txtItemCd1.value = arrRet(0) 
		frm1.txtItemNm1.value = arrRet(1)
		frm1.txtItemCd1.focus
		lgBlnFlgChgValue	  = True
	ElseIf iPos = 1 Then
		frm1.txtItemCd2.value = arrRet(0) 
		frm1.txtItemNm2.value = arrRet(1)
		frm1.txtItemCd2.focus
		lgBlnFlgChgValue	  = True
	End If
End Function


'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()

	Call InitVariables							
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, parent.gDateFormat, parent.gComNum1000, parent.gComNumDec)		
	Call ggoOper.LockField(Document, "N")						
	Call ggoOper.FormatDate(frm1.txtDt, parent.gDateFormat, 1)

	Call SetToolbar("10000000000011")
	Call SetDefaultVal
	If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm
		frm1.txtSlCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDt_KeyPress()
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call MainQuery()
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Call BtnPreview()
    FncQuery = True    
End Function

'========================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'========================================================================================
Function FncBtnPrint()
	Dim var1, var2, var3, var4, var5, var6
    Dim condvar 
	Dim ObjName
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then     
       Exit Function
    End If
    
    If Plant_Or_SLCd_Check = False Then 
		Exit Function
	End If
	
    var1 = UCase(Trim(frm1.txtPlantCd.value))
	var2 = "%" & UCase((Trim(frm1.txtSlCd.value))) & "%"
	
	var3 = UNIConvDateToYYYYMMDD(frm1.txtDt.text, parent.gDateFormat,"")
	
	If frm1.txtItemCd1.value = "" then 
		var4 = "0"	
	Else
		var4 = UCase((Trim(frm1.txtItemCd1.value)))
	End If
	
	If frm1.txtItemCd2.value = "" then 
		var5 = "zzzzzzzzzzzzzzzzzz"	
	Else
		var5 = UCase((Trim(frm1.txtItemCd2.value)))
	End If
	var6 = "%" & Trim(frm1.txtTrackingNo.value) & "%"
	
	condvar = condvar & "PLANTCD|"          & var1
    condvar = condvar & "|SLCD|"            & var2
    condvar = condvar & "|DocumentDt|"      & var3
    condvar = condvar & "|ItemCdFr|"        & var4
    condvar = condvar & "|ItemCdTo|"        & var5
    condvar = condvar & "|TrackingNo|"      & var6

	ObjName = AskEBDocumentName("i1831oa1", "ebr")
	Call FncEBRprint(EBAction, ObjName, condvar) 	    
	
End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview()
	Dim var1, var2, var3, var4, var5, var6
    Dim condvar 
	Dim ObjName
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "1") Then   
       Exit Function
    End If
    
    If Plant_Or_SLCd_Check = False Then 
		Exit Function
	End If
	
    var1 = UCase(Trim(frm1.txtPlantCd.value))
    var2 = "%" & UCase((Trim(frm1.txtSlCd.value))) & "%"
	
	var3 = UNIConvDateToYYYYMMDD(frm1.txtDt.text, parent.gDateFormat,"")
	
	If frm1.txtItemCd1.value = "" then 
		var4 = "0"	
	Else
		var4 = UCase((Trim(frm1.txtItemCd1.value)))
	End If
	
	If frm1.txtItemCd2.value = "" then 
		var5 = "zzzzzzzzzzzzzzzzzz"	
	Else
		var5 = UCase((Trim(frm1.txtItemCd2.value)))
	End If
	var6 = "%" & Trim(frm1.txtTrackingNo.value) & "%"
	
	condvar = condvar & "PLANTCD|"          & var1
    condvar = condvar & "|SLCD|"            & var2
    condvar = condvar & "|DocumentDt|"      & var3
    condvar = condvar & "|ItemCdFr|"        & var4
    condvar = condvar & "|ItemCdTo|"        & var5
    condvar = condvar & "|TrackingNo|"      & var6

	ObjName = AskEBDocumentName("i1831oa1", "ebr")
	Call FncEBRPreview(ObjName, condvar)    
	
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
    Call parent.FncExport(parent.C_SINGLE)                                      
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 화면 속성, Tab유무 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , True)   
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================
' Function Name : Plant_Or_SLCd_Check
' Function Desc : 
'========================================================================================
Function Plant_Or_SLCd_Check()
	'-----------------------
	'Check Plant CODE	
	'-----------------------
	If	CommonQueryRs("	PLANT_NM "," B_PLANT ",	" PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	= False	Then
		
		Call DisplayMsgBox("125000","X","X","X")
		frm1.txtPlantNm.Value = ""
		frm1.txtPlantCd.focus
		Plant_Or_SLCd_Check = False
		Exit function
	End	If
	lgF0 = Split(lgF0, Chr(11))
	frm1.txtPlantNm.Value = lgF0(0)

	If Trim(frm1.txtSLCd.Value)	<> "" Then
	'-----------------------
	'Check SLCd	CODE	
	'-----------------------
		If	CommonQueryRs("	SL_NM "," B_STORAGE_LOCATION ",	" SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
		lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	= False	Then
		
			Call DisplayMsgBox("125700","X","X","X")
			frm1.txtSLNm.Value = ""
			frm1.txtSLCd.focus
			Plant_Or_SLCd_Check = False
			Exit function
		End	If
		lgF0 = Split(lgF0, Chr(11))
		frm1.txtSLNm.Value = lgF0(0)

		If	CommonQueryRs("	SL_NM "," B_STORAGE_LOCATION ",	" PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)	= False	Then
			
			Call DisplayMsgBox("169922","X","X","X")
			frm1.txtSLCd.focus
			Plant_Or_SLCd_Check = False
			Exit function
		End	If
	End	If
	
	'-----------------------
	'Check Item CODE	
	'-----------------------
    If frm1.txtItemCd1.Value <> "" Then
		If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd1.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItemNm1.Value = lgF0(0)
		Else
			frm1.txtItemNm1.Value = ""
		End If
	End If
    
    If frm1.txtItemCd2.Value <> "" Then
		If 	CommonQueryRs(" ITEM_NM "," B_ITEM ", " ITEM_CD = " & FilterVar(frm1.txtItemCd2.Value, "''", "S"), _
			lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
			
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtItemNm2.Value = lgF0(0)
		Else
			frm1.txtItemNm2.Value = ""
		End If
	End If

	Plant_Or_SLCd_Check	= True
End	Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%> WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>재고일보출력</font></td>
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
					<TD>
						<TABLE <%=LR_SPACE_TYPE_60%> >		
							<TR>	
							    <TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 MAXLENGTH=40 tag="14" ALT="공장명">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>일자</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/i1831oa1_OBJECT1_txtDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>창고</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE="TEXT" NAME="txtSLCd" SIZE=8 MAXLENGTH=7 tag="11XXXU" ALT="창고" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSLCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE="TEXT" NAME="txtSLNm" SIZE=30 MAXLENGTH=40 tag="14" ALT="창고명">
								</TD>
							</TR>					
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd1" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd frm1.txtItemCd1.value, 0">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=30 MAXLENGTH=40 tag="14" ALT="품목명">&nbsp;~&nbsp;
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP>
								<INPUT TYPE=TEXT NAME="txtItemCd2" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd2" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenItemCd frm1.txtItemCd2.value, 1">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm2" SIZE=30 MAXLENGTH=40 tag="14" ALT="품목명">
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5">Tracking No.</TD>      
								<TD CLASS="TD6">
								<INPUT TYPE=TEXT SIZE=20 NAME="txtTrackingNo" MAXLENGTH="25"  tag="11XXXU" ALT = "Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo"  align="top" TYPE="BUTTON" ONCLICK="vbscript:OpenTrackingNo()">
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
				  <TD WIDTH = 10>&nbsp;</TD>
		          <TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>		
	            </TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
<input type="hidden" name="uname" TABINDEX="-1">
	<input type="hidden" name="dbname" TABINDEX="-1">
	<input type="hidden" name="filename" TABINDEX="-1">
	<input type="hidden" name="condvar" TABINDEX="-1">
	<input type="hidden" name="date" TABINDEX="-1">
</FORM>
</BODY>
</HTML>
