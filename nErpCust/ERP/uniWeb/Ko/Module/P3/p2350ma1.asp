<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2350ma1
'*  4. Program Name         : MRP예시전개 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002-04-16
'*  9. Modifier (First)     : Im Hyun Soo
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
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'==========================================================================================================
Dim EndDate

EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

Const BIZ_PGM_ID		= "p2350mb1.asp"
Const BIZ_PGM_RESULT_ID = "p2351ma1"
Const CookieSplit		= 4877						

'========================================================================================================= 
<!-- #Include file="../../inc/lgVariables.inc" -->
				

Dim IsOpenPop         
Dim lgInvCloseDt

'========================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    
    IsOpenPop = False

End Sub

'==========================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'==========================================================================================================
Sub LoadInfTB19029() 
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "BA") %>
End Sub 

'==========================================================================================================
'	Name : SetDefaultVal()
'	Description : 화면 초기화(수량 Field나 그 외 화면이 뜰 때 Default값을 정해줘야 하는 Field들 Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtFixExecFromDt.text = EndDate
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage()

	WriteCookie "txtPlantCd" , frm1.txtPlantCd.Value
	WriteCookie "txtPlantNm" , frm1.txtPlantNm.Value
	
	PgmJump(BIZ_PGM_RESULT_ID)
	
End Function

'------------------------------------------  OpenCondPlant()  -------------------------------------------------
'	Name : OpenCondPlant()
'	Description : Condition Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenConPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"						' 팝업 명칭 
	arrParam(1) = "B_PLANT"								' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPlantCd.Value)			' Code Condition
	arrParam(3) = ""									' Name Cindition
	arrParam(4) = ""									' Where Condition
	arrParam(5) = "공장"							' TextBox 명칭 
	
   	arrField(0) = "PLANT_CD"							' Field명(0)
    arrField(1) = "PLANT_NM"							' Field명(1)
	arrField(2) = "CONVERT(VARCHAR(4),PLAN_HRZN)"
	arrField(3) = "CONVERT(VARCHAR(4),PTF_FOR_MRP)"
    
    arrHeader(0) = "공장"							' Header명(0)
    arrHeader(1) = "공장명"							' Header명(1)
    arrHeader(2) = "계획기간"						' Header명(1)
    arrHeader(3) = "MRP확정기간"					' Header명(1)
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPlant(arrRet)
	End If	
	
End Function


'------------------------------------------  SetConPlant()  --------------------------------------------------
'	Name : SetConPlant()
'	Description : Condition Plant Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetConPlant(ByRef arrRet)
	frm1.txtPlantCd.Value    = arrRet(0)		
	frm1.txtPlantNm.Value    = arrRet(1)	
   
    frm1.txtFixExecToDt.text = UNIDateAdd("d", arrRet(3), EndDate, parent.gDateFormat)
    frm1.txtPlanExecToDt.text = UNIDateAdd("d", arrRet(2), EndDate, parent.gDateFormat)
    
    frm1.txtPlantCd.focus
    Set gActiveElement = document.activeElement    
End Function

Sub txtPlantCd_OnChange()
    If frm1.txtPlantCd.value = "" Then Exit Sub
    Call LookUpPlant
End Sub

'=======================================================================================================
'   Event Name : txtFixExecFromDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFixExecFromDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFixExecFromDt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtFixExecFromDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtFixExecFromDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtFixExecFromDt_Change() 
	lgBlnFlgChgValue = True 
End Sub

'=======================================================================================================
'   Event Name : txtFixExecFromDt_OnBlur()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtFixExecFromDt_OnBlur()
	Dim DtInvCloseDt
	Dim DtExecFromDt

	If frm1.txtFixExecFromDt.text = "" Then Exit Sub
	
	DtInvCloseDt = UniConvDateAToB(lgInvCloseDt, parent.gDateFormat, parent.gServerDateFormat)
	DtExecFromDt = UniConvDateAToB(frm1.txtFixExecFromDt.Text, parent.gDateFormat, parent.gServerDateFormat)
	
	If DtExecFromDt <= DtInvCloseDt Then
		Call DisplayMsgBox("189250", "x", "x", "x")
		frm1.txtFixExecFromDt.text = UNIDateAdd ("D", 1, lgInvCloseDt, parent.gDateFormat)
		frm1.txtFixExecFromDt.focus
		Set gActiveElement = document.activeElement
		Exit Sub
	End If
	
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFixExecToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtFixExecToDt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtFixExecToDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtFixExecToDt_Change() 
	lgBlnFlgChgValue = True 
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtPlanExecToDt_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtPlanExecToDt.Action = 7
		Call SetFocusToDocument("M")
        frm1.txtPlanExecToDt.Focus
	End If 
End Sub

'=======================================================================================================
'   Event Name : txtValidToDt_Change()
'   Event Desc : change flag setting
'=======================================================================================================
Sub txtPlanExecToDt_Change() 
	lgBlnFlgChgValue = True 
End Sub

'++++++++++++++++++++++++++++++++++++++++++  2.5 개발자 정의 함수  +++++++++++++++++++++++++++++++++++++++
'    개별적 프로그램 마다 필요한 개발자 정의 Procedure (Sub, Function, Validation & Calulation 관련 함수)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
    
	Select Case pRow
		Case "P"
			lgKeyStream = frm1.txtPlantCd.Value & parent.gColSep
	End Select
End Sub        

'++++++++++++++++++++++++++++++++++++++++++  2.5.2 ExecuteMRP  +++++++++++++++++++++++++++++++++++++++
'        Name : ExecuteMRP()    
'        Description : MRP 전개 Main Function          
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 
Function ExecuteMRP()
    
    Dim IntRetCD
	Dim strVal
	    
    If Not chkField(Document, "2") Then Exit Function
    
    If ValidDateCheck(frm1.txtFixExecFromDt, frm1.txtFixExecToDt) = False Then
		frm1.txtFixExecToDt.focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
		
	If ValidDateCheck(frm1.txtFixExecToDt, frm1.txtPlanExecToDt) = False Then
		frm1.txtPlanExecToDt.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End If   
	
	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO, "X", "X")
	
	If IntRetCD = vbNo Then	Exit Function
	
    Call LayerShowHide(1)
    
    With frm1
		.txtMode.value = parent.UID_M0002
		.txtFlgMode.value = lgIntFlgMode

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
    End With	
           
End Function

Sub LookUpPlant()
	Dim strVal
	
	If gLookUpEnable = False Then Exit Sub
    Err.Clear
	
    Call LayerShowHide(1)
    Call MakeKeyStream("P")
    
    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="    & parent.UID_M0001						         
        strVal = strVal     & "&txtKeyStream=" & lgKeyStream
    End With

    Call RunMyBizASP(MyBizASP, strVal)

End Sub

Sub LookUpPlantOk()

End Sub

Sub ExecuteOk()

End Sub

'=========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 

Sub Form_Load()
    Call SetDefaultVal
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
    
    Call SetToolbar("10000000000011")

    Call SetDefaultVal
    Call InitVariables
    If parent.gPlant <> "" Then
		frm1.txtPlantCd.value = UCase(parent.gPlant)
		frm1.txtPlantNm.value = parent.gPlantNm	
		gLookUpEnable = True
		Call LookUpPlant
		frm1.txtFixExecFromDt.focus 
	ELSE
		frm1.txtPlantCd.focus 
	End If   
	
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MRP예시전개</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS=TD5 NOWRAP>공장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="23XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPlant()" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>기준일자</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2350ma1_fpDateTime3_txtFixExecFromDt.js'></script>
							</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>확정전개기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2350ma1_fpDateTime4_txtFixExecToDt.js'></script>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>예시전개기간</TD>
								<TD CLASS=TD6 NOWRAP>
									<script language =javascript src='./js/p2350ma1_fpDateTime4_txtPlanExecToDt.js'></script>
								</TD>								
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>가용재고 감안여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAvailInvFlg ID=rdoAvailInvFlg1 VALUE="Y" CHECKED><LABEL FOR=rdoAvailInvFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoAvailInvFlg ID=rdoAvailInvFlg2 VALUE="N"><LABEL FOR=rdoAvailInvFlg2>감안안함</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>안전재고 감안여부</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSafeInvFlg ID=rdoSafeInvFlg1 VALUE="Y" CHECKED><LABEL FOR=rdoSafeInvFlg1>감안함</LABEL>
													 <INPUT TYPE="RADIO" CLASS="Radio" NAME=rdoSafeInvFlg ID=rdoSafeInvFlg2 VALUE="N"><LABEL FOR=rdoSafeInvFlg2>감안안함</LABEL></TD>
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
					<TD><BUTTON NAME="btnExec" CLASS="CLSMBTN" Flag=1 onclick="ExecuteMRP()">MRP 예시전개</BUTTON></TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:CookiePage">MRP예시전개전환</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
