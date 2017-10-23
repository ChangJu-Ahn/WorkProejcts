<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 매출채권집계 출력 
'*  3. Program ID           : s5111oa1
'*  4. Program Name         : 매출채권집계 출력 
'*  5. Program Desc         : 매출채권집계 출력 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/06/18
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Cho Sung Hyun
'* 10. Modifier (Last)      : Hwang Seongbae 
'* 11. Comment              : 표준적용 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

' Popup Index
Const C_PopItemCd = 1
Const C_PopSalesOrg	= 2

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size

Dim IsOpenPop          

'=========================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'=========================================
Sub SetDefaultVal()
    frm1.txtItem_Cd.focus 
	frm1.txtBillFromDt.Text = UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.txtBillToDt.Text = EndDate
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("Q","S","NOCOOKIE", "OA") %>	
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'=========================================
Function OpenConPop(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
		Select Case pvIntWhere
			Case C_PopItemCd
				iArrParam(1) = "B_ITEM"
				iArrParam(2) = Trim(.txtItem_cd.value)
				iArrParam(3) = ""
				iArrParam(4) = ""
				iArrParam(5) = .txtItem_cd.alt
					
				iArrField(0) = "ED15" & Parent.gColSep & "ITEM_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "ITEM_NM"
    
				iArrHeader(0) = .txtItem_cd.Alt
				iArrHeader(1) = .txtItem_Cd_Nm.Alt

				.txtItem_cd.focus	

			Case 2												
				iArrParam(1) = "B_SALES_ORG"
				iArrParam(2) = Trim(.txtSales_Org.value)
				iArrParam(3) = ""
				iArrParam(4) = ""
				iArrParam(5) = .txtSales_Org.Alt
						
			    iArrField(0) = "SALES_ORG"
			    iArrField(1) = "SALES_ORG_NM"
					    
			    iArrHeader(0) = .txtSales_Org.Alt
			    iArrHeader(1) = .txtSales_Org_Nm.Alt
			    
			    .txtSales_Org.focus
		End Select
	End With
	
	iArrParam(0) = iArrParam(5)
	
	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) <> "" Then
		Call SetConPop(iArrRet,pvIntWhere)
	End If

End Function

'=========================================
Function SetConPop(Byval pvArrRet,Byval pvIntWhere)
	With frm1	
		Select Case pvIntWhere
			Case C_PopItemCd
				.txtItem_Cd.Value		= pvArrRet(0)
				.txtItem_Cd_Nm.Value	= pvArrRet(1)
				
			Case C_PopSalesOrg	
				.txtSales_Org.Value		= pvArrRet(0)
				.txtSales_Org_Nm.Value	= pvArrRet(1)
			
		End Select	
	End With

End Function

'=========================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables
    Call GetValue_ko441()														'⊙: Initializes local global variables
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 

End Sub

'==========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
   
End Sub

'==========================================
Sub txtBillFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBillFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillFromDt.Focus
    End If
End Sub

'==========================================
Sub txtBillToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtBillToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBillToDt.Focus
    End If
End Sub

'========================================
Function BtnPreview_OnClick()
	Call BtnPrint("N")
End Function

'========================================
Function BtnPrint_OnClick()
	Call BtnPrint("Y")
End Function

'========================================
Function BtnPrint(ByVal pvStrPrint)

Dim vargBizArea,vargPlant,vargSalesGrp,vargSalesOrg
 
    If Not chkField(Document, "1") Then	Exit Function

	If ValidDateCheck(frm1.txtBillFromDt, frm1.txtBillToDt) = False Then Exit Function
    
	Dim iStrUrl, iStrParam1, iStrParam2, iStrParam3, iStrParam4, iStrParam5, iStrParam6
	
	If UCase(frm1.txtItem_Cd.value) = "" Then
		iStrParam2 = "%"
	Else
		iStrParam2 = Replace(Trim(UCase(frm1.txtItem_Cd.value)), "'" ,  "''")
	End If
    
    If UCase(frm1.txtSales_Org.value) = "" Then
		iStrParam3 = "%"
	Else
		iStrParam3 = Replace(Trim(UCase(frm1.txtSales_Org.value)), "'",  "''")
	End If

	iStrParam4 = UniConvDateToYYYYMMDD(frm1.txtBillFromDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	iStrParam5 = UniConvDateToYYYYMMDD(frm1.txtBillToDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	
    If frm1.Rb_WK1.checked = True Then
		iStrParam6 = "%"	
	ElseIf frm1.Rb_WK2.checked = True Then
		iStrParam6 = "N"
	ElseIf frm1.Rb_WK3.checked = True Then
		iStrParam6 = "Y"
    End If

	If lgBACd <> "" Then
		vargBizArea = " AND S_BILL_HDR.BIZ_AREA =  " & FilterVar(Trim(UCase(lgBACd)), "" ,  "S")
	Else
		vargBizArea = ""
	End If
	If lgPLCd <> "" Then
		vargPlant = " AND S_BILL_DTL.PLANT_CD = " & FilterVar(Trim(UCase(lgPLCd)), "" ,  "S")
	Else
		vargPlant = ""
	End If
	If lgSGCd <> "" Then
		vargSalesGrp = " AND S_BILL_HDR.SALES_GRP = " & FilterVar(Trim(UCase(lgSGCd)), "" ,  "S")
	Else
		vargSalesGrp = ""
	End If
	If lgSOCd <> "" Then
		vargSalesOrg = " AND S_BILL_HDR.SALES_ORG = " & FilterVar(Trim(UCase(lgSOCd)), "" ,  "S")
	Else
		vargSalesOrg = ""
	End If


   
	'--출력조건을 지정하는 부분 수정 
	iStrUrl = "CUR|" & iStrParam1 & "|ITEM_CD|" & iStrParam2 & "|SALES_ORG|" & iStrParam3 & "|BillFromDt|" & iStrParam4 & "|BillToDt|" & iStrParam5 & "|POST_FLAG|" & iStrParam6 
	istrUrl = istrUrl & "|gBizArea|" & vargBizArea 
	istrUrl = istrUrl & "|gPlant|" & vargPlant 
	istrUrl = istrUrl & "|gSalesGrp|" & vargSalesGrp
	istrUrl = istrUrl & "|gSalesOrg|" & vargSalesOrg 

	' Print 함수에서 호출 
    If frm1.Rb2_WK1.checked = True Then
		ObjName = AskEBDocumentName("s5112og1_KO441","ebr")
			
	ElseIf frm1.Rb2_WK2.checked = True Then
		ObjName = AskEBDocumentName("s5112og2_KO441","ebr")
		
	ElseIf frm1.Rb2_WK3.checked = True Then
		ObjName = AskEBDocumentName("s5112og3_KO441","ebr")
    End If

	If pvStrPrint = "N" Then
		' 미리보기 
		Call FncEBRPreview(ObjName, iStrUrl)
	Else
		' 출력 
		Call FncEBRprint(EBAction, ObjName, iStrUrl)
	End If
		
End Function

'========================================
 Function FncPrint() 
	Call parent.FncPrint()
End Function

'========================================
 Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================
Function FncExit()
	FncExit = True
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권집계</font></td>
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
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItem_Cd" ALT="품목" TYPE="Text" MAXLENGTH="18" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem_Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop C_PopItemCd">&nbsp;<INPUT NAME="txtItem_Cd_Nm" ALT="품목명" TYPE="Text" SIZE=30 tag="14"></TD>
								
								<TR>
									<TD CLASS=TD5 NOWRAP>영업조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSales_Org" ALT="영업조직" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSales_Org" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop C_PopSalesOrg">&nbsp;<INPUT NAME="txtSales_Org_Nm" ALT="영업조직명" TYPE="Text" MAXLENGTH="20" SIZE=30 tag="14"></TD>
								</TR>
								
								<TR>
									<TD CLASS="TD5" NOWRAP>매출채권일</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s5111oa1_fpDateTime1_txtBillFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s5111oa1_fpDateTime2_txtBillToDt.js'></script>
												</TD>
											
										    </TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>확정여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK1 Checked><LABEL FOR=Rb_WK1>전체</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK2><LABEL FOR=Rb_WK2>확정</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO1 ID=Rb_WK3><LABEL FOR=Rb_WK3>미확정</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>집계형태</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO2 ID=Rb2_WK1 Checked><LABEL FOR=Rb2_WK1>주문처별</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO2 ID=Rb2_WK2><LABEL FOR=Rb2_WK2>품목별</LABEL>&nbsp;
														   <INPUT TYPE="RADIO" CLASS="Radio" NAME=RADIO2 ID=Rb2_WK3><LABEL FOR=Rb2_WK3>영업그룹별</LABEL></TD>
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
					<TD>
					    <BUTTON NAME="BtnPreview" CLASS="CLSSBTN" Flag=1>미리보기</BUTTON>&nbsp;
					    <BUTTON NAME="BtnPrint" CLASS="CLSSBTN" Flag=1>인쇄</BUTTON>
					</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX ="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX ="-1"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname" TABINDEX ="-1">
    <input type="hidden" name="dbname" TABINDEX ="-1">
    <input type="hidden" name="filename" TABINDEX ="-1">
    <input type="hidden" name="condvar" TABINDEX ="-1">
	<input type="hidden" name="date" TABINDEX ="-1">
</FORM>
</BODY>
</HTML>
