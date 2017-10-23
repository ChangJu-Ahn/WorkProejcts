<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 품목별 수주집계 출력 
'*  3. Program ID           : S3113OA2
'*  4. Program Name         : 품목별 수주집계 출력 
'*  5. Program Desc         : 품목별 수주집계 출력 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/07/01
'*  8. Modified date(Last)  : 2003/07/01
'*  9. Modifier (First)     : Cho in kuk
'* 10. Modifier (Last)      : 
'* 11. Comment              : 
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

' Popup Index
Const C_PopSalesGrp		= 1
Const C_PopSoType		= 2
Const C_PopItemGrp		= 3
Const C_PopSalesType	= 4
Const C_PopBizParter	= 5

Dim EndDate

' 시스템 날짜 
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size

Dim IsOpenPop          

'=========================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
End Sub

'=========================================
Sub SetDefaultVal()
    frm1.txtConSoDtFromDt.focus 
	frm1.txtConSoDtFromDt.Text	= UNIGetFirstDay(EndDate, Parent.gDateFormat)
	frm1.txtConSoDtToDt.Text		= EndDate
End Sub

'=========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("Q","S","NOCOOKIE", "OA") %>	
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>
End Sub

'=========================================
Function OpenConPopUp(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
		Select Case pvIntWhere
			'영업그룹 
			Case C_PopSalesGrp	
				iArrParam(1) = "B_SALES_GRP"								
				iArrParam(2) = Trim(.txtConSalesGrp.value)			
				iArrParam(3) = ""											
				iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "							
				iArrParam(5) = .txtConSalesGrp.alt					
				
				iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"		
				iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"		
    
			    iArrHeader(0) = .txtConSalesGrp.alt					
			    iArrHeader(1) = .txtConSalesGrpNm.alt					

				.txtConSalesGrp.focus
			
			'수주형태	
			Case C_PopSoType						
				iArrParam(1) = "S_SO_TYPE_CONFIG"				
				iArrParam(2) = Trim(.txtConSoType.value)		
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  and SO_MGMT_FLAG <> " & FilterVar("N", "''", "S") & "  and STO_FLAG = " & FilterVar("N", "''", "S") & " "				
				iArrParam(5) = "수주형태"		
					
				iArrField(0) = "SO_TYPE"			
				iArrField(1) = "SO_TYPE_NM"		
				iArrField(2) = "EXPORT_FLAG"		
				iArrField(3) = "RET_ITEM_FLAG"	
				iArrField(4) = "AUTO_DN_FLAG"
				iArrField(5) = "CI_FLAG"	
							    
				iArrHeader(0) = "수주형태"					
				iArrHeader(1) = "수주형태명"					
				iArrHeader(2) = "수출여부"					
				iArrHeader(3) = "반품여부"					
				iArrHeader(4) = "자동출하생성여부"		
				iArrHeader(5) = "통관여부"			
					    
				frm1.txtConSoType.focus 
					
			'품목그룹	
			Case C_PopItemGrp
				iArrParam(1) = "B_ITEM_GROUP"								
				iArrParam(2) = Trim(.txtConItemGrp.value)				
				iArrParam(3) = ""											
				iArrParam(4) = "LEAF_FLG = " & FilterVar("Y", "''", "S") & " "								
				iArrParam(5) = .txtConItemGrp.alt						
				
				iArrField(0) = "ED15" & Parent.gColSep & "ITEM_GROUP_CD"	
				iArrField(1) = "ED30" & Parent.gColSep & "ITEM_GROUP_NM"	
    
			    iArrHeader(0) = .txtConItemGrp.alt					
			    iArrHeader(1) = .txtConItemGrpNm.alt					

				.txtConItemGrp.focus	

			'판매유형 
			Case C_PopSalesType
				iArrParam(1) = "B_MINOR"									
				iArrParam(2) = Trim(.txtConSalesType.value)			
				iArrParam(3) = ""											
				iArrParam(4) = "MAJOR_CD = " & FilterVar("S0001", "''", "S") & ""							
				iArrParam(5) = .txtConSalesType.alt					
				
				iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"			
				iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"			
    
			    iArrHeader(0) = .txtConSalesType.alt					
			    iArrHeader(1) = .txtConSalesTypeNm.alt					

				.txtConSalesType.focus
			
			'거래처 
			Case C_PopBizParter											
				iArrParam(1) = "B_BIZ_PARTNER"								
				iArrParam(2) = Trim(.txtConBizParter.value)			
				iArrParam(3) = ""											
				iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND BP_TYPE IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"	
				iArrParam(5) = .txtConBizParter.alt					
				
				iArrField(0) = "ED15" & Parent.gColSep & "BP_CD"			
				iArrField(1) = "ED30" & Parent.gColSep & "BP_NM"			
    
			    iArrHeader(0) = .txtConBizParter.alt				
			    iArrHeader(1) = .txtConBizParterNm.alt				

				.txtConBizParter.focus
				
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
			Case C_PopSalesGrp
				.txtConSalesGrp.Value	= pvArrRet(0)
				.txtConSalesGrpNm.Value	= pvArrRet(1)
				
			Case C_PopSoType	
				.txtConSoType.Value		= pvArrRet(0)
				.txtConSoTypeNm.Value	= pvArrRet(1)
				
			Case C_PopItemGrp
				.txtConItemGrp.Value	= pvArrRet(0)
				.txtConItemGrpNm.Value	= pvArrRet(1)
				
			Case C_PopSalesType	
				.txtConSalesType.Value		= pvArrRet(0)
				.txtConSalesTypeNm.Value	= pvArrRet(1)
				
			Case C_PopBizParter	
				.txtConBizParter.Value		= pvArrRet(0)
				.txtConBizParterNm.Value	= pvArrRet(1)
			
		End Select	
	End With

End Function

'=========================================
Sub Form_Load()
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														'⊙: Initializes local global variables
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 

End Sub


'==========================================
Sub txtConSoDtFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtConSoDtFromDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtConSoDtFromDt.Focus
    End If
End Sub

'==========================================
Sub txtConSoDtToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtConSoDtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtConSoDtToDt.Focus
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
    If Not chkField(Document, "1") Then	Exit Function

	If ValidDateCheck(frm1.txtConSoDtFromDt, frm1.txtConSoDtToDt) = False Then Exit Function
    
	Dim iStrUrl, iStrParam1, iStrParam2, iStrParam3, iStrParam4, iStrParam5, iStrParam6, iStrParam7
	
	iStrParam1 = UniConvDateToYYYYMMDD(frm1.txtConSoDtFromDt.Text,Parent.gDateFormat,Parent.gServerDateType)
	iStrParam2 = UniConvDateToYYYYMMDD(frm1.txtConSoDtToDt.Text,Parent.gDateFormat,Parent.gServerDateType)
		
	'영업그룹 
	If UCase(frm1.txtConSalesGrp.value) = "" Then
		iStrParam3 = "%"
	Else
		iStrParam3 = Replace(Trim(UCase(frm1.txtConSalesGrp.value)), "'" ,  "''")
	End If
	
	'수주형태 
	If UCase(frm1.txtConSoType.value) = "" Then
		iStrParam4 = "%"
	Else
		iStrParam4 = Replace(Trim(UCase(frm1.txtConSoType.value)), "'" ,  "''")
	End If
	
	'품목그룹 
	If UCase(frm1.txtConItemGrp.value) = "" Then
		iStrParam5 = "%"
	Else
		iStrParam5 = Replace(Trim(UCase(frm1.txtConItemGrp.value)), "'" ,  "''")
	End If
	
	'판매유형 
	If UCase(frm1.txtConSalesType.value) = "" Then
		iStrParam6 = "%"
	Else
		iStrParam6 = Replace(Trim(UCase(frm1.txtConSalesType.value)), "'" ,  "''")
	End If
	
	'거래처 
	If UCase(frm1.txtConBizParter.value) = "" Then
		iStrParam7 = "%"
	Else
		iStrParam7 = Replace(Trim(UCase(frm1.txtConBizParter.value)), "'" ,  "''")
	End If       	
      
	' 출력조건을 지정하는 부분 수정 
	iStrUrl = "txtConSoDtFromDt|" & iStrParam1 & "|txtConSoDtToDt|" & iStrParam2 & _
			  "|txtConSalesGrp|" & iStrParam3 & "|txtConSoType|" & iStrParam4 & _
			  "|txtConItemGrp|" & iStrParam5 & "|txtConSalesType|" & iStrParam6 & _
			  "|txtConBizParter|" & iStrParam7 
	
	' Print 함수에서 호출 
	ObjName = AskEBDocumentName("s3113oa2","ebr")

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목별 수주집계출력</font></td>
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
									<TD CLASS="TD5" NOWRAP>수주일</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s3113oa2_fpDateTime1_txtConSoDtFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s3113oa2_fpDateTime2_txtConSoDtToDt.js'></script>
												</TD>
											
										    </TR>
										</TABLE>
									</TD>							
								<TR>
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSalesGrp" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesGrp) ">
															<INPUT TYPE=TEXT NAME="txtConSalesGrpNm" SIZE=20 tag="14" ALT="영업그룹명"></TD>								
								<TR>
									<TD CLASS="TD5" NOWRAP>수주형태</TD>
									<TD CLASS="TD6" NOWRAP> <INPUT NAME="txtConSoType" ALT="수주형태" TYPE="Text" MAXLENGTH="18" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSoType" align=top TYPE="BUTTON"ONCLICK="vbscript:Call OpenConPopUp(C_PopSoType)">
														    <INPUT NAME="txtConSoTypeNm" ALT="수주형태명" TYPE="Text" SIZE=20 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>품목그룹</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConItemGrp" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConItemGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopItemGrp)">
															<INPUT TYPE=TEXT NAME="txtConItemGrpNm" SIZE=20 tag="14" ALT="품목그룹명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>판매유형</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSalesType" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="판매유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesType)">
															<INPUT TYPE=TEXT NAME="txtConSalesTypeNm" SIZE=20 tag="14" ALT="판매유형명"></TD>
								</TR>	
								<TR>
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBizParter" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="거래처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBizParter" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBizParter)">
															<INPUT TYPE=TEXT NAME="txtConBizParterNm" SIZE=20 tag="14" ALT="거래처명"></TD>
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
