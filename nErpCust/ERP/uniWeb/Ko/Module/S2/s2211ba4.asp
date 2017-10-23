<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 판매계획관리 
'*  3. Program ID           : S2211BA4
'*  4. Program Name			: 품목별공장배분비일괄생성 
'*  5. Program Desc         : 품목별공장배분비일괄생성 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/01/25
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park Yong Sik
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/lgvariables.inc"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             <% '☜: indicates that All variables must be declared in advance %>

Const C_PopItemGroupCd	= 1
Const C_PopItemCd		= 2

Const BIZ_PGM_ID = "s2211bb4.asp"											<% '☆: 비지니스 로직 ASP명 %>
Const BIZ_JUMP_ID = "B1B05MA1"				 '☆: JUMP시 비지니스 로직 ASP명 

Dim IsOpenPop

<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                       	              '⊙: Indicates that current mode is Create mode
    
    '----------  Coding part  -------------------------------------------------------------
    IsOpenPop = False							'☆: 사용자 변수 초기화 
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.cboConItemAcct.focus
End Sub

'========================================================================================================= 
Sub InitComboBox()	
		
	'품목계정 
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & " ", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	Call SetCombo2(frm1.cboConItemAcct, lgF0,lgF1, Chr(11))

End Sub

'========================================================================================================= 
Function OpenConPopup(ByVal pvIntWhere)
	Dim iarrRet
	Dim iArrParam(6), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	Select Case pvIntWhere
	Case C_PopItemGroupCd			'품목그룹 
		iArrParam(0) = "품목그룹"
		iArrParam(1) = "B_ITEM_GROUP"	
		iArrParam(2) = frm1.txtConItemGroupCd.value	
		iArrParam(3) = ""
		iArrParam(4) = ""							<%' Where Condition%>
		iArrParam(5) = "품목그룹"
			
		iArrField(0) = "ED15" & Parent.gColSep & "ITEM_GROUP_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "ITEM_GROUP_NM"	
			
		iArrHeader(0) = "품목그룹"
		iArrHeader(1) = "품목그룹명"
			
		frm1.txtConItemGroupCd.focus
			
		iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	Case C_PopItemCd '품목명 
		OpenConPopup = OpenConItemPopup(pvIntWhere)
		frm1.txtConItemCd.focus
		Exit Function
			
	End Select
	
	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet, pvIntWhere)
	End If

End Function

'========================================================================================================= 
Function OpenConItemPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(6)
	Dim iCalledAspName

	OpenConItemPopup = False
	
	iCalledAspName = AskPRAspName("s2210pa1")
	
	If Trim(iCalledAspName) = "" Then
		Call DisplayMsgBox("900040", parent.VB_INFORMATION, "s2210pa1", "X")
		lgBlnOpenPop = False
		Exit Function
	End If

	iArrParam(0) = Trim(frm1.txtConItemCd.value)
	iArrParam(1) = ""
	iArrParam(2) = ""
	iArrParam(3) = ""
	iArrParam(4) = frm1.cboConItemAcct.value
	iArrParam(5) = Trim(frm1.txtConItemGroupCd.value)
	
	iArrRet = window.showModalDialog(iCalledAspName, Array(window.parent,iArrParam), _
	 "dialogWidth=850px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConItemPopup = SetConPopup(iArrRet, pvIntWhere)
	End If	
End Function

'========================================================================================================= 
Function SetConPopup(ByVal iArrRet, ByVal pvIntWhere)

	SetConPopup = False
	
	Select Case pvIntWhere
		Case C_PopItemGroupCd
			frm1.txtConItemGroupCd.Value = iArrRet(0)
			frm1.txtConItemGroupNm.Value = iArrRet(1)
		Case C_PopItemCd
			frm1.txtConItemCd.Value = iArrRet(0)
			frm1.txtConItemNm.Value = iArrRet(1)
	End Select

	SetConPopup = True

End Function

'========================================================================================================= 
Sub Form_Load()
																		'⊙: Load Common DLL
    Call InitVariables																'⊙: Initializes local global variables    
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------           
    Call InitComboBox
    Call SetDefaultVal
    Call CookiePage(0)
    Call SetToolbar("1000000000000111") 

End Sub

'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================================= 
Function CookiePage(Byval pvKubun)

	On Error Resume Next
	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>
	Dim iStrItemCd

	With frm1
		If pvKubun = 1 Then
			WriteCookie CookieSplit , .txtConItemCd.value
		ElseIf pvKubun = 0 Then
			iStrItemCd = Trim(ReadCookie(CookieSplit))
			
			If iStrItemCd = "" then Exit Function
			.txtConItemCd.value = iStrItemCd			
			WriteCookie CookieSplit , ""
		End If
	End With
End Function

'========================================================================================================= 
' Function Desc : Jump시 데이타 변경여부 체크 
'==========================================================================================================
Function JumpChgCheck(byVal pvStrJumpPgmId)
	Call CookiePage(1)
	Call PgmJump(pvStrJumpPgmId)
End Function

'========================================================================================================= 
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================================================================================= 
Function FncExit()
    FncExit = True
End Function

'========================================================================================================= 
Function ExeReflect() 
	Call BtnDisabled(1)
	Dim strVal
	Dim IntRetCD

	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If
	
	If DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X") = vbNo Then
		Call BtnDisabled(0)
		Exit Function
	End If

	With frm1

		strVal = BIZ_PGM_ID & "?txtItemAcct="	& .cboConItemAcct.value
		strVal = strVal  & "&txtItemGroupCd=" & Trim(.txtConItemGroupCd.value)
		strVal = strVal  & "&txtItemCd=" & Trim(.txtConItemCd.value)
		strVal = strVal  & "&txtUserId=" & parent.gUsrId

	End With
	
	If LayerShowHide(1) = False then
		Call BtnDisabled(0)
		Exit Function 
	End if

	Call RunMyBizASP(MyBizASP, strVal)	                                        '☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG
End Function

'========================================================================================================= 
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	Call DisplayMsgBox("990000","X","X","X")
End Function

<%'==========================================================================================
'   Event Desc : 품목그룹 
'========================================================================================== %>
Function txtConItemGroupCd_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtConItemGroupCd.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("Y", "''", "S") & " ", "" & FilterVar("N", "''", "S") & " ", "default", "default", "" & FilterVar("IG", "''", "S") & "", C_PopItemGroupCd) Then
				.txtConItemGroupCd.value = ""
				.txtConItemGroupNm.value = ""
				.txtConItemGroupCd.focus
			End If
			txtConItemGroupCd_OnChange = False
		Else
			.txtConItemGroupNm.value = ""
		End If
	End With
End Function

<%'==========================================================================================
'   Event Desc : 품목 
'==========================================================================================
%>
Function txtConItemCd_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtConItemCd.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("IT", "''", "S") & "", C_PopItemCd) Then
				.txtConItemCd.value = ""
				.txtConItemNm.value = ""
				.txtConItemCd.focus
			End If
			txtConItemCd_OnChange = False
		Else
			.txtConItemNm.value = ""
		End If
	End With
End Function

<%'======================================   GetCodeName()  =====================================
'	Name : GetCodeName()
'	Description : 코드값에 해당하는 명을 Display한다.
'====================================================================================================
%>
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(5), iArrTemp
	
	GetCodeName = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		If err.number = 0 Then
			GetCodeName = OpenConPopup(pvIntWhere)
		Else
			MsgBox err.Description
		End If
	End if
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB4" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목별공장배분비일괄생성</font></td>
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
				<TR HEIGHT=20>
					<TD CLASS="TD5"></TD>
					<TD CLASS="TD6"></TD>
				</TR>
				<TR HEIGHT=35>		
					<TD CLASS="TD5">품목계정</TD>
					<TD CLASS="TD6"><SELECT NAME="cboConItemAcct" tag="1X"><OPTION value=""></OPTION></SELECT></TD>
				</TR>
				<TR HEIGHT=35>		
					<TD CLASS="TD5">품목그룹</TD>
					<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtConItemGroupCd" SIZE=20 MAXLENGTH=10 tag="1XXXXU" ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopItemGroupCd ">
												   <INPUT TYPE=TEXT NAME="txtConItemGroupNm" SIZE=40 MAXLENGTH=40 tag="14">	</TD>
				</TR>
				<TR HEIGHT=35>
					<TD CLASS=TD5 NOWRAP>품목</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtConItemCd" SIZE=20 MAXLENGTH=18 tag="1XXXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItem" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopItemCd ">
												   <INPUT TYPE=TEXT NAME="txtConItemNm" SIZE=40 MAXLENGTH=40 tag="14"></TD>
				</TR>
				<TR HEIGHT=20>
					<TD CLASS="TD5"></TD>
					<TD CLASS="TD6"></TD>				
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE WIDTH=100%>
				      <TR>
				         <TD WIDTH=10>&nbsp;</TD>
				         <TD><BUTTON NAME="btnRun" ONCLICK="vbscript:ExeReflect()" CLASS="CLSMBTN" flag=1>생성</BUTTON></TD>
						 <TD WIDTH=* ALIGN=RIGHT><a href = "VBSCRIPT:JumpChgCheck(BIZ_JUMP_ID)">품목별공장배분비등록</a></TD>
						 <TD WIDTH=10>&nbsp;</TD> 
				      </TR>
				</TABLE>
			</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
