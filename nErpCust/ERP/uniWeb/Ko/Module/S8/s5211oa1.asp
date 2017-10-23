<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S5211OA1
'*  4. Program Name         : B/L관리대장출력 
'*  5. Program Desc         : B/L관리대장출력 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/06/28
'*  8. Modified date(Last)  : 2000/06/28
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
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


Option Explicit																	'☜: indicates that All variables must be declared in advance
'==========================================================================================================
Dim iDBSYSDate
Dim EndDate,StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gDateFormat)
'=========================================================================================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgIntGrpCount              ' initializes Group View Size
Dim IsOpenPop          

'=========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
        
End Sub
'=========================================================================================================
Sub SetDefaultVal()

	frm1.txtConApplicant.focus 
	frm1.txtLCFromDt.Text = StartDate
	frm1.txtLCToDt.Text = EndDate
	
End Sub
'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "OA") %>	
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "OA") %>	
End Sub
'---------------------------------------------------------------------------------------------------------
Function OpenConPop()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

		arrParam(0) = "수입자"							<%' 팝업 명칭 %>
		arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>
		arrParam(2) = Trim(frm1.txtConApplicant.value)		<%' Code Condition%>
		arrParam(3) = ""	                                <%' Name Cindition%>
		arrParam(4) = "BP_TYPE IN (" & FilterVar("CS", "''", "S") & "," & FilterVar("C", "''", "S") & " )"			    <%' Where Condition%>
		arrParam(5) = "수입자"	                        <%' TextBox 명칭 %>
		
		arrField(0) = "BP_CD"								<%' Field명(0)%>
		arrField(1) = "BP_NM"								<%' Field명(1)%>
	   								
		arrHeader(0) = "수입자"							<%' Header명(0)%>
	    arrHeader(1) = "수입자명"

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	frm1.txtConApplicant.focus 
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPop(arrRet)
	End If

End Function
'---------------------------------------------------------------------------------------------------------
Function SetConPop(Byval arrRet)
	With frm1	
				.txtConApplicant.Value		= arrRet(0)
				.txtConApplicantNm.Value	= arrRet(1)
	End With

End Function
'=========================================================================================================
Sub Form_Load()
                                                         '⊙: Load Common DLL
    Call LoadInfTB19029														'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitVariables														'⊙: Initializes local global variables
    <% '----------  Coding part  -------------------------------------------------------------%>
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")											'⊙: 버튼 툴바 제어 

End Sub
'==========================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
   
End Sub
'*********************************************************************************************************
Sub txtLCFromDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtLCFromDt.Action = 7
		Call SetFocusToDocument("M")   
		frm1.txtLCFromDt.Focus				
    End If
End Sub

Sub txtLCToDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtLCToDt.Action = 7
		Call SetFocusToDocument("M")   
		frm1.txtLCToDt.Focus				
    End If
End Sub
'========================================================================================
 Function FncPrint() 
	Call parent.FncPrint()
End Function
'========================================================================================
 Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function
'========================================================================================
Function BtnPrint() 
	Dim strUrl
    
    '** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	If ValidDateCheck(frm1.txtLCFromDt, frm1.txtLCToDt) = False Then Exit Function

    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If

    <%'--출력조건을 지정하는 부분 수정 %>
	dim var1, var2 ,var3
	
	
	If UCase(frm1.txtConApplicant.value) = "" Then
		var1 = "%"
	Else
		var1 = FilterVar(Trim(UCase(frm1.txtConApplicant.value)), "" ,  "SNM")
	End If
    
  
		var2 = UniConvDateToYYYYMMDD(frm1.txtLCFromDt.Text,Parent.gDateFormat,Parent.gServerDateType)
		var3 = UniConvDateToYYYYMMDD(frm1.txtLCToDt.Text,Parent.gDateFormat,Parent.gServerDateType)

	
	<%'--출력조건을 지정하는 부분 수정 - 끝 %>
	
'    On Error Resume Next                                                    '☜: Protect system from crashing
    

    <%'--출력조건을 지정하는 부분 수정 %>
	strUrl = strUrl & "ConApplicant|" & var1
	strUrl = strUrl & "|LCFromDt|" & var2
	strUrl = strUrl & "|LCToDt|" & var3 


		'----------------------------------------------------------------
		' Print 함수에서 호출 
		'----------------------------------------------------------------
			ObjName = AskEBDocumentName("s5211oa1","ebr")
	
			Call FncEBRprint(EBAction, ObjName, strUrl)
		'----------------------------------------------------------------
		
End Function
'========================================================================================
Function BtnPreview() 
	'** ValidDateCheck(pObjFromDt, pObjToDt) : 'pObjToDt'이 'pObjFromDt'보다 크거나 같아야 할때 **
	If ValidDateCheck(frm1.txtLCFromDt, frm1.txtLCToDt) = False Then Exit Function


    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
	Dim var1, var2, var3
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader
	If UCase(frm1.txtConApplicant.value) = "" Then
		var1 = "%"
	Else
		var1 =  FilterVar(Trim(UCase(frm1.txtConApplicant.value)), "" ,  "SNM")
	End If
    
  
 		var2 = UniConvDateToYYYYMMDD(frm1.txtLCFromDt.Text,Parent.gDateFormat,Parent.gServerDateType)
		var3 = UniConvDateToYYYYMMDD(frm1.txtLCToDt.Text,Parent.gDateFormat,Parent.gServerDateType)


	strUrl = strUrl & "ConApplicant|" & var1
	strUrl = strUrl & "|LCFromDt|" & var2
	strUrl = strUrl & "|LCToDt|" & var3 
		
		ObjName = AskEBDocumentName("s5211oa1","ebr")
	
		Call FncEBRPreview(ObjName, strUrl)
		
End Function
'========================================================================================
Function FncExit()
	FncExit = True
End Function
'########################################################################################
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>B/L관리대장</font></td>
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
									<TD CLASS=TD5 NOWRAP>수입자</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConApplicant" ALT="수입자" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConApplicant" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPop()">&nbsp;<INPUT NAME="txtConApplicantNm" TYPE="Text" SIZE=25 tag="14"></TD>
								</TR>
								
								 <TR>
									<TD CLASS="TD5" NOWRAP>발행일</TD>
									<TD CLASS=TD6 NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s5211oa1_fpDateTime1_txtLCFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s5211oa1_fpDateTime2_txtLCToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
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
						    <TD WIDTH=10>&nbsp;</TD>
							<TD>
							    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;
							    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPrint()" Flag=1>인쇄</BUTTON>
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
