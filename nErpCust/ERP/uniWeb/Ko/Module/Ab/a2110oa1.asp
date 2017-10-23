<%@ LANGUAGE="VBSCRIPT" %>
<!--'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Basic Info.
'*  3. Program ID           : A2110OA1
'*  4. Program Name         : 계정코드출력 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2001/11/20
'*  8. Modified date(Last)  : 2002/11/26
'*  9. Modifier (First)     : 김호영 
'* 10. Modifier (Last)      : 김호영 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'======================================================================================================= -->


<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">


Option Explicit	

'========================================================================================================= 

Dim lgBlnFlgChgValue           ' Variable is for Dirty flag 
Dim lgIntFlgMode               ' Variable is for Operation Status 



'========================================================================================================= 

'----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
Dim IsOpenPop


'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed

End Sub



'========================================================================================================= 

Sub SetDefaultVal()
End Sub


'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'========================================================================================================= 
Function OpenPopup(Byval strCode, Byval Cond)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    Select case Cond
    Case "FrAcctCd"
	   arrParam(0) = "계정코드팝업"			' 팝업 명칭 
	   arrParam(1) = "A_ACCT"					' TABLE 명칭 
	   arrParam(2) = strCode      					' Code Condition
	   arrParam(3) = ""								' Name Cindition
	   arrParam(4) = ""								' Where Condition
	   arrParam(5) = "계정코드"				' 조건필드의 라벨 명칭 

       arrField(0) = "ACCT_CD"						' Field명(0)
       arrField(1) = "ACCT_NM"						' Field명(1)

       arrHeader(0) = "계정코드"			' Header명(0)
       arrHeader(1) = "계정코드명"			' Header명(1)

    Case "ToAcctCd"
	   arrParam(0) = "계정코드팝업"			' 팝업 명칭 
	   arrParam(1) = "A_ACCT"					' TABLE 명칭 
	   arrParam(2) = strCode      					' Code Condition
	   arrParam(3) = ""								' Name Cindition
	   arrParam(4) = ""								' Where Condition
	   arrParam(5) = "계정코드"				' 조건필드의 라벨 명칭 

       arrField(0) = "ACCT_CD"						' Field명(0)
       arrField(1) = "ACCT_NM"						' Field명(1)

       arrHeader(0) = "계정코드"			' Header명(0)
       arrHeader(1) = "계정코드명"			' Header명(1)

    End select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	Select case Cond
		case "FrAcctCd"
			frm1.txtFrAcctCd.focus
		case "ToAcctCd"
			frm1.txtToAcctCd.focus
	End select	
		Exit Function
	Else
		Call SetReturnVal(arrRet, Cond)
	End If	
End Function


'========================================================================================================= 

Function SetReturnVal(ByVal arrRet, ByVal field_fg)	
	Select case field_fg
		case "FrAcctCd"
			frm1.txtFrAcctCd.focus
			frm1.txtFrAcctCd.Value		= arrRet(0)
		case "ToAcctCd"
			frm1.txtToAcctCd.focus
			frm1.txtToAcctCd.Value		= arrRet(0)
	End select
End Function

'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")
    Call InitVariables 
    Call SetToolbar("10000000000011")
	frm1.txtFrAcctCd.focus
	frm1.PrintOpt1.checked = True
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Sub SetPrintCond(StrEbrFile,FrAcctCd,ToAcctCd)


	if frm1.PrintOpt1.checked = True then
	    StrEbrFile = "a2110oa1"
	else
	    StrEbrFile = "a2110oa2"
	end if

	If Len(frm1.txtFrAcctCd.value) < 1 Then
		FrAcctCd	= "0"
	Else
		FrAcctCd	= Trim(frm1.txtFrAcctCd.value)
	End If

	If Len(frm1.txtToAcctCd.value) < 1 Then
		ToAcctCd	= "ZZZZZZZZZZZZZZZZZZZZ"
	Else
		ToAcctCd	= Trim(frm1.txtToAcctCd.value)
	End If

End Sub

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
	Dim StrUrl, StrEbrFile, FrAcctCd,ToAcctCd
	Dim IntRetCd
    Dim ObjName

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	IF frm1.txtToAcctCd.value <> "" then
		If Trim(frm1.txtFrAcctCd.value) > Trim(frm1.txtToAcctCd.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtFrAcctCd.Alt, frm1.txtToAcctCd.Alt)
			frm1.txtFrAcctCd.focus
			Exit Function
		End If
	end if

    Call SetPrintCond(StrEbrFile,FrAcctCd,ToAcctCd)

	StrUrl = StrUrl & "FrAcctCd|" & FrAcctCd
	StrUrl = StrUrl & "|ToAcctCd|" & ToAcctCd

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPrint(EBAction,ObjName,StrUrl)
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	Dim StrUrl, StrEbrFile, FrAcctCd,ToAcctCd
	Dim IntRetCd
    Dim ObjName

    If Not chkField(Document, "1") Then
       Exit Function
    End If

	IF frm1.txtToAcctCd.value <> "" then
		If Trim(frm1.txtFrAcctCd.value) > Trim(frm1.txtToAcctCd.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtFrAcctCd.Alt, frm1.txtToAcctCd.Alt)
			frm1.txtFrAcctCd.focus
			Exit Function
		End If
	end if

    Call SetPrintCond(StrEbrFile,FrAcctCd,ToAcctCd)

	StrUrl = StrUrl & "FrAcctCd|" & FrAcctCd
	StrUrl = StrUrl & "|ToAcctCd|" & ToAcctCd
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
End Function


'========================================================================================
Function FncPrint()
    Call Parent.FncPrint()
End Function

'=======================================================================================================
Function FncFind()
	Call parent.FncFind(parent.C_SINGLE, False)
End Function


'========================================================================================
Function FncExit()
    FncExit = True
End Function



</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

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
<!--
'#########################################################################################################
'       					6. Tag부 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0 >
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><% ' 상위 여백 %></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>계정코드출력</font></td>
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
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>출력구분</TD>
								<TD CLASS="TD6" NOWRAP>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" CHECKED ID="PrintOpt1" VALUE="Y" tag="25"><LABEL FOR="PrintOpt1">상세</LABEL></SPAN>
									<SPAN STYLE="WIDTH: 80px"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="PrintOpt" ID="PrintOpt2" VALUE="N" tag="25"><LABEL FOR="PrintOpt2">요약</LABEL></SPAN></TD>
							</TR>
							<TR>
           						<TD CLASS="TD5" NOWRAP>계정코드</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtFrAcctCd" SIZE=15 MAXLENGTH=20 tag="11XXXU" ALT="시작계정코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtFrAcctCd.Value, 'FrAcctCd')">&nbsp;~&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtToAcctCd" SIZE=15 MAXLENGTH=20 tag="11XXXU" ALT="종료계정코드" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtToAcctCd.Value, 'ToAcctCd')"></TD>
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
					<TD><BUTTON NAME="btn배치" CLASS="CLSSBTN" OnClick="VBScript:Call FncBtnPreview()" Flag = 1>미리보기</BUTTON> &nbsp<BUTTON NAME="btn배치" CLASS="CLSSBTN" OnClick="VBScript:Call FncBtnPrint()" Flag = 1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  src="../../blank.htm"  WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1" >
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1" >
</FORM>
</BODY>
</HTML>
