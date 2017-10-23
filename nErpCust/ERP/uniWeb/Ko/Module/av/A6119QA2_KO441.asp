<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Account Management
'*  3. Program ID           : A6110MA1
'*  4. Program Name         : 매입세액 불공제분 계산근거 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/12/13
'*  8. Modified date(Last)  : 2001/01/03
'*  9. Modifier (First)     : SHH
'* 10. Modifier (Last)      : SHH
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '⊙: 화면처리ASP에서 서버작업이 필요한 경우  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '☆: 해당 위치에 따라 달라짐, 상대 경로  -->

<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                              '☜: indicates that All variables must be declared in advance 

'========================================================================================================= 
Dim lgMpsFirmDate, lgLlcGivenDt											 '☜: 비지니스 로직 ASP에서 참조하므로 Dim 

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
'Dim cboOldVal          
 Dim IsOpenPop          
'Dim lgCboKeyPress      
'Dim lgOldIndex								
'Dim lgOldIndex2        

'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'------------------------------------------  OpenPopUp()  -------------------------------------------------
'	Name : OpenPopUp()
'	Description : PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere		
		Case 0
			arrParam(0) = "세금신고사업장 팝업"				' 팝업 명칭 
			arrParam(1) = "B_TAX_BIZ_AREA"	 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "세금신고사업장코드"				' 조건필드의 라벨 명칭 

			arrField(0) = "TAX_BIZ_AREA_CD"					' Field명(0)
			arrField(1) = "TAX_BIZ_AREA_NM"					' Field명(0)
    
			arrHeader(0) = "세금신고사업장코드"				' Header명(0)
			arrHeader(1) = "세금신고사업장명"				' Header명(0)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/Commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function


'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : CtrlItem Popup에서 Return되는 값 setting
'----------------------------------------------------------------------------------------------------------

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' 사업장 
				.txtBizAreaCd.focus
				.txtBizAreaCd.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNm.value = arrRet(1)
		End Select
	End With	
End Function

Function FncBtnPrint() 
	On Error Resume Next
	
	Dim Var1
	Dim Var2
	Dim Var3
	Dim Var4
	Dim strUrl
	
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim ObjName
	
    lngPos = 0	

    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    If UniConvDateToYYYYMMDD(frm1.txtFromIssueDt.text + "-01", parent.gDateFormat, "") > UniConvDateToYYYYMMDD(frm1.txtToIssueDt.text + "-01", parent.gDateFormat, "") Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
    End If

	
	var1 = FilterVar(UCase(Trim(frm1.txtBizAreaCD.value)),"","SNM")
	var2 = replace(frm1.fpDateTime1.text,"-","")
	var3 = replace(frm1.fpDateTime2.text,"-","")
	var4 = UniConvDateToYYYYMMDD(frm1.fpDateTime3.text, parent.gDateFormat, "")
	
	If var1 = "" Then
		var1 = "%"
		frm1.txtBizAreaNM.value = ""
	Else
	    var1 = FilterVar(UCase(Trim(frm1.txtBizAreaCD.value)),"","SNM")
	End If
	
	For intCnt = 1 To 3
		lngPos = instr(lngPos + 1, GetUserPath, "/")
	Next
	
	StrEbrFile = "a6119oa2_KO441"
	
	StrUrl = StrUrl & "bizareacd|"	& var1
	StrUrl = StrUrl & "|frdt|"		& var2
	StrUrl = StrUrl & "|todt|"		& var3
	StrUrl = StrUrl & "|dt|"		& var4
	
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
	
End Function

Function FncBtnPreview()
	On Error Resume Next
	
	Dim Var1
	Dim Var2
	Dim Var3
	Dim Var4
	Dim strUrl
	
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim ObjName
	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	
    If UniConvDateToYYYYMM(frm1.txtFromIssueDt.text + "-01", parent.gDateFormat, "") > UniConvDateToYYYYMM(frm1.txtToIssueDt.text + "-01", parent.gDateFormat, "") Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
    End If

	
	var1 = FilterVar(UCase(Trim(frm1.txtBizAreaCD.value)),"","SNM")
	var2 = replace(frm1.fpDateTime1.text,"-","")
	var3 = replace(frm1.fpDateTime2.text,"-","")
	var4 = UniConvDateToYYYYMMDD(frm1.fpDateTime3.text, parent.gDateFormat, "")
	
	If var1 = "" Then
		var1 = "%"
		frm1.txtBizAreaNM.value = ""
	Else
	    var1 = FilterVar(UCase(Trim(frm1.txtBizAreaCD.value)),"","SNM")
	End If
	
	StrEbrFile = "a6119oa2_KO441"
	
	StrUrl = StrUrl & "bizareacd|"	& var1
	StrUrl = StrUrl & "|frdt|"		& var2
	StrUrl = StrUrl & "|todt|"		& var3
	StrUrl = StrUrl & "|dt|"		& var4
	
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
	
End Function


'===========================================  3.1.1 Form_Load()  =========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
    Dim svrDate
    Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
   
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000011")
	svrDate               = UNIDateClientFormat("<%=GetSvrDate%>")

	frm1.txtFromIssueDt.text = svrDate
	frm1.txtToIssueDt.text   = svrDate
	frm1.txtDrawnUpDt.text   = svrDate
	frm1.txtFromIssueDt.focus 
	
	Call ggoOper.FormatDate(frm1.txtFromIssueDt, parent.gDateFormat,2)
	Call ggoOper.FormatDate(frm1.txtToIssueDt, parent.gDateFormat, 2)

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'=======================================================================================================
'   Event Name : txtFromIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtFromIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromIssueDt.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtFromIssueDt.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtToIssueDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtToIssueDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToIssueDt.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtToIssueDt.Focus
    End If
End Sub


'=======================================================================================================
'   Event Name : txtDrawnDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtDrawnUpDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtDrawnUpDt.Action = 7
  		Call SetFocusToDocument("M")
		frm1.txtDrawnUpDt.Focus
    End If
End Sub

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
	Call FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '☜:화면 유형, Tab 유무 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
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

<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><!-- ' 상위 여백 --></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>매입세액불공제분계산근거출력</font></td>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
							</TR>
							
								<TR>
								<TD CLASS="TD5">세금신고사업장</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=12 MAXLENGTH=10 ALT="세금신고사업장" tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;
											    <INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=20 MAXLENGTH=50 ALT="세금신고사업장" tag="14X" ></TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">신고년월</TD>
								<TD CLASS="TD6"><OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtFromIssueDt CLASS=FPDTYYYYMM title=FPDATETIME ALT="신고년월" tag="12X1" VIEWASTEXT></OBJECT>
												 &nbsp;~&nbsp;
											    <OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtToIssueDt   CLASS=FPDTYYYYMM title=FPDATETIME ALT="신고년월" tag="12X1" VIEWASTEXT></OBJECT></TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">작성일</TD>
								<TD CLASS="TD6"><OBJECT classid=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtDrawnUpDt   CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="작성일" tag="12X1" VIEWASTEXT></OBJECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" OnClick="VBScript:FncBtnPreview()" Flag = 1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()" Flag = 1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname"  TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">	
</FORM>
</BODY>
</HTML>

