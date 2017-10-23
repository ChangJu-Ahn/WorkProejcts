<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Account Management
'*  3. Program ID           : A6106MA1
'*  4. Program Name         : 매입매출장출력 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/11/29
'*  8. Modified date(Last)  : 2000/11/29
'*  9. Modifier (First)     : Hersheys
'* 10. Modifier (Last)      : Hersheys
'* 11. Comment              :
'*                            
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
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                              '☜: indicates that All variables must be declared in advance 


 '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- 
<!-- #Include file="../../inc/lgvariables.inc" -->	
 '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim lgMpsFirmDate, lgLlcGivenDt											 '☜: 비지니스 로직 ASP에서 참조하므로 Dim 

Dim lgCurName()															'☆ : 개별 화면당 필요한 로칼 전역 변수 
'Dim cboOldVal          
 Dim IsOpenPop          
'Dim lgCboKeyPress      
'Dim lgOldIndex								
'Dim lgOldIndex2        

'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'*****************************************  2.1 Pop-Up 함수   ********************************************
'	기능: Pop-Up 
'********************************************************************************************************* 
 '------------------------------------------  OpenBp()  ---------------------------------------
'	Name : OpenBp()
'	Description : OpenAp Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function OpenBp(Byval strCode, byval iWhere)
	Dim arrRet
	Dim arrParam(5)

	If IsOpenPop = True Then Exit Function
		
	IsOpenPop = True

	arrParam(0) = strCode								'  Code Condition
   	arrParam(1) = ""							' 채권과 연계(거래처 유무)
	arrParam(2) = ""								'FrDt
	arrParam(3) = ""								'ToDt
	arrParam(4) = "T"							'B :매출 S: 매입 T: 전체 
	arrParam(5) = ""									'SUP :공급처 PAYTO: 지급처 SOL:주문처 PAYER :수금처 INV:세금계산 	
	
	arrRet = window.showModalDialog("../../Module/ar/BpPopup.asp", Array(window.Parent,arrParam), _
		"dialogWidth=780px; dialogHeight=450px; : Yes; help: No; resizable: No; status: No;")
		
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 1		' From 거래처 
				frm1.txtFromBpCd.focus
			Case 2		' To 거래처 
				frm1.txtToBpCd.focus
		End Select
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
		lgBlnFlgChgValue = True
	End If
	
End Function
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
		Case 1
			arrParam(0) = "거래처 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER" 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "거래처"					' 조건필드의 라벨 명칭 

			arrField(0) = "BP_CD"						' Field명(0)
			arrField(1) = "BP_NM"						' Field명(1)
    
			arrHeader(0) = "거래처코드"				' Header명(0)
			arrHeader(1) = "거래처명"				' Header명(1)
		Case 2
			arrParam(0) = "거래처 팝업"				' 팝업 명칭 
			arrParam(1) = "B_BIZ_PARTNER" 				' TABLE 명칭 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "거래처"					' 조건필드의 라벨 명칭 

			arrField(0) = "BP_CD"						' Field명(0)
			arrField(1) = "BP_NM"						' Field명(1)
    
			arrHeader(0) = "거래처코드"				' Header명(0)
			arrHeader(1) = "거래처명"				' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Select Case iWhere
			Case 0		' 사업장 
				frm1.txtBizAreaCd.focus
			Case 1		' From 거래처 
				frm1.txtFromBpCd.focus
			Case 2		' To 거래처 
				frm1.txtToBpCd.focus
		End Select
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
			Case 1		' From 거래처 
				.txtFromBpCd.focus
				.txtFromBpCd.value = UCase(Trim(arrRet(0)))
				.txtFromBpNm.value = arrRet(1)
			Case 2		' To 거래처 
				.txtToBpCd.focus
				.txtToBpCd.value = UCase(Trim(arrRet(0)))
				.txtToBpNm.value = arrRet(1)
		End Select
	End With
End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 
    On Error Resume Next                                                    '☜: Protect system from crashing
    
	Dim Var1, var2, var3, var4, var5, var6, var7
	Dim strUrl
	dim lngPos
	dim intCnt
	
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim ObjName

    lngPos = 0	

    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If

    If UniConvDateToYYYYMMDD(frm1.txtFromIssueDt.text, parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToIssueDt.text, parent.gDateFormat,"") Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
    End If
    
	StrEbrFile = "a6106ma1"

	If frm1.Rb_WK1.checked = True Then
		var1 = "I"
	Else
		var1 = "O"
	End If
	
	If Trim(frm1.txtBizAreaCD.value) = "" Then
		var2 = "%"
		frm1.txtBizAreaNM.value = ""
	Else
	    var2 = FilterVar(UCase(Trim(frm1.txtBizAreaCD.value)),"","SNM")
	End If

'	var3 = UNICDate(frm1.fpDateTime1.text)
'	var4 = UNICDate(frm1.fpDateTime2.Text)

	var3 = UniConvDateToYYYYMMDD(frm1.txtFromIssueDt.Text,parent.gDateFormat,"") 
	var4 = UniConvDateToYYYYMMDD(frm1.txtToIssueDt.Text,parent.gDateFormat,"") 
	
	If Trim(frm1.cboVatType.value) = "" Then
		var5 = "%"
	Else
		var5 = frm1.cboVatType.value
	End If

	If Trim(frm1.txtFromBPCd.value) = "" Then
		var6 = "*"
		frm1.txtFromBPNm.value = ""
	Else
		var6 = FilterVar(UCase(Trim(frm1.txtFromBPCd.value)),"","SNM")
	End If

	If Trim(frm1.txtToBPCd.Value) = "" Then
		var7 = "ZZZZZZZZZZ"
		frm1.txtToBPNm.value = ""
	Else
		var7 = FilterVar(UCase(Trim(frm1.txtToBPCd.Value)),"","SNM")
	End If
	
	For intCnt = 1 To 3
	    lngPos = instr(lngPos + 1, GetUserPath, "/")
	Next

	StrUrl = StrUrl & "IoFg|" & var1
	StrUrl = StrUrl & "|BizAreaCd|"	& var2
	StrUrl = StrUrl & "|FromIssueDt|" & var3
	StrUrl = StrUrl & "|ToIssueDt|" & var4
	StrUrl = StrUrl & "|VatType|" & var5
	StrUrl = StrUrl & "|FromBpCd|" & var6
	StrUrl = StrUrl & "|ToBpCd|" & var7
	
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
	
End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 
	On Error Resume Next
	
	Dim Var1, var2, var3, var4, var5, var6, var7
	Dim strUrl
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim ObjName

    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
	
    If UniConvDateToYYYYMMDD(frm1.txtFromIssueDt.text, parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToIssueDt.text, parent.gDateFormat,"") Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
    End If
	
	StrEbrFile = "a6106ma1"

	If frm1.Rb_WK1.checked = True Then
		var1 = "I"
	Else
		var1 = "O"
	End If
	
	If Trim(frm1.txtBizAreaCD.value) = "" Then
		var2 = "%"
		frm1.txtBizAreaNM.value = ""
	Else
	    var2 = FilterVar( UCase(Trim(frm1.txtBizAreaCD.value)),"","SNM")
	End If

'	var3 = UNICDate(frm1.fpDateTime1.text)
'	var4 = UNICDate(frm1.fpDateTime2.Text)

	var3 = UniConvDateToYYYYMMDD(frm1.txtFromIssueDt.Text,parent.gDateFormat,"") 
	var4 = UniConvDateToYYYYMMDD(frm1.txtToIssueDt.Text,parent.gDateFormat,"") 

	If Trim(frm1.cboVatType.value) = "" Then
		var5 = "%"
	Else
		var5 = frm1.cboVatType.value
	End If

	If Trim(frm1.txtFromBPCd.value) = "" Then
		var6 = "*"
		frm1.txtFromBPNm.value = ""
	Else
		var6 =FilterVar( UCase(Trim(frm1.txtFromBPCd.value)),"","SNM")
	End If

	If Trim(frm1.txtToBPCd.Value) = "" Then
		var7 = "ZZZZZZZZZZ"
		frm1.txtToBPNm.value = ""
	Else
		var7 = FilterVar( UCase(Trim(frm1.txtToBPCd.Value)),"","SNM")
	End If
	
	StrUrl = StrUrl & "IoFg|" & var1
	StrUrl = StrUrl & "|BizAreaCd|"	& var2
	StrUrl = StrUrl & "|FromIssueDt|" & var3
	StrUrl = StrUrl & "|ToIssueDt|" & var4
	StrUrl = StrUrl & "|VatType|" & var5
	StrUrl = StrUrl & "|FromBpCd|" & var6
	StrUrl = StrUrl & "|ToBpCd|" & var7
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
	
End Function


'===========================================  3.1.1 Form_Load()  =========================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 함수를 Call하는 부분 
'========================================================================================================= 
Sub Form_Load()
	Dim svrDate
	Dim strYear, strMonth, strDay
	Call LoadInfTB19029																'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'⊙: Lock  Suitable  Field
    
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000011")
    Call InitComboBox
    svrDate					 = "<%=GetSvrDate%>"
	Call ExtractDateFrom(svrDate, parent.gServerDateFormat, parent.gServerDateType, strYear,strMonth,strDay)
	frm1.txtFromIssueDt.focus 
	frm1.txtFromIssueDt.text =  UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)
	frm1.txtToIssueDt.text   =  UniConvYYYYMMDDToDate(parent.gDateFormat, strYear, strMonth, strDay)	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'==========================================  2.2.6 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("B9001", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboVatType ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitComboBox_Two()
	Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", " MAJOR_CD = " & FilterVar("A1008", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboApSts,lgF0  ,lgF1  ,Chr(11))
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


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    On Error Resume Next                                                   '☜: Protect system from crashing
    Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call Parent.FncExport(parent.C_SINGLE)												'☜: 화면 유형 
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
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매입매출장출력</font></td>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>입출구분</TD>
								<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK1 Checked><LABEL FOR=Rb_WK1>매입</LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio1 ID=Rb_WK2><LABEL FOR=Rb_WK2>매출</LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">세금신고사업장</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=12 MAXLENGTH=10  ALT="세금신고사업장" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;
											    <INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=20 MAXLENGTH=50  ALT="세금신고사업장" tag="14X" ></TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">발행일</TD>
								<TD CLASS="TD6"><script language =javascript src='./js/a6106oa1_fpDateTime1_txtFromIssueDt.js'></script>
								   &nbsp;~&nbsp;<script language =javascript src='./js/a6106oa1_fpDateTime2_txtToIssueDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">계산서유형</TD>
								<TD CLASS="TD6" COLSPAN=3><SELECT ID="cboVatType" NAME="cboVatType" ALT="계산서유형" STYLE="WIDTH: 200px" tag="1XX"><OPTION VALUE="" selected></OPTION></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">거래처</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtFromBPCd" NAME="txtFromBPCd" SIZE=12 MAXLENGTH=10  ALT="거래처" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtFromBpCd.Value, 1)">&nbsp;
											    <INPUT TYPE=TEXT ID="txtFromBPNm" NAME="txtFromBPNm" SIZE=20 MAXLENGTH=20  ALT="거래처" tag="14X" ></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">&nbsp;~&nbsp;</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtToBPCd" NAME="txtToBPCd" SIZE=12 MAXLENGTH=10 ALT="거래처" tag="1XNXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenBp(frm1.txtToBpCd.Value, 2)">&nbsp;
											    <INPUT TYPE=TEXT ID="txtToBPNm" NAME="txtToBPNm" SIZE=20 MAXLENGTH=20 ALT="거래처" tag="14X" ></TD>
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" OnClick="VBScript:FncBtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()"   Flag=1>인쇄</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC = "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX="-1">
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX="-1">	
</FORM>
</BODY>
</HTML>

