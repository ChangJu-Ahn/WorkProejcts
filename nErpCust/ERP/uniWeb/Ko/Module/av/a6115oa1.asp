<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Account Management
'*  3. Program ID           : A6115OA1
'*  4. Program Name         : �������������� 
'*  5. Program Desc         :
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/10/17
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Nam Yo
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc"  -->								<!-- '��: ȭ��ó��ASP���� �����۾��� �ʿ��� ���  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<!-- '��: �ش� ��ġ�� ���� �޶���, ��� ���  -->


<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                                                              '��: indicates that All variables must be declared in advance 
'=======================================================================================================
Dim lgMpsFirmDate, lgLlcGivenDt											 '��: �����Ͻ� ���� ASP���� �����ϹǷ� Dim 

Dim lgCurName()															'�� : ���� ȭ��� �ʿ��� ��Į ���� ���� 
'Dim cboOldVal          
 Dim IsOpenPop          
'Dim lgCboKeyPress      
'Dim lgOldIndex								
'Dim lgOldIndex2    
Dim  gSelframeFlg
Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2
    


'=======================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
End Sub


'========================================================================================================= 
Sub SetDefaultVal()
	

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
			arrParam(0) = "���ݽŰ����� �˾�"				' �˾� ��Ī 
			arrParam(1) = "B_TAX_BIZ_AREA"	 				' TABLE ��Ī 
			arrParam(2) = strCode						' Code Condition
			arrParam(3) = ""							' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "���ݽŰ������ڵ�"				' �����ʵ��� �� ��Ī 

			arrField(0) = "Tax_BIZ_AREA_CD"					' Field��(0)
			arrField(1) = "Tax_BIZ_AREA_NM"					' Field��(0)
    
			arrHeader(0) = "���ݽŰ������ڵ�"				' Header��(0)
			arrHeader(1) = "���ݽŰ������"				' Header��(0)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=480px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
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
'	Description : CtrlItem Popup���� Return�Ǵ� �� setting
'----------------------------------------------------------------------------------------------------------

Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0		' ����� 
				.txtBizAreaCd.focus
				.txtBizAreaCd.value = UCase(Trim(arrRet(0)))
				.txtBizAreaNm.value = arrRet(1)
		End Select
	End With	
End Function

Function FncBtnPrint() 
	On Error Resume Next
	

	Dim Var3
	Dim Var4
	Dim Var5
	Dim Var6
	Dim Var7
	
	Dim var8
	Dim var9
	Dim var10
	
	Dim varFromDt
	Dim VarToDt
	Dim strUrl
	
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim ObjName
	
    lngPos = 0	

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    
    If (frm1.txtFromIssueDt.year & Right(("0" & frm1.txtFromIssueDt.Month),2)  > frm1.txtToIssueDt.Year & Right(("0" & frm1.txtToIssueDt.Month),2)) Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'��: "Will you destory previous data"
		frm1.txtFromIssueDt.focus
		Exit Function
    End If
	
	If Trim(frm1.txtFiscCnt.text) <> "" Then
		If IsNumeric(frm1.txtFiscCnt.text) = False Then
			IntRetCD = DisplayMsgBox("229924", "X", "X", "X")							'�ʼ��Է� check!!
			frm1.txtFiscCnt.focus
			' ���ڸ� �Է��Ͻʽÿ� 
			Exit Function
		End If
	End If
	
	
	varFromDt = frm1.txtFromIssueDt.Year & "-" & Right(("0" & frm1.txtFromIssueDt.Month),2) & "-" & "01"
	VarToDt = FilterVar(frm1.txtToIssueDt.year,"2999","SNM") & "-" & Right("0" & FilterVar(frm1.txtToIssueDt.Month,"12","SNM"),2) & "-" & "01"
	VarToDt = DateAdd("D",-1, DateAdd("M",1,cdate(VarToDt)))
				
	
	var3 = UCase(Trim(frm1.txtBizAreaCD.value))
	var4 = UniConvDateToYYYYMMDD(UNIDateClientFormat(varFromDt), parent.gDateFormat, "")
	var5 = UniConvDateToYYYYMMDD(UNIDateClientFormat(VarToDt), parent.gDateFormat, "")
	var6 = UniConvDateToYYYYMMDD(frm1.fpDateTime3.text, parent.gDateFormat, "")
	var7 = frm1.txtFiscCnt.text 
	
	If var3 = "" Then
		var3 = "%"
		frm1.txtBizAreaNM.value = ""
	Else
	    var3 = UCase(Trim(frm1.txtBizAreaCD.value))
	End If
	If var7 = "" Then var7 = "_"

	For intCnt = 1 To 3
		lngPos = instr(lngPos + 1, GetUserPath, "/")
	Next

	If gSelframeFlg = TAB1 Then
		If frm1.Rb_AB1.checked = True Then    '�� 
			StrEbrFile = "a6115ma1"
		Else                                  '�� 
			StrEbrFile = "a6115ma1a"
		End IF	
	Else
		If frm1.Rb_AB1.checked = True  Then    '���� 
			StrEbrFile = "a6116ma1"
		Else                                  'ǥ�� 
			StrEbrFile = "a6116ma2"
		End IF		
	
	End If	
		
	var8 = frm1.txtCnt.text
	var9 = UNICDbl(frm1.txtDocAmt.text)
	var10 =UNICDbl(frm1.txtLocAmt.text)	
	
	StrUrl = StrUrl & "DrawnUpDt|"	      & var6
	StrUrl = StrUrl & "|FiscCnt|"	      & var7
	StrUrl = StrUrl & "|FromIssueDt|"	  & var4
	StrUrl = StrUrl & "|ReportBizAreaCd|" & var3
	StrUrl = StrUrl & "|ToIssueDt|"	      & var5
	StrUrl = StrUrl & "|Cnt|"			  & var8
	StrUrl = StrUrl & "|DocAmt|"          & var9
	StrUrl = StrUrl & "|LocAmt|"	      & var10
	
    ObjName = AskEBDocumentName(StrEbrFile,"ebr")	
	Call FncEBRPrint(EBAction,ObjName,StrUrl)	
	
		
End Function

Function FncBtnPreview()
	On Error Resume Next
	
	Dim Var3
	Dim Var4
	Dim Var5
	Dim Var6
	Dim Var7
	
	Dim var8
	Dim var9
	Dim var10
	
	Dim varFromDt
	Dim VarToDt
	Dim strUrl
	
	Dim arrParam, arrField, arrHeader
	Dim StrEbrFile
	Dim ObjName
	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
	
   If (frm1.txtFromIssueDt.year & Right(("0" & frm1.txtFromIssueDt.Month),2)  > frm1.txtToIssueDt.Year & Right(("0" & frm1.txtToIssueDt.Month),2)) Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'��: "Will you destory previous data"
		frm1.txtFromIssueDt.focus
		Exit Function
    End If

	If Trim(frm1.txtFiscCnt.text) <> "" Then
		If IsNumeric(frm1.txtFiscCnt.text) = False Then
			IntRetCD = DisplayMsgBox("229924", "X", "X", "X")							'�ʼ��Է� check!!
			frm1.txtFiscCnt.focus
			' ���ڸ� �Է��Ͻʽÿ� 
			Exit Function
		End If
	End If
	
	
	varFromDt = frm1.txtFromIssueDt.Year & "-" & Right(("0" & frm1.txtFromIssueDt.Month),2) & "-" & "01"
	VarToDt = FilterVar(frm1.txtToIssueDt.year,"2999","SNM") & "-" & Right("0" & FilterVar(frm1.txtToIssueDt.Month,"12","SNM"),2) & "-" & "01"
	VarToDt = DateAdd("D",-1, DateAdd("M",1,cdate(VarToDt)))
				
	
	var3 = UCase(Trim(frm1.txtBizAreaCD.value))
	var4 = UniConvDateToYYYYMMDD(UNIDateClientFormat(varFromDt), parent.gDateFormat, "")
	var5 = UniConvDateToYYYYMMDD(UNIDateClientFormat(VarToDt), parent.gDateFormat, "")
	var6 = UniConvDateToYYYYMMDD(frm1.fpDateTime3.text, parent.gDateFormat, "")
	
	var7 = frm1.txtFiscCnt.text 
	
	If var3 = "" Then
		var3 = "%"
		frm1.txtBizAreaNM.value = ""
	Else
	    var3 = UCase(Trim(frm1.txtBizAreaCD.value))
	End If
	If var7 = "" Then var7 = "_"

	var8 = frm1.txtCnt.text
	var9 = UNICDbl(frm1.txtDocAmt.text)
	var10 =UNICDbl(frm1.txtLocAmt.text)

	If gSelframeFlg = TAB1 Then
		If frm1.Rb_AB1.checked = True Then    '�� 
			StrEbrFile = "a6115ma1"
		Else                                  '�� 
			StrEbrFile = "a6115ma1a"
		End IF	
	Else
		If frm1.Rb_AB1.checked = True  Then    '���� 
			StrEbrFile = "a6116ma1"
		Else                                  'ǥ�� 
			StrEbrFile = "a6116ma2"
		End IF		
	
	End If
	
	StrUrl = StrUrl & "DrawnUpDt|"	      & var6
	StrUrl = StrUrl & "|FiscCnt|"	      & var7
	StrUrl = StrUrl & "|FromIssueDt|"	  & var4
	StrUrl = StrUrl & "|ReportBizAreaCd|" & var3
	StrUrl = StrUrl & "|ToIssueDt|"	      & var5
	StrUrl = StrUrl & "|Cnt|"			  & var8
	StrUrl = StrUrl & "|DocAmt|"          & var9
	StrUrl = StrUrl & "|LocAmt|"	      & var10

    ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	Call FncEBRPreview(ObjName,StrUrl)
	
	
End Function



'========================================================================================================= 
Sub Form_Load()
    Dim svrDate
    Call LoadInfTB19029																'��: Load table , B_numeric_format
    Call AppendNumberPlace("6","16","2")
    Call AppendNumberPlace("7","16","0")
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
	Call ggoOper.FormatNumber(frm1.txtFiscCnt, "99", "0", False)	   
	Call ggoOper.FormatNumber(frm1.txtCnt, "9999999", "0", False)	   
	 
    '----------  Coding part  -------------------------------------------------------------
    Call SetToolbar("10000000000011")
    'Call SetDefaultVal
	svrDate               =UNIDateClientFormat("<%=GetSvrDate%>")
	frm1.txtFiscCnt.text	= parent.gFiscCnt
	frm1.txtFromIssueDt.text = UNIDateClientFormat(parent.gFiscStart)
	frm1.txtToIssueDt.text   = svrDate
    Call ggoOper.FormatDate(frm1.txtFromIssueDt, parent.gDateFormat, 2)
    Call ggoOper.FormatDate(frm1.txtToIssueDt, parent.gDateFormat, 2)
  
	frm1.txtFromIssueDt.focus 
	Call ClickTab1()
	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub


'=======================================================================================================
'   Event Name : txtFromIssueDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
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
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtToIssueDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToIssueDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtToIssueDt.Focus
	End If
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
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
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function
'======================================================================================================
'	���: Tab Click
'	����: Tab Click�� �ʿ��� ����� �����Ѵ�.
'=======================================================================================================
Function ClickTab1()

	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ ù��° Tab 
	gSelframeFlg = TAB1
	Call visibleTab(gSelframeFlg)
	frm1.txtBizAreaCD.focus
End Function

Function ClickTab2()

	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ �ι�° Tab 
	gSelframeFlg = TAB2
	Call visibleTab(gSelframeFlg)
	frm1.txtBizAreaCD.focus
End Function



Function visibleTab(TabGubun)
	If TabGubun = TAB1 Then
		spnTitle.innerHTML = "��������" 
		spnRdo1.innerHTML = "��" 
		spnRdo2.innerHTML = "��" 
	ElseIF TabGubun = TAB2 Then
		spnTitle.innerHTML = "��±���" 
		spnRdo1.innerHTML = "����" 
		spnRdo2.innerHTML = "ǥ��" 
	End If

End Function 
'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
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
		<TD <%=HEIGHT_TYPE_00%>><!-- ' ���� ���� --></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>��������������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">	
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()" >
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>������ǥ�����(�������)</font></td>
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
								<TD CLASS="TD5" NOWRAP><span id="spnTitle">��������</span></TD>
								<TD CLASS="TD6"><INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio2 ID=Rb_AB1 Checked><LABEL FOR=Rb_AB1><span id="spnRdo1">��</span></LABEL>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								                <INPUT TYPE="RADIO" CLASS="Radio" NAME=Radio2 ID=Rb_AB2><LABEL FOR=Rb_AB2><span id="spnRdo2">��</span></LABEL></TD>
							</TR>
							<TR>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5">���ݽŰ�����</TD>
								<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtBizAreaCD" NAME="txtBizAreaCD" SIZE=12 MAXLENGTH=10 ALT="�����" tag="12NXXU" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btn" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopUp(frm1.txtBizAreaCD.Value, 0)">&nbsp;
											    <INPUT TYPE=TEXT ID="txtBizAreaNM" NAME="txtBizAreaNM" SIZE=20 MAXLENGTH=50 ALT="�����" tag="14X" ></TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">�ŷ��Ⱓ</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtFromIssueDt CLASS=FPDTYYYYMM title=FPDATETIME ALT="������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>
												 &nbsp;~&nbsp;
											    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtToIssueDt CLASS=FPDTYYYYMM title=FPDATETIME ALT="������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">ȸ��</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtFiscCnt" style="HEIGHT: 20px; WIDTH: 30px" tag="11X6Z" ALT="ȸ��" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							
							<TR>
							 	<TD CLASS="TD5">�������Ǽ�</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 NAME="txtCnt" style="HEIGHT: 20px; WIDTH: 100px" tag="11X" ALT="�������Ǽ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">��������ȭ�ݾ�</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name=txtDocAmt style="HEIGHT: 20px; WIDTH: 150px"CLASS=FPDS115 title=FPDOUBLESINGLE tag="11X6X" ALT="��������ȭ�ݾ�"></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
							 	<TD CLASS="TD5">��������ȭ�ݾ�</TD>
								<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name=txtLocAmt style="HEIGHT: 20px; WIDTH: 150px" CLASS=FPDS115 title=FPDOUBLESINGLE tag="11X7X" ALT="��������ȭ�ݾ�" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
							</TR>
							
							<TR>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
							</TR>
						</TABLE>
						<DIV ID="TabDiv"  SCROLL="no">
						</div>
						<DIV ID="TabDiv"  SCROLL="no">
						</div>
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
					<TD><BUTTON NAME="btnPreview" CLASS="CLSSBTN" OnClick="VBScript:FncBtnPreview()" Flag = 1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint"   CLASS="CLSSBTN" OnClick="VBScript:FncBtnPrint()" Flag = 1>�μ�</BUTTON></TD>
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

