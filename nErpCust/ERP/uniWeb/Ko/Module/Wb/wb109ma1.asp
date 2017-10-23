<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : �������� 
'*  3. Program ID           : WB117MA1
'*  4. Program Name         : WB117MA1.asp
'*  5. Program Desc         : �۾�������ȸ �� ���� 
'*  6. Modified date(First) : 2005/02/14
'*  7. Modified date(Last)  : 2005/02/14
'*  8. Modifier (First)     : �ֿ��� 
'*  9. Modifier (Last)      : �ֿ��� 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  �α����� ������ �����ڵ带 ����ϱ� ����  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<STYLE>

	.STATUS_FLG1 {
		color: red;
	}
	.STATUS_FLG2 {
		color: #228b22;
	}
	.STATUS_FLG3 {
		color: blue;
	}
	.link1 {
		color: black;
	}
</STYLE>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID = "wb109ma1"
Const BIZ_PGM_ID = "wb109mb1.asp"											 '��: �����Ͻ� ���� ASP�� 


Dim IsOpenPop    
Dim gSelframeFlg , lgCurrGrid      
Dim lgStrPrevKey2
Dim lgRefMode	' ���۷��� �������� : ERP ���������� 

'============================================  �ʱ�ȭ �Լ�  ====================================
Sub InitSpreadPosVariables()

End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  �Ű��� �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	Dim IntRetCD1
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
End Sub

Sub InitSpreadSheet()
	Dim ret
	
'    Call initSpreadPosVariables()  

End Sub


'============================================  �׸��� �Լ�  ====================================

Sub InitSpreadComboBox()
End Sub

Sub SetSpreadLock()

End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
 
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
 
End Sub

Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
	
End Sub

'============================================  ��ȸ���� �Լ�  ====================================

'DB���� �ε� �� ���� 
Function ShowDocList(Byval pMnuID)
	
	' XML ����Ÿ �������� 
	Dim xmlDoc, sURL, sXML
	Dim httpDoc		
	Set httpDoc = createObject("MSXML2.XMLHTTP")			
	Set xmlDoc = createObject("MSXML2.DOMDocument")
	
	sURL = BIZ_PGM_ID & "?txtFISC_YEAR=" & Frm1.txtFISC_YEAR.Text
	sURL = sURL & "&cboREP_TYPE=" & Frm1.cboREP_TYPE.Value
	sURL = sURL & "&txtMNU_ID=" & pMnuID
		
	httpDoc.open "GET", sURL, false	
		
	httpDoc.send
		
	sXML =  httpDoc.responseText
		
	Set httpDoc = Nothing 
	
	xmlDoc.loadXML  sXML 

	If xmlDoc.parseError.reason <> "" Then
		MsgBox xmlDoc.parseError.reason 
		Set xmlDoc = Nothing
		Exit Function
	End If

	' �޴�XML �� HTML ���� 
	Dim oNode, oNodeList, sHTML
	' VB�޴� 
	sHTML =  MakeMenuHTML(xmlDoc)
	divDoc.innerHTML = sHTML
		
	Set xmlDoc = Nothing
End Function

Function MakeMenuHTML(Byref pxmlDoc)
	Dim oNodeList, oNode, sHTML, sStatusFlg, sMnuNm, sTmp, sMnuID
	
	Set oNodeLIst = pxmlDoc.selectNodes("//row")

	sHTML = "<TABLE CLASS='BasicTB' CELLSPACING=0 border=1>" & vbCrLf
	sHTML = sHTML & "<TR HEIGHT=10>" & vbCrLf
	sHTML = sHTML & "	<TD COLSPAN=2 align=center CLASS=TD61>�������� ���ĸ��</TD>" & vbCrLf
	sHTML = sHTML & "</TR>" & vbCrLf
	sHTML = sHTML & "<TR HEIGHT=10>" & vbCrLf
	sHTML = sHTML & "	<TD WIDTH=80% CLASS=TD51>���ĸ�</TD>" & vbCrLf
	sHTML = sHTML & "	<TD WIDTH=20% CLASS=TD51>���൵</TD>" & vbCrLf
	sHTML = sHTML & "</TR>" & vbCrLf


	For Each oNode In oNodeLIst
		sStatusFlg	= oNode.attributes.getNamedItem("STATUS_FLG").text 
		sMnuNm		= oNode.attributes.getNamedItem("MNU_NM").text 
		sMnuID		= oNode.attributes.getNamedItem("MNU_ID").text 
		
		sTmp		= MakeStatus(sStatusFlg)
		
		sHTML = sHTML & "<TR HEIGHT=10>" & vbCrLf
		sHTML = sHTML & "<TD><a href=javascript:PgmJump('" & sMnuID & "') class='link1'>" & sMnuNm & "</a></TD>" & vbCrLf
		sHTML = sHTML & "<TD align=center>" & sTmp & "</TD>" & vbCrLf
		sHTML = sHTML & "</TR>" & vbCrLf

	Next
	sHTML = sHTML & "<TR height=*>" & vbCrLf
	sHTML = sHTML & "<TD>&nbsp;</TD><TD>&nbsp;</TD>" & vbCrLf
	sHTML = sHTML & "</TR>" & vbCrLf
	MakeMenuHTML = sHTML & "</TABLE>" & vbCrLf
	
	Set oNode = Nothing
	Set oNodeList = Nothing
End Function

Function MakeStatus(Byval pStatusFlg)
	Select Case pStatusFlg
		Case "1"
			MakeStatus = "<font class='STATUS_FLG1'>X</font> "	' -- ������ 
		Case "2"
			MakeStatus = "<font class='STATUS_FLG2'>��</font> "	' -- ���� 
		Case "3"
			MakeStatus = "<font class='STATUS_FLG3'>��</font> "	' -- �Ϸ� 
		Case Else
			MakeStatus = "�� "
	End Select
End Function

Function PgmJump(Byval pMnuID)
	Dim objConn , PostString
	WriteCookie "gActivePgmID",pMnuID
	
	Set objConn = CreateObject("uniConnector.cGlobal") 
	PostString = objConn.GetAspPostString 
	'window.open "../../SessionTrans.asp?" & PostString 
	
	window.open "../../uniToolbar.Asp?SLX=Y&DPCP=" & pMnuID & "&arg="
End Function

'====================================== �� �Լ� =========================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1000000000000111")										<%'��ư ���� ���� %>

	' �����Ѱ� 
    Call InitComboBox
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

	Call InitData()

     
    
    'Call FncQuery()
End Sub


'============================================  �̺�Ʈ �Լ�  ====================================
Sub txtFISC_YEAR_DblClick(Button)
    If Button = 1 Then
        frm1.txtFISC_YEAR.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFISC_YEAR.Focus
    End If
End Sub


'============================================  �������� �Լ�  ====================================

Function FncNew() 
    Dim IntRetCD 

    FncNew = False

    FncNew = True

End Function

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

<%  '-----------------------
    'Check previous data area
    '----------------------- %>
	Call MakeHTML

    
End Function

Function FncSave() 
 
    FncSave = True                                                          
    
End Function

Function FncCopy() 
 
End Function

Function FncCancel() 
                                                '��: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 

End Function


Function FncDeleteRow() 

End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '��: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'��: ȭ�� ���� %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'��:ȭ�� ����, Tab ���� %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False

    FncExit = True
End Function

'========================================================================================
Function FncDelete() 
    Dim IntRetCD

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF

    Call DbDelete

    FncDelete = True
End Function
'============================================  DB �＼�� �Լ�  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With Frm1
    
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001						         
        strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
        strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key   
        strVal = strVal     & "&txtCurrGrid="        & lgCurrGrid      
		
		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
  
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
		
End Function


</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 width=300>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD align=right><font class='STATUS_FLG1'>X</font> ������ <font class='STATUS_FLG2'>��</font> ������ <font class='STATUS_FLG3'>��</font> �Ϸ�&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">�������</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/wb109ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
									<TD CLASS="TD5">���θ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">�Ű���</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="�Ű���" STYLE="WIDTH: 50%" tag="14X"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
					<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%; overflow=auto"><% ' -- overflow=auto : ������ ������ ������ ũ�⿡ ���� ��ũ�ѹٰ� �����ǰ� �Ѵ� %>
						<TABLE <%=LR_SPACE_TYPE_20%> border=0>
							<TR>
								<TD WIDTH=70% valign=top align=center><img src="WB109MA1.JPG" usemap="#imagemap" width="513" height="435" border="0">
								</TD>
								<TD WIDTH=30% VALIGN=TOP>
								<DIV ID=divDoc>
								<TABLE <%=LR_SPACE_TYPE_20%> border=1 valign=top>
									<TR HEIGHT=10>
										<TD COLSPAN=2 align=center  CLASS=TD61>�������� ���ĸ��</TD>
									</TR>
									<TR HEIGHT=10>
										<TD WIDTH=80% CLASS=TD51>���ĸ�</TD>
										<TD WIDTH=20% CLASS=TD51>���൵</TD>
									</TR>
									<TR HEIGHT=*>
										<TD>&nbsp;&nbsp;</TD>
										</TD>&nbsp;&nbsp;</TD>
									</TR>
								</TABLE>
								</DIV>
								</TD>
							</TR>
						</TABLE>

					</TD>
				</TR>
			</TABLE>
			</DIV>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<map name="imagemap">
  <area shape="rect" alt="B. ��������" coords="184,3,330,39" href="javascript:ShowDocList('WB')">
  <area shape="rect" alt="1. �繫��ǥ �ۼ�" coords="184,45,331,81" href="javascript:ShowDocList('W1')">
  <area shape="rect" alt="2. ���Աݾ� ����" coords="184,116,331,151" href="javascript:ShowDocList('W2')">
  <area shape="rect" alt="3. �� ��������" coords="3,188,151,223" href="javascript:ShowDocList('W3')">
  <area shape="rect" alt="4. �غ������" coords="364,188,510,224" href="javascript:ShowDocList('W4')">
  <area shape="rect" alt="5. �ҵ�ݾ�����" coords="184,254,331,289" href="javascript:ShowDocList('W5')">
  <area shape="rect" alt="6. �������鼼������" coords="4,324,150,361" href="javascript:ShowDocList('W6')">
  <area shape="rect" alt="7. ��Ÿ��������" coords="366,324,512,359" href="javascript:ShowDocList('W7')">
  <area shape="rect" alt="8. ���μ�Ȯ��" coords="184,323,331,360" href="javascript:ShowDocList('W8')">
  <area shape="rect" alt="9. ��Ÿ����" coords="184,396,331,431" href="javascript:ShowDocList('W9')">
  <area shape="rect" alt="10. ���ڽŰ�" coords="367,395,511,430" href="javascript:ShowDocList('W10')">
</map>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<TEXTAREA CLASS="hidden" NAME="txtSpread2" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCurrGrid" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

