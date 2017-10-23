<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : ���μ� 
'*  2. Function Name        : ��Ÿ ���� 
'*  3. Program ID           : w9105mA1
'*  4. Program Name         : w9105mA1.asp
'*  5. Program Desc         : ��47ȣ �ֿ��������(��ǥ)
'*  6. Modified date(First) : 2005/02/23
'*  7. Modified date(Last)  : 2005/02/23
'*  8. Modifier (First)     : LSHSAT
'*  9. Modifier (Last)      : 
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

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../WCM/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'============================================  ���/���� ����  ====================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_MNU_ID		= "w9105mA1"
Const BIZ_PGM_ID		= "w9105mB1.asp"											 '��: �����Ͻ� ���� ASP�� 
Const EBR_RPT_ID		= "w9105OA1"

' -- �׸��� �÷� ���� 

Dim IsOpenPop  
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
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False

End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub



'============================================  �޺� �ڽ� ä���  ====================================

Sub InitComboBox()
	' ��ȸ����(����)
	Dim IntRetCD1
	
	Call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
	
End Sub




Sub InitData()
	frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"
	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"

	Call GetFISC_DATE
	
	'Exit Sub
		
End Sub




Sub GetFISC_DATE()	' ������ ��ȸ���ǿ� �����ϴ� ������,�������� �����´�.

		
End Sub

'============================================  ��ȸ���� �Լ�  ====================================


'============================================  �� �Լ�  ====================================
Sub Form_Load()
   
    Call LoadInfTB19029                                                       <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
		
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
	  
	' �����Ѱ� 
	Call AppendNumberPlace("6","15","1")
	
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)
 
	Call InitComboBox	' �����ؾ� �Ѵ�. ����� ȸ��������� �о���� ���� 
	Call ggoOper.ClearField(Document, "1")	
	Call InitData
	Call FncQuery()
	
	 
    
End Sub



'============================================  ����� �Լ�  ====================================

'============================================  �̺�Ʈ �Լ�  ====================================
Sub SetTxtDataChange()
	lgBlnFlgChgValue = True
End Sub
'============================================  �̺�Ʈ ȣ�� �Լ�  ====================================
'============================================  �������� �Լ�  ====================================

Function FncQuery() 
    Dim IntRetCD , i
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

	
<%  '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
<%  '-----------------------
    'Erase contents area
    '----------------------- %>
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
'    Call InitVariables													<%'Initializes local global variables%>
'    Call InitData 
	lgBlnFlgChgValue = False	' --     ClearField ����    SetTxtDataChange �߻��Ѵ�.                      
    															
<%  '-----------------------
    'Check condition area
    '----------------------- %>
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If    

    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>

     
    CALL DBQuery()
    
End Function

Function FncSave() 
    Dim i, sMsg
    
    FncSave = False                                                         
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    'On Error Resume Next                                                   <%'��: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    If lgBlnFlgChgValue = False Then
		Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
		Exit Function
	End If

	

    'If Verification = False Then Exit Function
    
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '��: Save db data
    
    FncSave = True                                                          
    
End Function

'========================================================================================
Function FncNew() 
    Dim IntRetCD 

    FncNew = False

  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

  '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call InitVariables
    Call InitData

    Call SetToolbar("1100100000000111")										<%'��ư ���� ���� %>
    lgIntFlgMode = parent.OPMD_CMODE

	frm1.txtCO_CD.focus

    FncNew = True

End Function


Function FncCopy() 
End Function

Function FncCancel() 
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
	    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
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


	    strVal = strVal     & "&lgStrPrevKey="		& lgStrPrevKey             '��: Next key tag

		Call RunMyBizASP(MyBizASP, strVal)   
    End With                                           '��:  Run biz logic

    DbQuery = True  
  
End Function

		
Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    Dim iNameArr
    Dim iDx, iRow, iMaxRows
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	lgBlnFlgChgValue = False
	
	'-----------------------
	'Reset variables area
	'-----------------------
	lgIntFlgMode = parent.OPMD_UMODE
	

    Call SetToolbar("1101100000000111")										<%'��ư ���� ���� %>
		
End Function


'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
    Dim lRow, lCol, lMaxRows, lMaxCols , i    
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow, lEndRow     
    Dim lRestGrpCnt 
    Dim strVal, strDel
 
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
		exit Function
	end if
    strVal = ""
    strDel = ""
    lGrpCnt = 1
    
	strVal = ""

	Frm1.txtMode.value        =  Parent.UID_M0002
	Frm1.txtFlgMode.Value	=	lgIntFlgMode
	'.txtUpdtUserId.value  =  Parent.gUsrID
	'.txtInsrtUserId.value =  Parent.gUsrID
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID) 
	
    DbSave = True                                                           
End Function


Function DbSaveOk()		
	Call InitVariables
	
    Call MainQuery()
End Function

'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal     & "&txtFISC_YEAR="       & Frm1.txtFISC_YEAR.Text      '��: Query Key        
    strVal = strVal     & "&cboREP_TYPE="        & Frm1.cboREP_TYPE.Value      '��: Query Key            
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call FncNew()
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
<SCRIPT LANGUAGE=javascript FOR=txtData EVENT=Change>
<!--
    SetTxtDataChange();
//-->
</SCRIPT>
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right></TD>
					<TD WIDTH=10>&nbsp;</TD>
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
									<TD CLASS="TD6"><script language =javascript src='./js/w9105ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script></TD>
									<TD CLASS="TD5">���θ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
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
						<TABLE <%=LR_SPACE_TYPE_60%> BORDER=0>
                            <TR>
                                <TD WIDTH="100%" VALIGN=TOP HEIGHT=100%>
									<TABLE <%=LR_SPACE_TYPE_60%> border="1" height=100% width="100%">
										<TR>
											<TD CLASS="TD61" COLSPAN=5 ALIGN=CENTER>�� ��</TD>
											<TD CLASS="TD61" COLSPAN=1 ALIGN=CENTER>�ݾ�</TD>
										</TR>
										<TR>
											<TD CLASS="TD51" width="10%" ROWSPAN=18 ALIGN=CENTER>
												<br><br>��<br><br><br>��<br><br><br>��<br><br><br><br>
											</TD>
											<TD CLASS="TD51" width="10%" ROWSPAN=4 ALIGN=CENTER>
												��<br><br>��	
											</TD>
											<TD CLASS="TD51" width="10%" ROWSPAN=2 ALIGN=CENTER>
												����� 
											</TD>
											<TD CLASS="TD51" width="30%" COLSPAN=2 >(1) ��ΰ���</TD>
											<TD CLASS="TD51" width="40%"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" COLSPAN=2>(2) ��������</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>										   
											<TD CLASS="TD51" ROWSPAN=2 ALIGN=CENTER>
												��Ÿ 
											</TD>
											<TD CLASS="TD51" COLSPAN=2 >(3) ��ΰ���</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" COLSPAN=2>(4) ��������</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" ROWSPAN=2 ALIGN=CENTER>
												��<br><br>�� 
											</TD>
											<TD CLASS="TD51" ALIGN=CENTER>
												����� 
											</TD>
											<TD CLASS="TD51" COLSPAN=2 >(5) ��ΰ���</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>									  
										<TR>
											<TD CLASS="TD51" ALIGN=CENTER>
												��Ÿ 
											</TD>
											<TD CLASS="TD51" COLSPAN=2>(6) ��ΰ���</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>	
										<TR>
											<TD CLASS="TD51" ROWSPAN=6 ALIGN=CENTER>
												��<br>��<br>��<br>��<br>�� 
											</TD>
											<TD CLASS="TD51" ROWSPAN=2 ALIGN=CENTER>
												����� 
											</TD>
											<TD CLASS="TD51" COLSPAN=2>(7) ���</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" COLSPAN=2>(8) ��ΰ���</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>										   
											<TD CLASS="TD51" ROWSPAN=4 ALIGN=CENTER>
												��Ÿ<br>(�¿���)
											</TD>
											<TD CLASS="TD51" width="10%" ROWSPAN=2 ALIGN=CENTER>���</TD>
											<TD CLASS="TD51" width="20%">(9) ��ⷮ 2,000cc����</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51">(10) ��ⷮ 2,000cc�ʰ�</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" ROWSPAN=2 ALIGN=CENTER>��ΰ���</TD>
											<TD CLASS="TD51" >(11) ��ⷮ 2,000cc����</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51">(12) ��ⷮ 2,000cc�ʰ�</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" ROWSPAN=6 ALIGN=CENTER>
												ȸ<br><br>��<br><br>�� 
											</TD>
											<TD CLASS="TD51" ROWSPAN=2 ALIGN=CENTER>
												���� 
											</TD>
											<TD CLASS="TD51" COLSPAN=2 >(13) ���¼�</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" COLSPAN=2>(14) ��ΰ���</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>										   
											<TD CLASS="TD51" ROWSPAN=2 ALIGN=CENTER>
												�ܵ� 
											</TD>
											<TD CLASS="TD51" COLSPAN=2 >(15) ���¼�</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" COLSPAN=2>(16) ��ΰ���</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>										   
											<TD CLASS="TD51" ROWSPAN=2 ALIGN=CENTER>
												��Ÿ(�ｺŬ����)
											</TD>
											<TD CLASS="TD51" COLSPAN=2 >(17) ���¼�</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
										<TR>
											<TD CLASS="TD51" COLSPAN=2>(18) ��ΰ���</TD>
											<TD CLASS="TD51"><script language =javascript src='./js/w9105ma1_txtData_txtData.js'></script></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						</DIV>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE CLASS="TB3" CELLSPACING=0>
	
		
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="bttnPreview"  CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('VIEW')" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="bttnPrint"	CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint('PRINT')"   Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>d=e