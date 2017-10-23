<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*
'*  1. Module Name          : HR
'*  2. Function Name        : 
'*  3. Program ID           : h9121ma1
'*  4. Program Name         : (��������Ű������������)
'*  5. Program Desc         : ��������Ű������������ 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/04/18
'*  8. Modified date(Last)  : 2003/06/13
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Lee SiNa
'* 11. Comment              :
'*                            
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->				

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">	

<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

'============================================  1.2.1 Global ��� ����  ====================================

Const BIZ_PGM_ID = "h9121mb1.asp"											 '��: �����Ͻ� ���� ASP�� 

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '��: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '��: Indicates that no value changed
    lgIntGrpCount = 0                                                       '��: Initializes Group View Size
    IsOpenPop = False														'��: ����� ���� �ʱ�ȭ 
End Sub

'============================================= 2.1.2 LoadInfTB19029() ====================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("I", "H", "NOCOOKIE", "MA") %>
End Sub
'==========================================  2.4.3 Set???()  ===============================================
'	Name : OpenyearareaInfo()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 

Function OpenyearareaInfo(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��������Ű����� �˾�"			' �˾� ��Ī 
	arrParam(1) = "HFA100T"						' TABLE ��Ī 
	arrParam(2) = strCode						' Code Condition
	arrParam(3) = ""							' Name COndition
	arrParam(4) = ""							' Where Condition
	arrParam(5) = "��������Ű�����"			
	
    arrField(0) = "YEAR_AREA_CD"				' Field��(0)
    arrField(1) = "YEAR_AREA_NM"						' Field��(1)
    
    arrHeader(0) = "��������Ű����� �ڵ�"					' Header��(0)
    arrHeader(1) = "��������Ű������"					' Header��(1)
    
	arrRet = window.showModalDialog("../../comasp/commonpopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
    	frm1.txtYearAreaCd.focus	
	    Exit Function
	Else
		With frm1
			.txtYearAreaCd.value = arrRet(0)
			.txtYearAreaNm.value = arrRet(1)
			.txtYearAreaCd.focus
		End With

	End If	

End Function

'==========================================  3.1.1 Form_Load()  ===========================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'==========================================================================================================
Sub Form_Load()

    Call InitVariables																'��: Initializes local global variables
    Call LoadInfTB19029																'��: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)	'��: Format Numeric Contents Field
    Call ggoOper.LockField(Document, "N")											'��: Lock  Suitable  Field
    Call SetToolBar("1110100000001111")
	
	frm1.txtYearAreaCd.focus	
End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '��: Processing is NG
    Err.Clear                                                               '��: Protect system from crashing

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")				'��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If
    frm1.txtYearAreaNm.value = ""
    
    Call DbQuery																'��: Query db data

    FncQuery = True																'��: Processing is OK
        
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                     '��: Processing is NG
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X")           '��: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")                                      '��: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field
    Call InitVariables															'��: Initializes local global variables
    
    Call SetToolBar("1110100000001111")

	frm1.txtYearAreaCd.focus
	
    FncNew = True																'��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
    Dim IntRetCD
    
    FncDelete = False														'��: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003", parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF	
    
    Call DbDelete															'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing
    
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")                          '��: No data changed!!
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                             '��: Check contents area
       Exit Function
    End If
	if txtYearAreaCd_OnChange() then
		Exit Function
	end if    

    Call DbSave				                                                '��: Save db data
    
    FncSave = True                                                          '��: Processing is OK
    
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE											'Indicates that current mode is Crate mode
    
     ' ���Ǻ� �ʵ带 �����Ѵ�. 
    Call ggoOper.ClearField(Document, "1")                              'Clear Condition Field
    Call ggoOper.LockField(Document, "N")								'This function lock the suitable field
    
	lgBlnFlgChgValue = True

    frm1.txtYearAreaCd_Body.value = ""

    frm1.txtYearAreaCd_Body.focus
    
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
    On Error Resume Next                                                    '��: Protect system from crashing
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    On Error Resume Next                                                    '��: Protect system from crashing
    
    parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    Dim strVal
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '��: 
        Exit Function
    ElseIf lgPrevNo = "" then
		Call DisplayMsgBox("900011", "X", "X", "X")
	End IF	
        
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						   '��: 
    strVal = strVal & "&txtYearAreaCd = " & lgPrevNo							   '��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    Dim strVal

    If lgIntFlgMode <> parent.OPMD_UMODE Then                                     'Check if there is retrived data
        Call DisplayMsgBox("900002", "X", "X", "X")                                '��: 
        Exit Function
    End If
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						  '��: �����Ͻ� ó�� ASP�� ���°� 
    strVal = strVal & "&txtYearAreaCd=" & lgNextNo							  '��: ��ȸ ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)

End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(parent.C_SINGLE)												'��: ȭ�� ���� 
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
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")			'��: "Will you destory previous data"
		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================

Function DbDelete() 
    Err.Clear                                                               '��: Protect system from crashing
    
    DbDelete = False														'��: Processing is NG
    
    Call LayerShowHide(1)                                                   '��: Protect system from crashing
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtYearAreaCd_Body=" & Trim(frm1.txtYearAreaCd_Body.value)				'��: ���� ���� ����Ÿ 
    strVal = strVal & "&txtOwnRgstNo=" & Trim(frm1.txtOwnRgstNo.value)				'��: ���� ���� ����Ÿ 
    
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbDelete = True                                                         '��: Processing is NG

End Function


'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================

Function DbDeleteOk()														'��: ���� ������ ���� ���� 
	lgBlnFlgChgValue = False
	Call FncNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    
    Err.Clear                                                               '��: Protect system from crashing
    
    DbQuery = False                                                         '��: Processing is NG
    Call LayerShowHide(1)                                                   '��: Protect system from crashing
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'��: �����Ͻ� ó�� ASP�� ���� 
    strVal = strVal & "&txtYearAreaCd=" & Trim(frm1.txtYearAreaCd.value)				'��: ��ȸ ���� ����Ÿ 
    
    
    call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
	
    DbQuery = True                                                          '��: Processing is NG

End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()														'��: ��ȸ ������ ������� 

    Call SetToolBar("1111100000011111")
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field

    lgIntFlgMode = parent.OPMD_UMODE												'��: Indicates that current mode is Update mode
    frm1.txtYearAreaNm_Body.focus
End Function


'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave() 

    Err.Clear																'��: Protect system from crashing

	DbSave = False															'��: Processing is NG

    Dim strVal
    Call LayerShowHide(1)                                                   '��: Protect system from crashing

	With frm1
		.txtMode.value = parent.UID_M0002											'��: �����Ͻ� ó�� ASP �� ���� 
		.txtFlgMode.value = lgIntFlgMode
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With
	
    DbSave = True                                                           '��: Processing is NG
    
End Function


'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================

Function DbSaveOk()															'��: ���� ������ ���� ���� 

    frm1.txtYearAreaCd.value = frm1.txtYearAreaCd_Body.value 
    lgBlnFlgChgValue = False
    FncQuery

End Function

Function txtYearAreaCd_OnChange()
    If  frm1.txtYearAreaCd.value = "" Then
        frm1.txtYearAreaNm.value = ""
        frm1.txtYearAreaCd.focus
        Set gActiveElement = document.ActiveElement
    Else
        If  CommonQueryRs(" YEAR_AREA_NM "," HFA100T "," YEAR_AREA_CD =  " & FilterVar(frm1.txtYearAreaCd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            Call  DisplayMsgBox("970000", "x","��������Ű����� �ڵ�","x")

            frm1.txtYearAreaNm.value = ""
	        frm1.txtYearAreaCd.focus
	        Set gActiveElement = document.ActiveElement
			txtYearAreaCd_OnChange = true	        
	        exit function
	    Else
	        frm1.txtYearAreaNm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>��������Ű�������</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH="100%" CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH="100%">
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>��������Ű�����</TD>
									<TD CLASS="TD656" NOWRAP><INPUT NAME="txtYearAreaCd" MAXLENGTH="10" SIZE=10 ALT ="��������Ű����� �ڵ�" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenyearareaInfo(frm1.txtYearAreaCd.value,0)">
													         <INPUT NAME="txtYearAreaNm" MAXLENGTH="30" SIZE=30 ALT ="��������Ű������" tag="14X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>��������Ű����� �ڵ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtYearAreaCd_Body" ALT="��������Ű����� �ڵ�" MAXLENGTH="10" SIZE=10 tag = "23XXXU"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��������Ű������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtYearAreaNm_Body" ALT="��������Ű������" MAXLENGTH="30" SIZE=30 tag="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����ڵ�Ϲ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOwnRgstNo" ALT="����ڵ�Ϲ�ȣ" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" tag ="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>���ε�Ϲ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCoownRgstNo" ALT="���ε�Ϲ�ȣ" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" tag ="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>��ǥ�ڸ�</TD>
    						    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepreNm" ALT="��ǥ�ڸ�" MAXLENGTH="30" SIZE=30 STYLE="TEXT-ALIGN:left" tag="22"></TD>				    					    			
							</TR>
							<TR>
 							    <TD CLASS=TD5 NOWRAP>��ȭ��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTelNo" ALT="��ȭ��ȣ" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" tag  ="22"></TD>
 							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�ּ�</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtAddr"  ALT="�ּ�"     MAXLENGTH="100" SIZE="100" STYLE="TEXT-ALIGN:left" tag="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�������ڵ�/��������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTaxOffice"  ALT="�������ڵ�"  MAXLENGTH="3" SIZE="10" STYLE="TEXT-ALIGN:left" tag="22">
								                     <INPUT NAME="txtTaxOfficeNm"  ALT="��������"  MAXLENGTH="30" SIZE="30" STYLE="TEXT-ALIGN:left" tag="22"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����ڸ�</TD>
    						    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtWorkerNm" ALT="����ڸ�" MAXLENGTH="30" SIZE=30 STYLE="TEXT-ALIGN:left" tag="22"></TD>				    					    			
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>����ںμ�</TD>
    						    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtWorkerDeptNm" ALT="����ںμ�" MAXLENGTH="30" SIZE=30 STYLE="TEXT-ALIGN:left" tag="22"></TD>				    					    			
							</TR>
							<TR>
 							    <TD CLASS=TD5 NOWRAP>�������ȭ��ȣ</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtWorkerTel" ALT="�������ȭ��ȣ" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" tag  ="22"></TD>
 							</TR>	
							<TR>
								<TD CLASS=TD5 NOWRAP>Ȩ�ؽ�ID</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtHometaxID" ALT="Ȩ�ؽ�ID" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" tag ="21"></TD>
							</TR> 														
							<% Call SubFillRemBodyTd56(2) %>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=hidden NAME="txtMode" tag="24" tabindex="-1">
<INPUT TYPE=hidden NAME="txtFlgMode" tag="24" tabindex="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

