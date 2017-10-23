<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7112ma1
'*  4. Program Name         : ������ ����ȸ 
'*  5. Program Desc         : �����ڻ꺰 �������� ��ȸ 
'*  6. Comproxy List        : +As0069LookupSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2002/03/26
'*  8. Modified date(Last)  : 2002/03/26
'*  9. Modifier (First)     : Ȳ���� 
'* 10. Modifier (Last)      : Ȳ���� 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                       
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'=======================================================================================================
'												1. �� �� �� 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc ����   
'	���: Inc. Include
'======================================================================================================= -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID 		= "A7112MB1.asp"                              '��: Biz Logic ASP Name

'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const C_SHEETMAXROWS_D  = 30                                          '��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
Const C_MaxKey          = 3					                          '��: SpreadSheet�� Ű�� ���� 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgIsOpenPop                                          

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


'======================================================================================================
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'=======================================================================================================

'======================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=======================================================================================================

Sub InitVariables()
    lgStrPrevKey     = ""
'    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          'Indicates that current mode is Create mode
	lgBlnFlgChgValue = False
    lgSortKey        = 1
'    lgSaveRow        = 0

End Sub


Sub InitData
	With frm1
		.txtAcctCd.value				= ""
		.txtAcctNm.value				= ""				'������ 
		.cboDeprMthd.value				= ""			'�󰢹�� 
		.txtRegDt.value					= 0			'������� 
		.txtLocAcqAmt.value				= 0			'���ݾ�(�ڱ�)
		.txtAcqQty.value				= 0				'������ 
        .txtInvQty.Value				= 0			'������ 
		.txtDurYrs.value				= 0				'���뿬�� 
		.txtDeprRate.value				= 0			'���� 
		'''''���⸻ 
		.txtEndLTermAcqAmt.value		= 0			'��氡�� 
		.txtEndLTermCptAmt.value		= 0			'�ں������� 
		.txtEndLTermDeprAmt.value		= 0			'�󰢾� 
		.txtEndLTermBalAmt.value		= 0			'�̻󰢾� 
		.txtEndLTermInvQty.value		= 0			'��� 
		
		'''''�����			
		.txtFMnthAcqAmt.value			= 0			'��氡�� 
		.txtFMnthCptAmt.value			= 0			'�ں������� 
		.txtFMnthDeprAmt.value			= 0			'�󰢾� 
		.txtFMnthBalAmt.value			= 0			'�̻󰢾� 
		.txtFMnthInvQty.value			= 0			'��� 
		'''''����߻�			
		.txtMnthCptAmt.value			= 0			'�ں������� 
		.txtMnthDeprAmt.value			= 0			'�󰢾� 
		.txtMnthDisUseQty.value			= 0			'�Ű���ⷮ 
		'''''�����	
		.txtLMnthAcqAmt.value			= 0			'��氡�� 
		.txtLMnthCptAmt.value			= 0			'�ں������� 
		.txtLMnthDeprAmt.value			= 0			'�󰢾� 
		.txtLMnthBalAmt.value			= 0			'�̻󰢾� 
		.txtLMnthInvQty.value			= 0			'��� 
			      
	End With
End Sub

Sub InitComboBox()
	Dim strName, strCode      
	
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Call CommonQueryRs("rTrim(minor_nm), rTrim(minor_nm)", "b_minor", "major_cd=" & FilterVar("A2002", "''", "S") & " ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
	Call SetCombo2(frm1.cboDeprMthd ,lgF0  ,lgF1  ,Chr(11))
	
   	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
<% 
   BaseDate     = GetSvrDate                                                                  'Get DB Server Date
%>   

	frm1.txtDeprYyyymm.Text = UNIConvDateAToB("<%=BaseDate%>" ,parent.gServerDateFormat,parent.gDateFormat)               'Convert DB date type to Company
	Call ggoOper.FormatDate(frm1.txtDeprYyyymm, parent.gDateFormat, 2)

End Sub
'======================================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'=======================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 

	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "MA") %>  ' check

	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
    
End Sub


'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()
   
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal lRow)
    
End Sub

 '******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 
 '------------------------------------------  OpenMasterRef()  -------------------------------------------------
'	Name : OpenMasterRef()
'	Description : Asset Master Condition PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenMasterRef()


	Dim arrRet
	Dim arrParam(8)
	Dim iCalledAspName
	
	If lgIsOpenPop = True Then Exit Function
	
	iCalledAspName = AskPRAspName("a7103ra1")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "a7103ra1" ,"x")
		IsOpenPop = False
		Exit Function
	End If

	' ���Ѱ��� �߰� 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	lgIsOpenPop = True

	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
			"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPoRef(arrRet)
	End If	

		
End Function

 '------------------------------------------  SetPoRef()  -------------------------------------------------
'	Name : SetPoRef()
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub SetPoRef(strRet)       
	frm1.txtAsstNo.value = strRet(0)
	frm1.txtAsstNm.value = strRet(1)		
End Sub


'======================================================================================================
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'=======================================================================================================

'======================================================================================================
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'=======================================================================================================

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'=======================================================================================================
Sub Form_Load()
	
	Err.Clear                                                                        '��: Clear err status
    
	Call LoadInfTB19029                                                              '��: Load table , B_numeric_format
		
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
 <%   
    'Call GetAdoFieldInf("C3605MA101","S","A")                                ' S for Sort , A for SpreadSheet No('A','B',....             
%>

    Call ggoOper.LockField(Document, "N")                                   
		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	
'    lgMaxFieldCount =  UBound(gFieldNM)                      
	
'    ReDim lgPopUpR(C_MaxSelList - 1,1)

'    Call MakePopData(gDefaultT,gFieldNM,gFieldCD,lgPopUpR,lgSortFieldNm,lgSortFieldCD,C_MaxSelList)

	Call InitVariables														
    frm1.txtAsstNo.focus
	Call SetDefaultVal	
    Call SetToolbar("11000000000011")										
    Call InitComboBox
    frm1.txtAsstNo.focus

	' ���Ѱ��� �߰� 
	Dim xmlDoc

	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc)

	' ����� 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ� 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text

	' ���κμ�(��������)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text

	' ���� 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text

	Set xmlDoc = Nothing

	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'=======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'=======================================================================================================
'   Event Name : txtYyyymm_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDeprYyyymm_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub


Sub txtDeprYyyymm_DblClick(Button)
    If Button = 1 Then
       frm1.txtDeprYyyymm.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtDeprYyyymm.Focus       
    End If
End Sub

Sub txtDeprYyyymm_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call MainQuery
End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'=======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
  
End Sub

'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

End Sub

'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	
End Sub

'======================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : ��ư �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
	
End Sub

'======================================================================================================
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'=======================================================================================================

'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'=======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               'Protect system from crashing
	'-----------------------
    'Erase contents area
    '----------------------- 
    'Call ggoOper.ClearField(Document, "2")									'Clear Contents  Field
    Call InitVariables     
    Call InitData                                                   
    															
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								'This function check indispensable field
       Exit Function
    End If
    
  '-----------------------
    'Query function call area
    '----------------------- 
	IF DbQuery = False Then
		Exit Function
	END IF
	      
    FncQuery = True															
End Function


'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'=======================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'=======================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)											 '��: ȭ�� ���� 
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     '��:ȭ�� ����, Tab ���� 
End Function


'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'=======================================================================================================
Function FncExit()
Dim IntRetCD
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim strVal
	Dim DurYrsFg
	Dim strFrDt
	Dim strYear
	Dim strMonth
	Dim strDay
		
	Err.Clear                                                                   '��: Protect system from crashing
	DbQuery = False
	
	Call LayerShowHide(1)
	With frm1	
    '---------Developer Coding part (Start)----------------------------------------------------------------
		if .rdoDurYrsFg(0).checked then DurYrsFg = "C"
    	if .rdoDurYrsFg(1).checked then DurYrsFg = "T"
    	
		Call ExtractDateFrom(frm1.fpDateTime1.Text,frm1.fpDateTime1.UserDefinedFormat,parent.gComDateType,strYear,strMonth,strDay)

		strFrDt = strYear & strMonth

		strVal	= BIZ_PGM_ID & "?txtMode="		& parent.UID_M0001						'��: �����Ͻ� ó�� ASP�� ���� 
		strVal	= strVal & "&txtAsstNo="		& Trim(.txtAsstNo.value)				'��: ��ȸ ���� ����Ÿ 
		strVal	= strVal & "&txtDepryyyymm="	& strFrdt
		strVal	= strVal & "&DurYrsFg="			& DurYrsFg

		' ���Ѱ��� �߰� 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' ����� 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' ���κμ� 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' ���κμ�(��������)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' ���� 


	'---------Developer Coding part (End)----------------------------------------------------------------
		Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
        
    End With

    DbQuery = True

 End Function
'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=======================================================================================================
Function DbQueryOk()													'��ȸ ������ ������� 
	
    '-----------------------
    'Reset variables area
    '-----------------------

    lgIntFlgMode = parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")									'This function lock the suitable field

    Call SetToolbar("11000000000111")
   	
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc"  -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><% ' ���� ���� %></TD>
	</TR>

	<!-- �Ǳ���  -->
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>

	<!-- ��������  -->
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">

		<!-- ù��° �� ����  -->
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>�ڻ��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAsstNo" SIZE=15 MAXLENGTH=18 tag="12XXXU" ALT="�ڻ��ȣ"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAssetCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:OpenMasterRef()"> <INPUT TYPE="Text" NAME="txtAsstNm" SIZE=25 MAXLENGTH=30 tag="14" ALT="�ڻ��"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>�󰢳��</TD>
									<TD CLASS="TD6" NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtDeprYyyymm" CLASS=FPDTYYYYMM tag="12X1" Title="FPDATETIME" ALT=�󰢳�� id=fpDateTime1> </OBJECT>');</SCRIPT>
									</TD>		
<!--									<TD CLASS="HIDDEN"><INPUT TYPE="RADIO" CLASS="RADIO" checked NAME="rdoDurYrsFg" TAG="12" ID="rbYrsTax" ><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDurYrsFg" TAG="12" ID="rbYrsCas"></TD>																																	
								</TR>
								<TR>		���� -->	
									<TD CLASS="TD5" NOWRAP>����������</TD>
									<TD CLASS="TD6" COLSPAN="3" NOWRAP> <SPAN STYLE="width:120;"><INPUT TYPE="RADIO" CLASS="RADIO" checked NAME="rdoDurYrsFg" TAG="12" ID="rbYrsTax"><LABEL TYPE="HIDDEN" FOR="rbYrsCas">���ȸ�����</LABEL></SPAN>
																		<SPAN STYLE="width:120;"><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoDurYrsFg" TAG="12" ID="rbYrsCas"><LABEL TYPE="HIDDEN" FOR="rbYrsTax">��������</LABEL></SPAN></TD>
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
						<TR>
							<TD HEIGHT=30% WIDTH=100%>
								<FIELDSET CLASS="CLSFLD" STYLE="HEIGHT: 100%">
								<TABLE CLASS="TB2" CELLSPACING=0 STYLE="HEIGHT: 100%">
									<TR>
										<TD CLASS="TD5" NOWRAP>�����ڵ�</TD>
										<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtAcctCd" SIZE=20 MAXLENGTH=20 tag="24XXXU" ALT="�����ڵ�"> <INPUT TYPE="Text" NAME="txtAcctNm" SIZE=20 MAXLENGTH=20 tag="24" ALT="������"></TD>
										<TD CLASS="TD5" NOWRAP>�󰢹��</TD>
										<TD CLASS="TD6" NOWRAP><SELECT NAME="cboDeprMthd" STYLE="WIDTH:150px;" tag="24"><OPTION VALUE=""></OPTION></SELECT></TD>
<!--										<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDeprMthd" SIZE=20 MAXLENGTH="10" tag="24" STYLE="TEXT-ALIGN: center" ALT="�󰢹��"></TD>  -->
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>�������</TD>
										<TD CLASS="TD6" NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtRegDt" CLASS=FPDTYYYYMMDD tag="24X1" Title="FPDATETIME" ALT=������� id=fpDateTime2> </OBJECT>');</SCRIPT>
										</TD>												
										<TD CLASS=TD5 NOWRAP>���ݾ�(�ڱ�)</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle1 name=txtLocAcqAmt style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 150px" title="FPDOUBLESINGLE" ALT="���ݾ�(�ڱ�)" tag="24X2"> </OBJECT>');</SCRIPT>&nbsp;
										</TD>											
									</TR>
									<TR>
										<TD CLASS=TD5 NOWRAP>������</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle2 name="txtAcqQty" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 80px" title="FPDOUBLESINGLE" ALT="������" tag="24X3"> </OBJECT>');</SCRIPT>&nbsp;
										</TD>
										<TD CLASS="TD5" NOWRAP>������</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 name="txtInvQty" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 80px" title="FPDOUBLESINGLE" ALT="������" tag="24X3"> </OBJECT>');</SCRIPT>&nbsp;
										</TD>										
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP>���뿬��</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle4 name="txtDurYrs" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 80px" title="FPDOUBLESINGLE" ALT="���뿬��" tag="24X3"> </OBJECT>');</SCRIPT>&nbsp;
										</TD>											
										<TD CLASS="TD5" NOWRAP>����</TD>
										<TD CLASS=TD6 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle5 name="txtDeprRate" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 80px" title="FPDOUBLESINGLE" ALT="����" tag="24X5"> </OBJECT>');</SCRIPT>&nbsp;%
										</TD>
									</TR>
								</TABLE>
								</FIELDSET>
							</TD>
						</TR>
						<TR>
							<TD HEIGHT=* WIDTH=100%>
								<FIELDSET CLASS="CLSFLD" STYLE="HEIGHT: 100%">
								<TABLE CLASS="TB2" CELLSPACING=0 STYLE="HEIGHT: 100%">
									<TR>
										<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
										<TD CLASS="TD19" NOWRAP>[���⸻ ����]</TD>
										<TD CLASS="TD19" NOWRAP>[�����]</TD>
										<TD CLASS="TD19" NOWRAP>[����߻�]</TD>
										<TD CLASS="TD19" NOWRAP>[�����]</TD>
									</TR>
									<TR>
										<TD CLASS="TD5" NOWRAP STYLE="Text-Align: right">��氡��</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle6 name="txtEndLTermAcqAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" ALT="��氡��" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>																			
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle7 name="txtFMnthAcqAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS="TD19" NOWRAP></TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle9 name="txtLMnthAcqAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>											
										</TD>																		
									</TR>
									<TR>
										<TD CLASS=TD5>�ں�������</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle10 name="txtEndLTermCptAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" ALT="�ں�������" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle11 name="txtFMnthCptAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle12 name="txtMnthCptAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle13 name="txtLMnthCptAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5>�󰢾�</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle22 name="txtEndLTermDeprAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle23 name="txtFMnthDeprAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle24 name="txtMnthDeprAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle25 name="txtLMnthDeprAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>										
									</TR>
									<TR>
										<TD CLASS=TD5>�̻󰢾�</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle26 name="txtEndLTermBalAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>									
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle27 name="txtFMnthBalAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS="TD19" NOWRAP></TD>																		
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle29 name="txtLMnthBalAmt" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5>�Ű���ⷮ</TD>
										<TD CLASS="TD19" NOWRAP></TD>
										<TD CLASS="TD19" NOWRAP></TD>											
									<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle32 name="txtMnthDisUseQty" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X2"> </OBJECT>');</SCRIPT>
										</TD>
									</TR>
									<TR>
										<TD CLASS=TD5>���</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle34 name="txtEndLTermInvQty" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X3"> </OBJECT>');</SCRIPT>
										</TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle35 name="txtFMnthInvQty" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X3"> </OBJECT>');</SCRIPT>
										</TD>
											<TD CLASS="TD19" NOWRAP></TD>
										<TD CLASS=TD19 NOWRAP>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle37 name="txtLMnthInvQty" style="HEIGHT: 20px; RIGHT: 0px; TOP: 0px; WIDTH: 130px" title="FPDOUBLESINGLE" tag="24X3"> </OBJECT>');</SCRIPT>
										</TD>
										
									</TR>
								</TABLE>
								</FIELDSET>
							</TD>
						</TR>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=10>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
