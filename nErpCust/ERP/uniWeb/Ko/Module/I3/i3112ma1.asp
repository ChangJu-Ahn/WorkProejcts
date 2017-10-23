<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Long-term Inv Analysis
'*  2. Function Name        : 
'*  3. Program ID           : I3112MA1
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Component List       : 
'*  7. Modified date(First) : 2006/05/25
'*  8. Modified date(Last)  : 2006/09/01
'*  9. Modifier (First)     : KiHong Han
'* 10. Modifier (Last)      : LEE SEUNG WOOK
'* 11. Comment
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>


Option Explicit                                                           

Const BIZ_PGM_ANALYSIS_ID = "i3112mb1.asp"
Const BIZ_PGM_JUMP1_ID = "i3111qa1"
Const BIZ_PGM_JUMP2_ID = "i3112qa1"

'=========================================================================================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgVariables.inc" -->

'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 
Dim IsOpenPop       

Dim CompanyYM

CompanyYM = UNIMonthClientFormat(UniConvDateAToB("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gAPDateFormat))


'=========================================================================================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()
    IsOpenPop = False
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()
	
	Dim	arrVar1, arrVar2, arrVar3
	
	If ReadCookie("txtPlantCd") = "" Then
		If Parent.gPlant <> "" Then
			
			frm1.txtPlantCd.value = Ucase(Parent.gPlant)
			frm1.txtPlantNm.value = Parent.gPlantNm
			
			Call CommonQueryRs(" LONGTERM_STOCK_CAL_PERIOD, PERNICIOUS_STOCK_CAL_PERIOD, PLAN_STOCK_CAL_PERIOD ", " I_LONGTERM_INV_ANAL_CONFG ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
			arrVar1 = split(lgF0, chr(11))
			arrVar2 = split(lgF1, chr(11))
			arrVar3 = split(lgF2, chr(11))
			
			frm1.txtLongtermStockCalPeriod.Value = arrVar1(0)		
			frm1.txtPerniciousStockCalPeriod.Value = arrVar2(0)
			'frm1.txtplanStockCalPeriod.Value = arrVar3(0)
			
		End If
    Else
		frm1.txtPlantCd.Value = ReadCookie("txtPlantCd")
	End If
	
	If ReadCookie("txtPlantNm") <> "" Then
		frm1.txtPlantNm.Value = ReadCookie("txtPlantNm")
	End If	
	
	If ReadCookie("txtYYYYMM") = "" Then
		frm1.txtAnalYYYYMM.Text = CompanyYM
	Else
		frm1.txtAnalYYYYMM.Text = ReadCookie("txtYYYYMM")
	End If	
	 
	WriteCookie "txtPlantCd", ""
	WriteCookie "txtPlantNm", ""
	WriteCookie "txtYYYYMM", ""
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.3 InitComboBox()  =======================================
'	Name : InitComboBox()
'	Description : Combo Display
'========================================================================================================= 
Sub InitComboBox()
	On Error Resume Next
    Err.Clear
    
End Sub

'=======================================================================================================
'   Event Name : txtPlantCd_LostFocus()
'   Event Desc : ������ ��������رⰣ, �Ǽ������رⰣ, �����ȹ�Ⱓ�� ã�´�.
'=======================================================================================================
'Sub txtPlantCd_LostFocus()
Sub txtPlantCd_onchange()
    Dim strYear
    Dim strMonth
    Dim strDay
	Dim	arrVar1, arrVar2, arrVar3
	
	If frm1.txtPlantCd.value <> "" Then
		Call CommonQueryRs(" LONGTERM_STOCK_CAL_PERIOD, PERNICIOUS_STOCK_CAL_PERIOD, PLAN_STOCK_CAL_PERIOD ", " I_LONGTERM_INV_ANAL_CONFG ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
		
		
		arrVar1 = split(lgF0, chr(11))
		arrVar2 = split(lgF1, chr(11))
		arrVar3 = split(lgF2, chr(11))
		
		frm1.txtLongtermStockCalPeriod.Value = arrVar1(0)		
		frm1.txtPerniciousStockCalPeriod.Value = arrVar2(0)
		'frm1.txtplanStockCalPeriod.Value = arrVar3(0)
	Else
		frm1.txtPlantNm.Value  = ""
		frm1.txtLongtermStockCalPeriod.Value = ""
		frm1.txtPerniciousStockCalPeriod.Value = ""
		'frm1.txtplanStockCalPeriod.Value = ""
	End If
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
'	Name : OpenPlant()
'	Description : Plant PopUp
'--------------------------------------------------------------------------------------------------------- 
Sub OpenPlant() 
	Dim arrRet, arr_var
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim	arrVar1, arrVar2, arrVar3
	
	If IsOpenPop = True Then Exit Sub

	IsOpenPop = True

	arrParam(0) = "�����˾�"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "����"			
	
	arrField(0) = "PLANT_CD"	
	arrField(1) = "PLANT_NM"	
	
	arrHeader(0) = "����"		
	arrHeader(1) = "�����"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) <> "" Then
		frm1.txtPlantCd.Value = arrRet(0)		
		frm1.txtPlantNm.Value = arrRet(1)
		
		Call CommonQueryRs(" LONGTERM_STOCK_CAL_PERIOD, PERNICIOUS_STOCK_CAL_PERIOD, PLAN_STOCK_CAL_PERIOD ", " I_LONGTERM_INV_ANAL_CONFG ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)	
		
		arrVar1 = split(lgF0, chr(11))
		arrVar2 = split(lgF1, chr(11))
		arrVar3 = split(lgF2, chr(11))
			
		frm1.txtLongtermStockCalPeriod.Value = arrVar1(0)		
		frm1.txtPerniciousStockCalPeriod.Value = arrVar2(0)
		'frm1.txtplanStockCalPeriod.Value = arrVar3(0)
	End If	

	frm1.txtPlantCd.Focus	
	Set gActiveElement = document.activeElement
End Sub

'=============================================  2.5.2 JumpToLongtermInvList()  ======================================
'=	Event Name : JumpToLongtermInvList
'=	Event Desc : ��������Ȳ���� Jump
'========================================================================================================
Function JumpToLongtermInvList()
	With frm1
		'�����ڵ�/��/�м����� 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtYYYYMM", .txtAnalYYYYMM.Text
	End With
	
	PgmJump(BIZ_PGM_JUMP1_ID)
End Function

'=============================================  2.5.2 JumpToLongtermInvChange()  ======================================
'=	Event Name : JumpToLongtermInvChange
'=	Event Desc : ���������̷� Jump
'========================================================================================================
Function JumpToLongtermInvChange()
	With frm1
		'�����ڵ�/��/�м����� 
		WriteCookie "txtPlantCd", Trim(.txtPlantCd.value)
		WriteCookie "txtPlantNm", Trim(.txtPlantNm.value)
		WriteCookie "txtYYYYMM", .txtAnalYYYYMM.Text
	End With
	
	PgmJump(BIZ_PGM_JUMP2_ID)
End Function

'=========================================================================================================
' Name : Analysis()   
' Description : ������м� Function          
'========================================================================================================= 
Function Analysis()
    
    Analysis = False
	Dim IntRetCD
	
	'-----------------------
	'Check condition area
	'----------------------- 
	If Not chkField(Document, "1") Then
		Exit Function
	End If
	
	'-----------------------
	'Save function call area
	'-----------------------
	IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO,"X","X")
	If IntRetCD = vbNo Then Exit Function
	
	If DbAnalysis("N") = False then Exit Function	
	
    Analysis = True 
End Function

'========================================================================================
' Function Name : DbAnalysis
' Function Desc : 
'========================================================================================
Function DbAnalysis(Byval pvReAnalFlag) 
	DbAnalysis = False 
	
	err.Clear 
	On Error Resume Next    
	
	Call LayerShowHide(1)
	
	Dim strVal
	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate

	Call ExtractDateFrom(frm1.txtAnalYYYYMM.Text,frm1.txtAnalYYYYMM.UserDefinedFormat,Parent.gComDateType,strYear,strMonth,strDay)
	
	With frm1
		
		strVal = BIZ_PGM_ANALYSIS_ID & "?txtReAnalFlag="	& pvReAnalFlag _
										& "&txtAnalYYYYMM=" & strYear&strMonth _
										& "&txtPlantCd="	& UCASE(.txtPlantCd.value) _
										& "&txtLongterm="	& UCASE(.txtLongtermStockCalPeriod.value) _
										& "&txtPernicious=" & UCASE(.txtPerniciousStockCalPeriod.value)
										
	
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbAnalysis = True
End Function

'========================================================================================
' Function Name : DbAnalysisOk
' Function Desc : 
'========================================================================================
Function DbAnalysisOk() 
	DbAnalysisOk = False
	Call DisplayMsgBox("990000","X", "X", "X")
    DbAnalysisOk = True
End Function

'========================================================================================
' Function Name : DbReAnalysis
' Function Desc : 
'========================================================================================
Function DbReAnalysis() 
	DbReAnalysis = False
	Dim IntRetCD
	
	IntRetCD = DisplayMsgBox("U00002", Parent.VB_YES_NO, "X", "X") 
	If IntRetCD = vbNo Then	Exit Function
	
	If DbAnalysis("Y") = False then Exit Function
	
    DbReAnalysis = True
End Function


'=========================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()
	Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")  
    
    Call ggoOper.FormatDate(frm1.txtAnalYYYYMM, Parent.gDateFormat, 2)
    Call SetDefaultVal
    
    Call SetToolBar("10000000000011")
    Call InitVariables    
    Call InitComboBox
    
    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantCd.focus 
    Else
		frm1.txtAnalYYYYMM.focus
	End If
    
End Sub

'=========================================================================================================
'   Event Name : txtAnalYYYYMM_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=========================================================================================================
Sub txtAnalYYYYMM_DblClick(Button) 
	If Button = 1 Then 
		frm1.txtAnalYYYYMM.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtAnalYYYYMM.Focus
	End If 
End Sub

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
    Call parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE, False)
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB1" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD HEIGHT=5>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������м�</font></td>
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
		<TD WIDTH=100% CLASS="Tab11" valign=top>
			<TABLE WIDTH="100%" HEIGHT="100%" CELLSPACING=0 CELLPADDING=0>				
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
									<TD CLASS="TD6" NOWRAP>
										<INPUT TYPE=TEXT NAME="txtPlantCd" SIZE="10" MAXLENGTH="4" ALT="����" TAG="12XXXU"><IMG ALIGN=top HEIGHT=20 NAME=btnPlantPopup ONCLICK=vbscript:OpenPlant() SRC="../../../CShared/image/btnPopup.gif" WIDTH=16  TYPE="BUTTON">&nbsp;<INPUT NAME="txtPlantNm" TAG="14X">
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�м�����</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> ID="txtAnalYYYYMM" NAME="txtAnalYYYYMM" CLASS="FPDTYYYYMM" title=FPDATETIME ALT="�м�����" TAG="12"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLongtermStockCalPeriod" SIZE="4" MAXLENGTH="4" ALT="������" TAG="12XXXU">&nbsp;���� �̻� ���(���)������ ���� ǰ��</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>�Ǽ����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPerniciousStockCalPeriod" SIZE="4" MAXLENGTH="4" ALT="�Ǽ����" TAG="12XXXU">&nbsp;���� �̻� ���(���)������ ���� ǰ��</TD>
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
					<TD><BUTTON NAME="btnSummary" CLASS="CLSMBTN" Flag=1 onclick="Analysis()">������м�</BUTTON></TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:JumpToLongtermInvList">��������Ȳ</A>&nbsp;|&nbsp;<A href="vbscript:JumpToLongtermInvChange">����������</A></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT= <%=BizSize%>><IFRAME NAME="MyBizASP" tabindex=-1 SRC="../../blank.htm" WIDTH=100% HEIGHT= <%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

