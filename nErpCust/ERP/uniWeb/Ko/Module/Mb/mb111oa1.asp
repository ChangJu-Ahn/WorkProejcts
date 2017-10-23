<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : mb111oa1
'*  4. Program Name         : ��� �ҿ䷮ ��� 
'*  5. Program Desc         : ��� �ҿ䷮ ��� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/07/01
'*  8. Modified date(Last)  : 2003/07/01
'*  9. Modifier (First)     : KANG SU HWAN
'* 10. Modifier (Last)      : KANG SU HWAN
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit	
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.3 Global Variable�� ����  ===============================
'=========================================================================================================
Dim lblnWinEvent
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE 
    lgBlnFlgChgValue = False  
    lgIntGrpCount = 0         
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	Dim StartDate, EndDate
	
	StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", parent.gServerDateFormat)
    StartDate = UniConvDateAToB(StartDate, parent.gServerDateFormat, parent.gDateFormat)
    EndDate   = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
    
	frm1.txtFrDt.Text	= StartDate
	frm1.txtToDt.Text	= EndDate
	
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","QA") %>
End Sub

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'=========================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����"	
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
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value=arrRet(0)		
		frm1.txtPlantNm.value=arrret(1)
		frm1.txtPlantCd.focus
	End If	
End Function

'================================================================================================================================
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSpplCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"				
	arrParam(1) = "B_BIZ_PARTNER"			

	arrParam(2) = Trim(frm1.txtSpplCd.Value)
	'arrParam(3) = Trim(frm1.txtSpplNm.Value)	
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "����ó"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "����ó"				
    arrHeader(1) = "����ó��"			
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSpplCd.Value = arrRet(0)
		frm1.txtSpplNm.Value = arrRet(1)
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'*********************************************************************************************************
 Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFrDt.Focus
	End if
End Sub

 Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToDt.Focus
	End if
End Sub
'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029                                                     
    Call ggoOper.LockField(Document, "N")   
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitVariables                                                      
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										
    
    frm1.txtPlantCd.focus 
	Set gActiveElement = document.activeElement
End Sub

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
 Function FncPrint() 
	Call parent.FncPrint()
End Function
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
 Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)  
End Function

'==========================================  2.2.6 ChkKeyField()  =======================================
'	Name : ChkKeyField()
'	Description : 
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	
	ChkKeyField = true
	
	'����	
	strWhere = " PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "
	Call CommonQueryRs(" PLANT_NM "," B_PLANT ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("17a003","X","����","X")
		frm1.txtPlantCd.focus
		frm1.txtPlantNm.value = ""
		ChkKeyField = False
		Exit Function
	End If
	
	strDataNm = split(lgF0,chr(11))
	frm1.txtPlantNm.value = strDataNm(0)

	'����ó	
	If Trim(frm1.txtSpplCd.value) <> "" Then
		strWhere = " BP_CD =  " & FilterVar(frm1.txtSpplCd.value, "''", "S") & "  "
		Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","����ó","X")
			frm1.txtSpplCd.focus
			frm1.txtSpplNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
		
		strDataNm = split(lgF0,chr(11))
		frm1.txtSpplNm.value = strDataNm(0)
	End If
End Function


'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
 Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
	dim var1,var2,var3,var4,var5,var6,var7
    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
    IF ChkKeyField() = False Then 
		frm1.txtPlantCd.focus
		Exit Function
    End if
    
    with frm1
        If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","��������", "X")
			Exit Function
		End if   
	End with

	On Error Resume Next                   
	
	lngPos = 0
	
	var1 = UCase(frm1.txtPlantCd.value)
	var2 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,Parent.gDateFormat,Parent.gServerDateType) 'uniCdate(frm1.txtFrDt.text)
	var3 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,Parent.gDateFormat,Parent.gServerDateType)'uniCdate(frm1.txtToDt.text)
	
	If Trim(frm1.txtSpplCd.value) = "" Then
		var4 = ""
		var5 = "ZZZZZZZZZZ"
	Else
		var4 = UCase(frm1.txtSpplCd.value)
		var5 = UCase(frm1.txtSpplCd.value)
	End If
	
	If frm1.rdoflg1.checked = True Then
		var6 = ""
		var7 = "Z"
	ElseIf frm1.rdoflg2.checked = True Then
		var6 = "C"
		var7 = "C"
	ElseIf frm1.rdoflg3.checked = True Then
		var6 = "F"
		var7 = "F"
	End If
        		
	strUrl = strUrl & "plant|" 			& var1	
	strUrl = strUrl & "|fr_podt|" 		& var2
	strUrl = strUrl & "|to_podt|"		& var3
	strUrl = strUrl & "|fr_spplcd|" 	& var4
	strUrl = strUrl & "|to_spplcd|" 	& var5
	strUrl = strUrl & "|fr_sppltype|"	& var6
	strUrl = strUrl & "|to_sppltype|"	& var7
	
'----------------------------------------------------------------
' Print �Լ����� ȣ�� 
'----------------------------------------------------------------
	
	ObjName = AskEBDocumentName("mb111oa1","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)
'----------------------------------------------------------------
	
	Call BtnDisabled(0)	
		
End Function

'========================================================================================
' Function Name : BtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================
Function BtnPreview() 
	On Error Resume Next                      
    
    If Not chkField(Document, "1") Then		
       Exit Function
    End If
    
    IF ChkKeyField() = False Then 
		frm1.txtPlantCd.focus
		Exit Function
    End if
    
    With frm1
        If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","��������", "X")
			Exit Function
		End if   
	End With

	dim var1,var2,var3,var4,var5,var6,var7
	dim strUrl
	dim arrParam, arrField, arrHeader
		
	var1 = UCase(frm1.txtPlantCd.value)
	var2 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,Parent.gDateFormat,Parent.gServerDateType) 'uniCdate(frm1.txtFrDt.text)
	var3 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,Parent.gDateFormat,Parent.gServerDateType)'uniCdate(frm1.txtToDt.text)
	
	If Trim(frm1.txtSpplCd.value) = "" Then
		var4 = ""
		var5 = "ZZZZZZZZZZ"
	Else
		var4 = UCase(frm1.txtSpplCd.value)
		var5 = UCase(frm1.txtSpplCd.value)
	End If
	
	If frm1.rdoflg1.checked = True Then
		var6 = ""
		var7 = "Z"
	ElseIf frm1.rdoflg2.checked = True Then
		var6 = "C"
		var7 = "C"
	ElseIf frm1.rdoflg3.checked = True Then
		var6 = "F"
		var7 = "F"
	End If
        		
	strUrl = strUrl & "plant|" 			& var1	
	strUrl = strUrl & "|fr_podt|" 		& var2
	strUrl = strUrl & "|to_podt|"		& var3
	strUrl = strUrl & "|fr_spplcd|" 	& var4
	strUrl = strUrl & "|to_spplcd|" 	& var5
	strUrl = strUrl & "|fr_sppltype|"	& var6
	strUrl = strUrl & "|to_sppltype|"	& var7

	ObjName = AskEBDocumentName("mb111oa1","ebr")
	Call FncEBRPreview(ObjName, strUrl)
	
	Call BtnDisabled(0)	
		
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
    FncExit = True
End Function

'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

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
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../SChared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��� �ҿ䷮ ���</font></td>
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
		<TD WIDTH=100% CLASS="Tab11" HEIGHT=* colspan="2">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT ALT="����" NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 tag="14X"></TD>
							</TR>
							<TR><TD CLASS="TD5" NOWRAP>��������</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<script language =javascript src='./js/mb111oa1_fpDateTime2_txtFrDt.js'></script>
											</td>
											<td>~</td>
											<td>
												<script language =javascript src='./js/mb111oa1_fpDateTime2_txtToDt.js'></script>
											</td>
										<tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSpplCd"   SIZE=10 MAXLENGTH=10 ALT="����ó" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSupplierCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
													   <INPUT TYPE=TEXT NAME="txtSpplNm" SIZE=20 MAXLENGTH=18 ALT="����ó" tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoflg" id = "rdoflg1" Value="A"  checked tag="12"><label for="rdoflg1">&nbsp;��ü&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoflg" id = "rdoflg2" Value="C"  tag="12"><label for="rdoflg2">&nbsp;����&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="����" NAME="rdoflg" id = "rdoflg3" Value="F"  tag="12"><label for="rdoflg3">&nbsp;����&nbsp;</label>
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;
						<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<!-- Print Program must contain this HTML Code -->
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
<!-- End of Print HTML Code -->
</BODY>
</HTML>
