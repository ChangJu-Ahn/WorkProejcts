<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : Offer Sheet
'*  5. Program Desc         : Offer Sheet
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/06/29
'*  8. Modified date(Last)  : 2001/10
'*  9. Modifier (First)     : shin Jin Hyun
'* 10. Modifier (Last)      : Ma Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- '******************************************  1.1 Inc ����   **************************************** -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'��: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim StartDate, EndDate

	StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
	StartDate = UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat) 
	EndDate   = UniConvDateAToB("<%=GetSvrDate%>"  ,Parent.gServerDateFormat,Parent.gDateFormat)

Dim lblnWinEvent
Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size         
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtFrDt.Text	= StartDate
	frm1.txtToDt.Text	= EndDate
	
	frm1.txtPlantCd.value=parent.gPlant
	frm1.txtPlantNm.value=parent.gPlantNm
End Sub

'==========================================  LoadInfTB19029()  ================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
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
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.value= arrRet(1)		
		lgBlnFlgChgValue = True
		frm1.txtPlantCd.focus
	End If	
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenSupplier()  -------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"			
	arrParam(1) = "B_Biz_Partner"			
	
	arrParam(2) = Trim(frm1.txtSupplierCd.Value)	
	arrParam(3) = ""
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "����ó"			
	
	arrField(0) = "BP_Cd"				
	arrField(1) = "BP_NM"				

	arrHeader(0) = "����ó"				
	arrHeader(1) = "����ó��"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		lgBlnFlgChgValue = True
		frm1.txtSupplierCd.focus
	End If
	Set gActiveElement = document.activeElement
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	Call LoadInfTB19029                                                     '��: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitVariables                                                      '��: Initializes local global variables
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'��: ��ư ���� ���� 
    frm1.txtPlantCd.focus 
    Set gActiveElement = document.activeElement
End Sub

'=========================================  Form_QueryUnload()  =====================================
Sub Form_QueryUnload(Cancel, UnloadMode)
   
End Sub
'==========================================================================================
'   Event Name : txtFrDt  	 
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrDt.focus
	End if
End Sub
'========================================================================================
'   Event Name : txtToDt  	 
'==========================================================================================

Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.focus
	End if
End Sub

'===============================  FncPrint()  ============================================
Function FncPrint() 
	Call parent.FncPrint()
End Function
'===============================  FncFind()  ============================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                                     <%'��:ȭ�� ����, Tab ���� %>
End Function

'==========================================  2.2.6 ChkKeyField()  =======================================
 Function ChkKeyField()
	
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	
	ChkKeyField = true
	
	strWhere = " PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "
	
	Call CommonQueryRs(" PLANT_NM "," B_PLANT ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("17a003","X","����","X")
		frm1.txtPlantNm.value = ""
		ChkKeyField = False
		Exit Function
	End If
	
	strDataNm = split(lgF0,chr(11))
	
	frm1.txtPlantNm.value = strDataNm(0)
	
End Function

'==========================================  2.2.6 ChkKeyField1()  =======================================
 Function ChkKeyField1()
	
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	
	ChkKeyField1 = true
	
	strWhere = " BP_CD =  " & FilterVar(frm1.txtSupplierCd.value, "''", "S") & "  "
	
	Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("17a003","X","����ó","X")
		frm1.txtSupplierNM.value = ""
		ChkKeyField1 = False
		Exit Function
	End If
	
	strDataNm = split(lgF0,chr(11))
	
	frm1.txtSupplierNM.value = strDataNm(0)
	
End Function
'===============================  FncBtnPrint()  ============================================
Function FncBtnPrint() 
	Dim StrUrl
	dim var1,var2,var3,var4
    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If
	
	 IF ChkKeyField() = False Then 
		frm1.txtPlantCd.focus
		Exit Function
    End if
    
    IF ChkKeyField1() = False Then 
		frm1.txtSupplierCd.focus
		Exit Function
    End if
    
	On Error Resume Next                                                    '��: Protect system from crashing
	

	var1 = UCase(frm1.txtPlantCd.value)
	var2 = UCase(frm1.txtSupplierCd.value)
	var3 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,parent.gDateFormat,parent.gServerDateType)  'uniCdate(frm1.txtFrDt.Text)
	var4 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,parent.gDateFormat,parent.gServerDateType)  'uniCdate(frm1.txtToDt.Text)

	with frm1
	    If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then	
            Call DisplayMsgBox("17a003","X","��꼭����","X")				      
            Exit Function
        End if   
	End with
	        		
	strUrl = strUrl & "plant|" & var1 & "|bp_cd|" & var2 & "|fr_dt|" & var3 & "|to_dt|" & var4 
	
'----------------------------------------------------------------
' Print �Լ����� ȣ�� 
'----------------------------------------------------------------
	ObjName = AskEBDocumentName("m5111oa1","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)
'----------------------------------------------------------------	
End Function

'===============================  BtnPreview()  ============================================
Function BtnPreview() 
    
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If

	IF ChkKeyField() = False Then 
		frm1.txtPlantCd.focus
		Exit Function
    End if
    
   IF ChkKeyField1() = False Then 
		frm1.txtSupplierCd.focus
		Exit Function
    End if
    
	Dim var1,var2,var3,var4
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader
	
	var1 = UCase(frm1.txtPlantCd.value)
	var2 = UCase(frm1.txtSupplierCd.value)
	var3 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,parent.gDateFormat,parent.gServerDateType) 'uniCdate(frm1.txtFrDt.Text)
	var4 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,parent.gDateFormat,parent.gServerDateType) 'uniCdate(frm1.txtToDt.Text)
	
	with frm1
	    If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then	
            Call DisplayMsgBox("17a003","X","��꼭����","X")				      
            Exit Function
        End if 	   
	End with
	

	strUrl = strUrl & "plant|" & var1 & "|bp_cd|" & var2 & "|fr_dt|" & var3 & "|to_dt|" & var4 

	ObjName = AskEBDocumentName("m5111oa1","ebr")
	Call FncEBRPreview(ObjName, strUrl)	
End Function

'===============================  FncExit()  ============================================
Function FncExit()
	   
    FncExit = True
    
End Function

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
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>���Űŷ�����</font></td>
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
					<TD HEIGHT=20 WIDTH=100%>						
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="����" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 ALT="����" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSupplierCd" SIZE=10 MAXLENGTH=10 ALT="����ó" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()">
													   <INPUT TYPE=TEXT NAME="txtSupplierNm" SIZE=20 MAXLENGTH=18 ALT="����ó" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>��꼭����</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m5111oa1_fpDateTime1_txtFrDt.js'></script> ~
													   <script language =javascript src='./js/m5111oa1_fpDateTime2_txtToDt.js'></script></TD>
							</TR>
						</TABLE>
						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD Width = 10>&nbsp</TD>
					<TD Valign=top>				
					    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;		    
					    <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON>		
					</TD>
					<TD Width = 10>&nbsp</TD>
				</TR>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SRC="m3112mb1.asp" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST" >
    <input type="hidden" name="uname">
    <input type="hidden" name="dbname">
    <input type="hidden" name="filename">
    <input type="hidden" name="condvar">
	<input type="hidden" name="date">
</FORM>
</BODY>
</HTML>
