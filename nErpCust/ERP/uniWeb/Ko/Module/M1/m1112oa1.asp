<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M1112OA1
'*  4. Program Name         : ����ó���ܰ���� 
'*  5. Program Desc         : ����ó���ܰ���� 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/06/29
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : shin Jin Hyun
'* 10. Modifier (Last)      : Kang Su Hwan
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
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit			
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================  1.2.1 Global ��� ����  ======================================
Dim StartDate
	StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
	StartDate = UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat) 

Dim lgIsOpenPop
Dim lblnWinEvent
Dim IsOpenPop
'=========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE   
    lgBlnFlgChgValue = False    
    lgIntGrpCount = 0           
End Sub
'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtValidFrDt.Text	= StartDate 'UNIFormatDate("<%= StartDate %>")
	frm1.txtPlantCd.value=parent.gPlant
	frm1.txtPlantNm.value=parent.gPlantNm
End Sub
'===================================  LoadInfTB19029()  =======================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'------------------------------------------  OpenPlantCd()  -------------------------------------------------
Function OpenPlantCd()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����"						
	arrParam(1) = "B_PLANT"      					
	arrParam(2) = Trim(frm1.txtPlantCd.Value)		
'	arrParam(3) = Trim(frm1.txtPlantNm.Value)		
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
		frm1.txtPlantCd.Value = arrRet(0)
		frm1.txtPlantNm.Value = arrRet(1)
		frm1.txtPlantCd.focus
	End If	
	frm1.txtItemCd1.value=""
	frm1.txtItemNm1.value=""
	frm1.txtItemCd2.value=""
	frm1.txtItemNm2.value=""
End Function
'------------------------------------------  OpenItemCd()  -------------------------------------------------

Function OpenItemCd1()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd1.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)

	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd1.focus
		Exit Function
	Else
		frm1.txtItemCd1.Value	= arrRet(0)
		frm1.txtItemNm1.Value	= arrRet(1)
		frm1.txtItemCd1.focus
	End If
End Function

'------------------------------------------  OpenItemCd()  -------------------------------------------------
Function OpenItemCd2()
	Dim arrRet
	Dim arrParam(5), arrField(1)
	Dim iCalledAspName
	Dim IntRetCD

	If lgIsOpenPop = True Then Exit Function

	if  Trim(frm1.txtPlantCd.Value) = "" then
		Call DisplayMsgBox("17A002", "X", "����", "X")
		frm1.txtPlantCd.focus
		Exit Function
	End if

	lgIsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd2.value)		' Item Code
	arrParam(2) = "!"	' ��12!MO"�� ���� -- ǰ����� ������ ���ޱ��� 
	arrParam(3) = "30!P"
	arrParam(4) = ""		'-- ��¥ 
	arrParam(5) = ""		'-- ����(b_item_by_plant a, b_item b: and ���� ����)

	arrField(0) = 1 ' -- ǰ���ڵ� 
	arrField(1) = 2 ' -- ǰ��� 

	iCalledAspName = AskPRAspName("B1B11PA3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		lgIsOpenPop = False
		Exit Function
	End If

	arrRet = window.showModalDialog(iCalledAspName, Array(parent.window, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd2.focus
		Exit Function
	Else
		frm1.txtItemCd2.Value	= arrRet(0)
		frm1.txtItemNm2.Value	= arrRet(1)
		frm1.txtItemCd2.focus
	End If
End Function
'------------------------------------------  OpenBpCd()  -------------------------------------------------
Function OpenBpCd1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBpCd1.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"						
	arrParam(1) = "B_Biz_Partner"					
	arrParam(2) = Trim(frm1.txtBpCd1.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)			
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "����ó"						
	
    arrField(0) = "BP_CD"							
    arrField(1) = "BP_NM"							
    
    arrHeader(0) = "����ó"						
    arrHeader(1) = "����ó��"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd1.focus
		Exit Function
	Else
		frm1.txtBpCd1.Value = arrRet(0)
		frm1.txtBpNm1.Value = arrRet(1)
		frm1.txtBpCd1.focus
	End If		
End Function

'------------------------------------------  OpenBpCd()  -------------------------------------------------
Function OpenBpCd2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtBpCd2.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"						
	arrParam(1) = "B_Biz_Partner"					
	arrParam(2) = Trim(frm1.txtBpCd2.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)			
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "					
	arrParam(5) = "����ó"						
	
    arrField(0) = "BP_CD"							
    arrField(1) = "BP_NM"							
    
    arrHeader(0) = "����ó"						
    arrHeader(1) = "����ó��"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBpCd2.focus
		Exit Function
	Else
		frm1.txtBpCd2.Value = arrRet(0)
		frm1.txtBpNm2.Value = arrRet(1)
		frm1.txtBpCd2.focus
	End If		
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
		frm1.txtPlantCd.value = ""
		ChkKeyField = False
		Exit Function
	End If
	
	strDataNm = split(lgF0,chr(11))
	
	frm1.txtPlantNm.value = strDataNm(0)
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029         
    Call ggoOper.LockField(Document, "N")        
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitVariables                                                
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")								

    frm1.txtValidFrDt.focus 
	Set gActiveElement = document.activeElement
End Sub

'=====================================  txtValidFrDt_DblClick()  =====================================
Sub txtValidFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtValidFrDt.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtValidFrDt.Focus
	End If
End Sub

'=====================================  FncPrint()  =====================================
 Function FncPrint() 
	Call parent.FncPrint()
End Function

'=====================================  FncFind()  =====================================
 Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                           
End Function
'=====================================  FncBtnPrint()  =====================================
 Function FncBtnPrint() 
	Dim StrUrl
	dim var1,var2,var3,var4,var5,var6
    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
    IF ChkKeyField() = False Then 
		frm1.txtPlantCd.focus
		Exit Function
    End if
    
    with frm1
		if (frm1.txtItemCd1.value <> "") AND  (frm1.txtItemCd2.value <> "") then
			if  UCase(frm1.txtItemCd1.value) > UCase(frm1.txtItemCd2.value)  then	
				Call DisplayMsgBox("17a003","X","ǰ��","X")
				frm1.txtItemCd1.focus 
				Exit Function
			End if  
		End if 
		
		if (frm1.txtBpCd1.value <> "") AND  (frm1.txtBpCd2.value <> "") then
			if  UCase(frm1.txtBpCd1.value) > UCase(frm1.txtBpCd2.value)  then	
				Call DisplayMsgBox("17a003","X","����ó","X")
				frm1.txtBpCd1.focus 
				Exit Function
			End if   
		End if 
	End with
	
	On Error Resume Next                                            
	var1 = UniConvDateToYYYYMMDD(frm1.txtValidFrDt.Text,parent.gDateFormat,parent.gServerDateType)'uniCdate(frm1.txtValidFrDt.text)
	var2 = FilterVar(Trim(UCase(frm1.txtPlantCd.value)), "" ,  "SNM") 
	var3 = FilterVar(Trim(UCase(frm1.txtItemCd1.value)), "" ,  "SNM") 
	var4 = FilterVar(Trim(UCase(frm1.txtItemCd2.value)), "" ,  "SNM")
	var5 = FilterVar(Trim(UCase(frm1.txtBpCd1.value)), "" ,  "SNM") 
	var6 = FilterVar(Trim(UCase(frm1.txtBpCd2.value)), "" ,  "SNM")  
	
	'if	var3="" then
	'	var3="%"
	'end if
	'if	var5="" then
	'	var5="%"
	'end if
	if	var4="" then 
		var4="ZZZZZZZZZZZZZZZZZZ"
	end if
	
	if	var6="" then 
		var6="ZZZZZZZZZZ"
	end if
        		
	strUrl = strUrl & "fr_dt|" & var1 
	strUrl = strUrl & "|plant_cd|" & var2
	strUrl = strUrl & "|fr_item|" & var3
	strUrl = strUrl & "|to_item|" & var4
	strUrl = strUrl & "|fr_bp|" & var5
	strUrl = strUrl & "|to_bp|" & var6
	
'----------------------------------------------------------------
' Print �Լ����� ȣ�� 
'----------------------------------------------------------------
	ObjName = AskEBDocumentName("m1112oa1","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)
'----------------------------------------------------------------	
End Function

'=====================================  BtnPreview()  =====================================
Function BtnPreview() 
    If Not chkField(Document, "1") Then								
       Exit Function
    End If
	
	IF ChkKeyField() = False Then 
		frm1.txtPlantCd.focus
		Exit Function
    End if
    
             
	with frm1
		if (frm1.txtItemCd1.value <> "") AND  (frm1.txtItemCd2.value <> "") then
			if  UCase(frm1.txtItemCd1.value) > UCase(frm1.txtItemCd2.value)  then	
				Call DisplayMsgBox("17a003","X","ǰ��","X")
				frm1.txtItemCd1.focus 
				Exit Function
			End if  
		End if 
		
		if (frm1.txtBpCd1.value <> "") AND  (frm1.txtBpCd2.value <> "") then
			if  UCase(frm1.txtBpCd1.value) > UCase(frm1.txtBpCd2.value)  then	
				Call DisplayMsgBox("17a003","X","����ó","X")
				frm1.txtBpCd1.focus 
				Exit Function
			End if   
		End if 
	End with
	
	dim var1,var2,var3,var4,var5,var6
	dim strUrl
	dim arrParam, arrField, arrHeader
	
	var1 = UniConvDateToYYYYMMDD(frm1.txtValidFrDt.Text,parent.gDateFormat,parent.gServerDateType)'uniCdate(frm1.txtValidFrDt.text)
	var2 = FilterVar(Trim(UCase(frm1.txtPlantCd.value)), "" ,  "SNM") 
	var3 = FilterVar(Trim(UCase(frm1.txtItemCd1.value)), "" ,  "SNM") 
	var4 = FilterVar(Trim(UCase(frm1.txtItemCd2.value)), "" ,  "SNM")
	var5 = FilterVar(Trim(UCase(frm1.txtBpCd1.value)), "" ,  "SNM") 
	var6 = FilterVar(Trim(UCase(frm1.txtBpCd2.value)), "" ,  "SNM")  
	
	'if	var3="" then
	'	var3="%"
	'end if
	'if	var5="" then
	'	var5="%"
	'end if
	if	var4="" then 
		var4="ZZZZZZZZZZZZZZZZZZ"
	end if
	
	if	var6="" then 
		var6="ZZZZZZZZZZ"
	end if
	
	strUrl = strUrl & "fr_dt|" & var1 
	strUrl = strUrl & "|plant_cd|" & var2
	strUrl = strUrl & "|fr_item|" & var3
	strUrl = strUrl & "|to_item|" & var4
	strUrl = strUrl & "|fr_bp|" & var5
	strUrl = strUrl & "|to_bp|" & var6	
	
	ObjName = AskEBDocumentName("m1112oa1","ebr")
	Call FncEBRPreview(ObjName, strUrl)
End Function

'=====================================  FncExit()  =====================================
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
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ó���ܰ���Ȳ</font></td>
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
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
	    		<TR>
					<TD WIDTH=100%>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>���������</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME style="LEFT: 0px; WIDTH: 100px; TOP: 0px; HEIGHT: 20px" name=txtValidFrDt CLASSID=<%=gCLSIDFPDT%> tag="12X1" ALT="���������"></OBJECT>');</SCRIPT>										
								</TD>																												  
								<TD CLASS="TD6" NOWRAP>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>����</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="����"   NAME="txtPlantCd" SIZE=10 LANG="ko" MAXLENGTH=4 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPlantCd()">
													   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14"></TD>													   
								<TD CLASS="TD6" NOWRAP></TD>
																	
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>ǰ��</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" MAXLENGTH=18   SIZE=10 MAXLENGTH=10 ALT="ǰ��" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd1()">
													   <INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=20 ALT="ǰ��" tag="14"></TD>
								<TD CLASS="TD6" NOWRAP>~ <INPUT TYPE=TEXT NAME="txtItemCd2" MAXLENGTH=18   SIZE=10 MAXLENGTH=10 ALT="ǰ��" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd2()">
													   <INPUT TYPE=TEXT NAME="txtItemNm2" SIZE=20 ALT="ǰ��" tag="14"></TD>					   
							</TR>							
							<TR>
								<TD CLASS="TD5" NOWRAP>����ó</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd1"   SIZE=10 MAXLENGTH=10 ALT="����ó" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBpCd1()">
													   <INPUT TYPE=TEXT NAME="txtBpNm1" SIZE=20 MAXLENGTH=18 ALT="����ó" tag="14"></TD>
								<TD CLASS="TD6" NOWRAP>~ <INPUT TYPE=TEXT NAME="txtBpCd2"  SIZE=10 MAXLENGTH=10 ALT="����ó" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBpCd2()">
													   <INPUT TYPE=TEXT NAME="txtBpNm2" SIZE=20 MAXLENGTH=18 ALT="����ó" tag="14"></TD>
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>�̸�����</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 tabindex = -1></IFRAME>
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
