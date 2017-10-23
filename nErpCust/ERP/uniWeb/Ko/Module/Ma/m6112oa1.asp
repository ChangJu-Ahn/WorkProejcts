<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m6112oa1
'*  4. Program Name         : 부대비배부내역출력 
'*  5. Program Desc         : 부대비배부내역출력(발주번호별, 품목별)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/03/08
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Lee JeongTae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
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

Dim StartDate, EndDate
	StartDate = uniDateAdd("m", -1, "<%=GetSvrDate%>", Parent.gServerDateFormat)
	StartDate = UniConvDateAToB(StartDate,Parent.gServerDateFormat,Parent.gDateFormat) 
	EndDate   = UniConvDateAToB("<%=GetSvrDate%>"  ,Parent.gServerDateFormat,Parent.gDateFormat)

Dim lblnWinEvent
Dim IsOpenPop
'==============================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0        
End Sub
'==============================================================================================================================
Sub SetDefaultVal()
    frm1.txtBizAreaCd.value=parent.gBizArea
	frm1.txtFrDt.Text	= StartDate
	frm1.txtToDt.Text	= EndDate
End Sub
'==============================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","OA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub
'==============================================================================================================================
Function OpenBizAreaCd()
	
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "사업장"						
	arrParam(1) = "B_Biz_Area"					
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)			
	'arrParam(4) = "BP_TYPE In ('S','CS') And usage_flag='Y'"					
	arrParam(5) = "사업장"						
	
    arrField(0) = "Biz_Area_Cd"							
    arrField(1) = "Biz_Area_Nm"							
    
    arrHeader(0) = "사업장"						
    arrHeader(1) = "사업장명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus	
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtBizAreaCd.Value = arrRet(0)
		frm1.txtBizAreaNm.Value = arrRet(1)
		frm1.txtBizAreaCd.focus	
		Set gActiveElement = document.activeElement		
	End If		
		
End Function
'==============================================================================================================================
Function OpenPoNo()
	
	Dim strRet
	Dim arrParam(2)
	Dim iCalledAspName
			
	If lblnWinEvent = True Or UCase(frm1.txtPoNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
		
	arrParam(0) = ""  'Return Flag
	arrParam(1) = ""  'Release Flag
	arrParam(2) = ""  'STO Flag
	
	iCalledAspName = AskPRAspName("m3111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", PopupParent.VB_INFORMATION, "m3111pa1", "X")
		lblnWinEvent = False
		Exit Function
	End If
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0) = "" Then
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement	
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus	
		Set gActiveElement = document.activeElement		
	End If	
		
End Function
'==============================================================================================================================
Function OpenItemCd1()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = "품목"						
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"	
	arrParam(2) = Trim(frm1.txtItemCd1.Value)		
'	arrParam(3) = Trim(frm1.txtItemNm.Value)		
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
	arrParam(5) = "품목"						

    arrField(0) = "B_Item.Item_Cd"					
    arrField(1) = "B_Item.Item_NM"					
    arrField(2) = "B_Plant.Plant_Cd"				
    arrField(3) = "B_Plant.Plant_NM"				
    
    arrHeader(0) = "품목"						
    arrHeader(1) = "품목명"						
    arrHeader(2) = "공장"						
    arrHeader(3) = "공장명"						
    
    iCalledAspName = AskPRAspName("m1111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m1111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam,arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd1.focus	
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtItemCd1.Value = arrRet(0)
		frm1.txtItemNm1.Value = arrRet(1)
		frm1.txtItemCd1.focus	
		Set gActiveElement = document.activeElement		
	End If	

End Function
'==============================================================================================================================
Function OpenItemCd2()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	
	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True

	arrParam(0) = "품목"						
	arrParam(1) = "B_Item_By_Plant,B_Plant,B_Item"	
	arrParam(2) = Trim(frm1.txtItemCd2.Value)		
'	arrParam(3) = Trim(frm1.txtItemNm.Value)		
	arrParam(4) = "B_Item_By_Plant.Plant_Cd = B_Plant.Plant_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd and B_Item.phantom_flg = " & FilterVar("N", "''", "S") & "  "
	arrParam(5) = "품목"						

    arrField(0) = "B_Item.Item_Cd"					
    arrField(1) = "B_Item.Item_NM"					
    arrField(2) = "B_Plant.Plant_Cd"				
    arrField(3) = "B_Plant.Plant_NM"				
    
    arrHeader(0) = "품목"						
    arrHeader(1) = "품목명"						
    arrHeader(2) = "공장"						
    arrHeader(3) = "공장명"						
    
    iCalledAspName = AskPRAspName("m1111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m1111pa1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam, arrField, arrHeader), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

   IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd2.focus	
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtItemCd2.Value = arrRet(0)
		frm1.txtItemNm2.Value = arrRet(1)
		frm1.txtItemCd2.focus	
		Set gActiveElement = document.activeElement		
	End If	

End Function



Function OpenDistRefNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "배부참조번호"				    
	arrParam(1) = "M_DISB_HIST"					    
	arrParam(2) = Trim(frm1.txtDistRefNo.value)	 
'	arrParam(3) = trim(frm1.txtDistTypenm.value)	
	arrParam(4) = " DIST_REF_NO IS NOT NULL "			
	arrParam(5) = "배부유형"			
	
    arrField(0) = "ED15" & Chr(11) & "DIST_REF_NO"					
    arrField(1) = "ED15" & Chr(11) & "ITEM_DOCUMENT_NO"
    arrField(2) = "ED15" & Chr(11) & "PROCESS_STEP"	
    arrField(3) = "DD10" & Chr(11) & "DISB_QRY_FR_DT"					
    arrField(4) = "DD10" & Chr(11) & "DISB_QRY_TO_DT"
    arrField(5) = "DD10" & Chr(11) & "DISB_DT"					
    arrField(6) = "DD10" & Chr(11) & "DISB_JOB_DT"				
			
    
    arrHeader(0) = "배부참조번호"				
    arrHeader(1) = "재고처리번호"
    arrHeader(2) = "경비발생단계"				
    arrHeader(3) = "배부대상기간(From)"
    arrHeader(4) = "배부대상기간(To)"				
    arrHeader(5) = "작업일"
    arrHeader(6) = "배부년월"									
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtDistRefNo.focus	
		Set gActiveElement = document.activeElement		
		Exit Function
	Else
		frm1.txtDistRefNo.Value = arrRet(0)	
		frm1.txtDistRefNo.focus	
		Set gActiveElement = document.activeElement		
	End If	
End Function

'==============================================================================================================================
Sub Form_Load()
    
    Call LoadInfTB19029                  
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")
    Call InitVariables                   
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")	
    
    frm1.txtBizAreaCd.focus 
	Set gActiveElement = document.activeElement
End Sub
'==============================================================================================================================
Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtFrDt.Focus
	End If
End Sub
'==============================================================================================================================
Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")	
        frm1.txtToDt.Focus
	End If
End Sub
'==============================================================================================================================
Function ChkKeyField()
	
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	
	ChkKeyField = true
	
	strWhere = " BIZ_AREA_CD =  " & FilterVar(frm1.txtBizAreaCd.value, "''", "S") & "  "
	
	Call CommonQueryRs(" BIZ_AREA_NM "," B_BIZ_AREA ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
	IF Len(lgF0) < 1 Then 
		Call DisplayMsgBox("17a003","X","사업장","X")
		frm1.txtBizAreaNm.value = ""
		ChkKeyField = False
		Exit Function
	End If
	
	strDataNm = split(lgF0,chr(11))
	
	frm1.txtBizAreaNm.value = strDataNm(0)
	
End Function
'==============================================================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'==============================================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False) 
    Set gActiveElement = document.activeElement
End Function
'==============================================================================================================================
Function FncBtnPrint() 
 
	Dim StrUrl	
	Dim intCnt
	dim var1,var2,var3,var4,var5,var6, var7
    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If
	
	IF ChkKeyField() = False Then 
		frm1.txtBizAreaCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
    End if
	
	with frm1
	    If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" Then
			Call DisplayMsgBox("17a003", "X","발생일", "X")			
			Exit Function
		End if   
		
		if (frm1.txtItemCd1.value <> "") AND  (frm1.txtItemCd2.value <> "") then
			if  UCase(frm1.txtItemCd1.value) > UCase(frm1.txtItemCd2.value)  then	
				Call DisplayMsgBox("17a003","X","품목","X")
				frm1.txtItemCd1.focus 
				Set gActiveElement = document.activeElement
				Exit Function
			End if  
		End if 
	End with
	
	On Error Resume Next    
	
	var1= UCase(frm1.txtBizAreaCd.value)
	
	If UCase(frm1.txtPoNo.value) = "" Then
		var2 = "%"
	Else
		var2 = UCase(frm1.txtPoNo.value)
	End If
	
	If UCase(frm1.txtItemCd1.value) = "" Then
		var3 = "%"
	Else
		var3= UCase(frm1.txtItemCd1.value)
	End If
	
	If UCase(frm1.txtItemCd2.value) = "" Then
		var4 = "ZZZZZZZZZZZZZZZZZZ"
	Else
		var4= UCase(frm1.txtItemCd2.value)
	End If
			
	If UCase(frm1.txtFrDt.text) = "" Then
		var5 = ("1900-01-01")
	Else
		var5 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,parent.gDateFormat,parent.gServerDateType) 'uniCdate(frm1.txtFrDt.text)
	End If
	
	If UCase(frm1.txtToDt.text) = "" Then
		var6 = ("2999-12-31")
	Else
		var6 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,parent.gDateFormat,parent.gServerDateType) 'uniCdate(frm1.txtToDt.text)
	End If
	
	
	If UCase(frm1.txtDistRefNo.value) = "" Then
		var7 = "%"
	Else
		var7 = UCase(Trim(frm1.txtDistRefNo.value))
	End If
					
	strUrl = strUrl & "bizarea|"		& var1
	strUrl = strUrl & "|pono|"			& var2
	strUrl = strUrl & "|fritem|"		& var3
	strUrl = strUrl & "|toitem|"		& var4
	strUrl = strUrl & "|frdt|"			& var5
	strUrl = strUrl & "|todt|"			& var6
	StrUrl = StrUrl & "|distrefno|"				& var7
	
	if frm1.rdoflg1.checked = True then
		ObjName = AskEBDocumentName("m6112oa1","ebr")
		Call FncEBRprint(EBAction, ObjName, strUrl)
	else
		ObjName = AskEBDocumentName("m6112oa2","ebr")
		Call FncEBRprint(EBAction, ObjName, strUrl)
	End if
	
	Call BtnDisabled(0)	
	Set gActiveElement = document.activeElement	
End Function
'==============================================================================================================================
Function BtnPreview() 
	On Error Resume Next                       
    
	Err.Clear                                                       
    
    If Not chkField(Document, "1") Then	
       Exit Function
    End If
    
    IF ChkKeyField() = False Then 
		frm1.txtBizAreaCd.focus
		Exit Function
    End if
    
    Dim strVal
    dim var1,var2,var3,var4,var5,var6, var7
	dim strUrl
	dim arrParam, arrField, arrHeader
    
    with frm1
	     If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then
			Call DisplayMsgBox("17a003", "X","발생일", "X")			
			Exit Function
		End if   

		if (frm1.txtItemCd1.value <> "") AND  (frm1.txtItemCd2.value <> "") then
			if  UCase(frm1.txtItemCd1.value) > UCase(frm1.txtItemCd2.value)  then	
				Call DisplayMsgBox("17a003","X","품목","X")
				frm1.txtItemCd1.focus 
				Set gActiveElement = document.activeElement
				Exit Function
			End if  
		End if 
	END WITH
	
	var1= UCase(frm1.txtBizAreaCd.value)
	
	If UCase(frm1.txtPoNo.value) = "" Then
		var2 = "%"
	Else
		var2 = UCase(frm1.txtPoNo.value)
	End If	
	
	If UCase(frm1.txtItemCd1.value) = "" Then
		var3 = "%"
	Else
		var3= UCase(frm1.txtItemCd1.value)
	End If
	
	If UCase(frm1.txtItemCd2.value) = "" Then
		var4 = "ZZZZZZZZZZZZZZZZZZ"
	Else
		var4= UCase(frm1.txtItemCd2.value)
	End If
		
	If UCase(frm1.txtFrDt.text) = "" Then
		var5 = ("1900-01-01")
	Else
		var5 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,parent.gDateFormat,parent.gServerDateType) 'uniCdate(frm1.txtFrDt.text)
	End If
	
	If UCase(frm1.txtToDt.text) = "" Then
		var6 = ("2999-12-31")
	Else
		var6 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,parent.gDateFormat,parent.gServerDateType) 'uniCdate(frm1.txtToDt.text)
	End If
	
	If UCase(frm1.txtDistRefNo.value) = "" Then
		var7 = "%"
	Else
		var7 = UCase(Trim(frm1.txtDistRefNo.value))
	End If	
				
	strUrl = strUrl & "bizarea|"		& var1
	strUrl = strUrl & "|pono|"			& var2
	strUrl = strUrl & "|fritem|"		& var3
	strUrl = strUrl & "|toitem|"		& var4
	strUrl = strUrl & "|frdt|"			& var5
	strUrl = strUrl & "|todt|"			& var6
	strUrl = strUrl & "|distrefno|"		& var7
	
	if frm1.rdoflg1.checked = True then		
		ObjName = AskEBDocumentName("m6112oa1","ebr")
		Call FncEBRPreview(ObjName, strUrl)		
	else		
		ObjName = AskEBDocumentName("m6112oa2","ebr")
		Call FncEBRPreview(ObjName, strUrl)
	End if
	
	Call BtnDisabled(0)	
	Set gActiveElement = document.activeElement		
End Function
'==============================================================================================================================
Function FncExit()
    FncExit = True
	Set gActiveElement = document.activeElement    
End Function
'☜: 아래 OBJECT Tag는 InterDev 사용자를 위한것으로 프로그램이 완성되면 아래 Include 코드로 대체되어야 한다 
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers
function setCookie(name, value, expire)
{
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>부대비배부내역</font></td>
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
								<TD CLASS="TD5" NOWRAP>발행유형</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="발행유형" NAME="rdoflg" id = "rdoflg1" Value="Y"  checked tag="12"><label for="rdoflg1">&nbsp;발주번호별&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="발행유형" NAME="rdoflg" id = "rdoflg2" Value="N"  tag="12"><label for="rdoflg2">&nbsp;품목별&nbsp;</label>
								</TD>									
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBizAreaCd"   SIZE=10 MAXLENGTH=10 ALT="사업장" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBizAreaCd()">
													   <INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=20 MAXLENGTH=18 ALT="사업장" tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>배부참조번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtDistRefNo"  SIZE=32 MAXLENGTH=18 ALT="배부참조번호" tag="11XXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDistRefNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenDistRefNo()">
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>발주번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo"  SIZE=32 MAXLENGTH=18 ALT="발주번호" tag="11XXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()">
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd1" SIZE=10 MAXLENGTH=18 ALT="품목" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd1()">
													   <INPUT TYPE=TEXT NAME="txtItemNm1" SIZE=20 MAXLENGTH=18 ALT="품목" tag="14"> ~</TD>
							</TR> 
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>						   
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd2" SIZE=10 MAXLENGTH=18 ALT="품목" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemCd2()">
													   <INPUT TYPE=TEXT NAME="txtItemNm2" SIZE=20 MAXLENGTH=18 ALT="품목" tag="14"></TD>					   
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>발생일</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발생일 NAME="txtFrDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</td>
											<td>&nbsp;~&nbsp;</td>
											<td>
												<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT=발생일 NAME="txtToDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
											</td>
										<tr>
									</table>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR >
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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
