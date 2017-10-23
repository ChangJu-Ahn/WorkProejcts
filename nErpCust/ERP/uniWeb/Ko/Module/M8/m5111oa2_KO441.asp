<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 거래실적집계표(구매)
'*  5. Program Desc         : 거래실적집계표(구매)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/06/29
'*  8. Modified date(Last)  : 2001/06/29
'*  9. Modifier (First)     : shin Jin Hyun
'* 10. Modifier (Last)      : Ma Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'******************************************  1.1 Inc 선언   ****************************************-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  =====================================-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ===================================-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance
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
	
	frm1.rdoApFlg(0).checked = True
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPlantCd, "Q") 
		frm1.txtPlantCd.Tag = left(frm1.txtPlantCd.Tag,1) & "4" & mid(frm1.txtPlantCd.Tag,3,len(frm1.txtPlantCd.Tag))
        frm1.txtPlantCd.value = lgPLCd
	End If
End Sub

'========================================  LoadInfTB19029()  =================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "M","NOCOOKIE","OA") %>
	<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub

'------------------------------------------  OpenCommPopup()  -------------------------------------------------
Function OpenCommPopup(arrHeader, arrField, arrParam, arrRet)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	If arrRet(0) = "" Then
		OpenCommPopup = False
	Else
		OpenCommPopup = True
		lgBlnFlgChgValue = True
	End If
	
End Function

'------------------------------------------  OpenIvType()  -------------------------------------------------
Function OpenIvType()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	If lblnWinEvent = True Or UCase(frm1.txtIvTypeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	lblnWinEvent = True
	
	arrHeader(0) = "매입형태"						' Header명(0)
    arrHeader(1) = "매입형태명"						' Header명(1)
    
    arrField(0) = "IV_TYPE_CD"							' Field명(0)
    arrField(1) = "IV_TYPE_NM"							' Field명(1)
    
	arrParam(0) = "매입형태"						' 팝업 명칭 
	arrParam(1) = "M_IV_TYPE"							' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			' Code Condition
	'arrParam(2) = Trim(frm1.txtIvTypeCd.Value)			' Code Condition
	'arrParam(3) = Trim(frm1.txtIvTypeNm.Value)			' Name Cindition
	'arrParam(4) = "import_flg='N'"						' Where Condition
	arrParam(5) = "매입형태"						' TextBox 명칭 
	
    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
		frm1.txtIvTypeCd.Value = arrRet(0)
		frm1.txtIvTypeNm.Value = arrRet(1)
    End If
    frm1.txtIvTypeCd.focus
    lblnWinEvent = False
    Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenSppl()  -------------------------------------------------
Function OpenSppl()
	Dim arrRet
	Dim arrHeader(6), arrField(6), arrParam(5) 
	
	arrHeader(2) = "사업자등록번호"									' Header명(2)
    arrField(0) = "B_BIZ_PARTNER.BP_Cd"									' Field명(0)
    arrField(1) = "B_BIZ_PARTNER.BP_Nm"								    ' Field명(1)
    arrField(2) = "B_BIZ_PARTNER.BP_RGST_NO"							' Field명(2)

	If lblnWinEvent = True Or UCase(frm1.txtPayeeCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function    
	lblnWinEvent = True

	arrHeader(0) = "지급처"											' Header명(0)
	arrHeader(1) = "지급처명"										' Header명(1)

	arrParam(0) = "지급처"											' 팝업 명칭 
	arrParam(1) = "B_BIZ_PARTNER, B_BIZ_PARTNER_FTN	"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtPayeeCd.Value)							' Code Condition%>
	'arrParam(2) = Trim(frm1.txtPayeeCd.Value)							' Code Condition%>
	arrParam(4) = "B_BIZ_PARTNER.BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And B_BIZ_PARTNER.usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(4) = arrParam(4) & " AND B_BIZ_PARTNER.BP_CD = B_BIZ_PARTNER_FTN.PARTNER_BP_CD " & " AND  B_BIZ_PARTNER_FTN.PARTNER_FTN = " & FilterVar("MPA", "''", "S") & " " 			<%' Where Condition%>
    '지급처팝업오류수정(2003.09.18)
    'If Trim(frm1.txtPayeeCd.Value) <> "" then
	'	arrParam(4) = arrParam(4) & " AND B_BIZ_PARTNER_FTN.BP_CD = '" & FilterVar(Trim(frm1.txtPayeeCd.Value), " " , "SNM") & "'"  
	'End If
	arrParam(5) = "지급처"											' TextBox 명칭 

    If OpenCommPopup(arrHeader, arrField, arrParam, arrRet) then
				frm1.txtPayeeCd.Value = arrRet(0) : frm1.txtPayeeNm.Value = arrRet(1)
    end if
    frm1.txtPayeeCd.focus
    lblnWinEvent = False
    Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenIvNo()  -------------------------------------------------
Function OpenIvNo()
	
	Dim strRet
	Dim arrParam(0)
	Dim iCalledAspName
		
	If lblnWinEvent = True Or UCase(frm1.txtIvNo.className) = UCase(parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
	arrParam(0) = ""
		
	iCalledAspName = AskPRAspName("m5111pa1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "m5111pa1", "X")
		lgIsOpenPop = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")


	lblnWinEvent = False
	
	If strRet(0) = "" Then
		frm1.txtIvNo.focus
		Exit Function
	Else
		frm1.txtIvNo.value = strRet(0)
		frm1.txtIvNo.focus
	End If	
	Set gActiveElement = document.activeElement	
End Function

'------------------------------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
    If frm1.txtPlantCd.className = "protected" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "PLANT_CD"	
    arrField(1) = "PLANT_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		
    
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

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
	
    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call GetValue_ko441()
    Call SetDefaultVal
    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 
	frm1.txtIvTypeCd.focus
   
End Sub
'=======================================  Form_QueryUnload()  ======================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    
End Sub

'==================================  OCX_EVENT  ==========================================
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrDt.focus
	End if
End Sub

Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.focus
	End if
End Sub

'======================================  FncPrint()  =========================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'======================================  FncFind()  =========================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)  
    Set gActiveElement = document.activeElement
End Function

'==========================================  2.2.6 ChkKeyField()  =======================================
 Function ChkKeyField()
	
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	
	ChkKeyField = true
	
	If Trim(frm1.txtPlantCd.value) <> "" Then
		strWhere = " PLANT_CD =  " & FilterVar(frm1.txtPlantCd.value, "''", "S") & "  "
	
		Call CommonQueryRs(" PLANT_NM "," B_PLANT ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","공장","X")
			frm1.txtPlantNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPlantNm.value = strDataNm(0)
	End If	
	

	If Trim(frm1.txtIvTypeCd.value) <> "" Then
		strWhere = " IV_TYPE_CD =  " & FilterVar(frm1.txtIvTypeCd.value, "''", "S") & "  "
	
		Call CommonQueryRs(" IV_TYPE_NM "," M_IV_TYPE ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","매입형태","X")
			frm1.txtIvTypeNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtIvTypeNm.value = strDataNm(0)
	End If	
	

	If Trim(frm1.txtPayeeCd.value) <> "" Then
		strWhere = " BP_CD =  " & FilterVar(frm1.txtPayeeCd.value, "''", "S") & "  "
		Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","지급처","X")
			frm1.txtPayeeNm.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPayeeNm.value = strDataNm(0)
	End If
End Function
'======================================  FncBtnPrint()  =========================================
Function FncBtnPrint() 
    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
    IF ChkKeyField() = False Then 
		frm1.txtIvTypeCd.focus
		Exit Function
    End if

	With frm1
		If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then	
            Call DisplayMsgBox("17a003","X","계산서일자","X")				      
            Exit Function
        End if   
	End with
	On Error Resume Next                                                    '☜: Protect system from crashing
	
	Dim var1,var2,var3,var4,var5,var6,var7,var8,var9
	
	Dim strUrl
	Dim arrParam, arrField, arrHeader
		
	If Trim(frm1.txtIvTypeCd.value) <> "" Then
		var1 = Trim(frm1.txtIvTypeCd.value) : var2 = Trim(frm1.txtIvTypeCd.value) 
	Else
		var1 = "" : var2 = "ZZZZZ"
	End If
	
	If Trim(frm1.txtPayeeCd.value) <> "" Then
		var3 = Trim(frm1.txtPayeeCd.value) : var4 = Trim(frm1.txtPayeeCd.value) 
	Else
		var3 = "" : var4 = "ZZZZZZZZZZ"
	End If

	If Trim(frm1.txtIvNo.value) <> "" Then
		var5 = Trim(frm1.txtIvNo.value) : var6 = Trim(frm1.txtIvNo.value) 
	Else
		var5 = "" : var6 = "ZZZZZZZZZZZZZZZZZZ"
	End If

	var7 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,parent.gDateFormat,parent.gServerDateType)  'uniCdate(frm1.txtFrDt.Text)
	var8 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,parent.gDateFormat,parent.gServerDateType)  'uniCdate(frm1.txtToDt.Text)
			
	var9 = Trim(frm1.txtPlantCd.value)

	strUrl = strUrl & "fr_ivtype|" & var1 & "|to_ivtype|" & var2 & "|fr_payeecd|" & var3 & "|to_payeecd|" & var4 & "|fr_ivno|" & var5 & "|to_ivno|" & var6 & "|fr_dt|" & var7 & "|to_dt|" & var8 & "|plantcd|" & var9

    If lgBACd<>"" Then
        strUrl = strUrl & "|FR_BIZ_AREA|" & lgBACd 
        strUrl = strUrl & "|TO_BIZ_AREA|" & lgBACd 
    Else
        strUrl = strUrl & "|FR_BIZ_AREA|" & "" 
        strUrl = strUrl & "|TO_BIZ_AREA|" & "ZZZZZZZZZZ" 
    End If

    If lgPGCd<>"" Then
        strUrl = strUrl & "|FR_PUR_GRP|" & lgPGCd 
        strUrl = strUrl & "|TO_PUR_GRP|" & lgPGCd 
    Else
        strUrl = strUrl & "|FR_PUR_GRP|" & "" 
        strUrl = strUrl & "|TO_PUR_GRP|" & "ZZZZZZZZZZ" 
    End If

    If lgPOCd<>"" Then
        strUrl = strUrl & "|FR_PUR_ORG|" & lgPOCd 
        strUrl = strUrl & "|TO_PUR_ORG|" & lgPOCd 
    Else
        strUrl = strUrl & "|FR_PUR_ORG|" & "" 
        strUrl = strUrl & "|TO_PUR_ORG|" & "ZZZZZZZZZZ" 
    End If

'----------------------------------------------------------------
' Print 함수에서 호출 
'----------------------------------------------------------------
    If frm1.rdoApFlg(0).checked = True then								'집계 
		ObjName = AskEBDocumentName("m5111oa3_KO441","ebr")
	Else																'상세 
		ObjName = AskEBDocumentName("m5111oa2_KO441","ebr")
	End If

	Call FncEBRprint(EBAction, ObjName, strUrl)
'----------------------------------------------------------------	
End Function

'======================================  BtnPreview()  =========================================
Function BtnPreview() 
'On Error Resume Next                                                    '☜: Protect system from crashing
    
    If Not chkField(Document, "1") Then									<%'⊙: This function check indispensable field%>
       Exit Function
    End If
    
    IF ChkKeyField() = False Then 
		frm1.txtIvTypeCd.focus
		Exit Function
    End if
    
	with frm1
	    If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then	
            Call DisplayMsgBox("17a003","X","계산서일자","X")				      
            Exit Function
        End if 	
	End with

	dim var1,var2,var3,var4,var5,var6,var7,var8,var9
	
	dim strUrl
	dim arrParam, arrField, arrHeader
		
	If Trim(frm1.txtIvTypeCd.value) <> "" Then
		var1 = Trim(frm1.txtIvTypeCd.value) : var2 = Trim(frm1.txtIvTypeCd.value) 
	else
		var1 = "" : var2 = "ZZZZZ"
	End If
	
	If Trim(frm1.txtPayeeCd.value) <> "" Then
		var3 = Trim(frm1.txtPayeeCd.value) : var4 = Trim(frm1.txtPayeeCd.value) 
	else
		var3 = "" : var4 = "ZZZZZZZZZZ"
	End If

	If Trim(frm1.txtIvNo.value) <> "" Then
		var5 = Trim(frm1.txtIvNo.value) : var6 = Trim(frm1.txtIvNo.value) 
	else
		var5 = "" : var6 = "ZZZZZZZZZZZZZZZZZZ"
	End If

	var7 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,parent.gDateFormat,parent.gServerDateType)  'uniCdate(frm1.txtFrDt.Text)
	var8 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,parent.gDateFormat,parent.gServerDateType)  'uniCdate(frm1.txtToDt.Text)
			
	var9 = Trim(frm1.txtPlantCd.value) 

	strUrl = strUrl & "fr_ivtype|" & var1 & "|to_ivtype|" & var2 & "|fr_payeecd|" & var3 & "|to_payeecd|" & var4 & "|fr_ivno|" & var5 & "|to_ivno|" & var6 & "|fr_dt|" & var7 & "|to_dt|" & var8 & "|plantcd|" & var9

    If lgBACd<>"" Then
        strUrl = strUrl & "|FR_BIZ_AREA|" & lgBACd 
        strUrl = strUrl & "|TO_BIZ_AREA|" & lgBACd 
    Else
        strUrl = strUrl & "|FR_BIZ_AREA|" & "" 
        strUrl = strUrl & "|TO_BIZ_AREA|" & "ZZZZZZZZZZ" 
    End If

    If lgPGCd<>"" Then
        strUrl = strUrl & "|FR_PUR_GRP|" & lgPGCd 
        strUrl = strUrl & "|TO_PUR_GRP|" & lgPGCd 
    Else
        strUrl = strUrl & "|FR_PUR_GRP|" & "" 
        strUrl = strUrl & "|TO_PUR_GRP|" & "ZZZZZZZZZZ" 
    End If

    If lgPOCd<>"" Then
        strUrl = strUrl & "|FR_PUR_ORG|" & lgPOCd 
        strUrl = strUrl & "|TO_PUR_ORG|" & lgPOCd 
    Else
        strUrl = strUrl & "|FR_PUR_ORG|" & "" 
        strUrl = strUrl & "|TO_PUR_ORG|" & "ZZZZZZZZZZ" 
    End If


    If frm1.rdoApFlg(0).checked = True then								'집계 
		ObjName = AskEBDocumentName("m5111oa3_KO441","ebr")
	Else																'상세 
		objName = AskEBDocumentName("m5111oa2_KO441","ebr")
	End If
	Call FncEBRPreview(ObjName, strUrl)		
	
End Function
'======================================  FncExit()  =========================================
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>구매거래실적집계표</font></td>
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
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 ALT="공장" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">
													   <INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=20 ALT="공장" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>매입형태</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtIvTypeCd" SIZE=10 MAXLENGTH=5 ALT="매입형태" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvType()">
													   <INPUT TYPE=TEXT NAME="txtIvTypeNm" SIZE=20 MAXLENGTH=20 ALT="매입형태" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>지급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPayeeCd" SIZE=10 MAXLENGTH=10 ALT="지급처" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSppl()">
													   <INPUT TYPE=TEXT NAME="txtPayeeNm" SIZE=20 MAXLENGTH=20 ALT="지급처" tag="14X"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>매입번호</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtIvNo" SIZE=32 MAXLENGTH=18 ALT="매입번호" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvNo()"></TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>계산서일자</TD>
								<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m5111oa2_fpDateTime1_txtFrDt.js'></script> ~
													   <script language =javascript src='./js/m5111oa2_fpDateTime2_txtToDt.js'></script></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>출력형태</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio Class="Radio" ALT="출력형태" NAME="rdoApFlg" id = "rdoApflg1" Value="Y" tag="12" checked><label for="rdoApFlg">&nbsp;집계&nbsp;</label>
													   <INPUT TYPE=radio Class="Radio" ALT="출력형태" NAME="rdoApFlg" id = "rdoApflg2" Value="N" tag="12"><label for="rdoApFlg">&nbsp;상세&nbsp;</label>
								</TD>									
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
					    <BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;		    
					    <BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON>		
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
