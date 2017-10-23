<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111oa4_ko441
'*  4. Program Name         : Offer 대장 
'*  5. Program Desc         : Offer 대장 
'*  6. Component List       : 
'*  7. Modified date(First) : 2008/01/03
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
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
<!--'#########################################################################################################
'												1. 선 언 부 
'##########################################################################################################-->
<!-- '******************************************  1.1 Inc 선언   **********************************************
'	기능: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit		

Dim StartDate
Dim EndDate

EndDate = "<%=GetSvrDate%>"
StartDate = UNIDateAdd("m", -1, EndDate, Parent.gServerDateFormat)
EndDate   = UniConvDateAToB(EndDate, Parent.gServerDateFormat, Parent.gDateFormat)
StartDate = UniConvDateAToB(StartDate, Parent.gServerDateFormat, Parent.gDateFormat)  


Dim lgBlnFlgChgValue          
Dim lgIntFlgMode              
Dim lgIntGrpCount             

'==========================================  1.2.3 Global Variable값 정의  ===============================
'=========================================================================================================
Dim lblnWinEvent
Dim IsOpenPop

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                           
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtFrDt.Text	= StartDate 
	frm1.txtToDt.Text	= EndDate
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd1, "Q") 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd2, "Q") 
		frm1.txtPurGrpCd1.Tag = left(frm1.txtPurGrpCd1.Tag,1) & "4" & mid(frm1.txtPurGrpCd1.Tag,3,len(frm1.txtPurGrpCd1.Tag))
		frm1.txtPurGrpCd2.Tag = left(frm1.txtPurGrpCd2.Tag,1) & "4" & mid(frm1.txtPurGrpCd2.Tag,3,len(frm1.txtPurGrpCd2.Tag))
        frm1.txtPurGrpCd1.value = lgPGCd
        frm1.txtPurGrpCd2.value = lgPGCd
	End If
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M","NOCOOKIE","OA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "OA") %>
End Sub

'------------------------------------------  OpenBpCd()  -------------------------------------------------
Function OpenBpCd1()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"						
	arrParam(1) = "B_Biz_Partner"					
	arrParam(2) = Trim(frm1.txtBpCd1.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)			
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_CD"							
    arrField(1) = "BP_NM"							
    
    arrHeader(0) = "공급처"						
    arrHeader(1) = "공급처명"					
    
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

'------------------------------------------  OpenPurGrpCd()  -------------------------------------------------

Function OpenPurGrpCd1()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
    If frm1.txtPurGrpCd1.className = "protected" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd1.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd1.focus
		Exit Function
	Else
		frm1.txtPurGrpCd1.Value = arrRet(0)
		frm1.txtPurGrpNm1.Value = arrRet(1)
		frm1.txtPurGrpCd1.focus
	End If
	
End Function

'------------------------------------------  OpenBpCd()  -------------------------------------------------
Function OpenBpCd2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"						
	arrParam(1) = "B_Biz_Partner"					
	arrParam(2) = Trim(frm1.txtBpCd2.Value)			
	'arrParam(3) = Trim(frm1.txtBpNm.Value)			
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "
	arrParam(5) = "공급처"						
	
    arrField(0) = "BP_CD"							
    arrField(1) = "BP_NM"							
    
    arrHeader(0) = "공급처"						
    arrHeader(1) = "공급처명"					
    
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
'------------------------------------------  OpenPurGrpCd()  -------------------------------------------------
Function OpenPurGrpCd2()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
    If frm1.txtPurGrpCd2.className = "protected" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd2.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd2.focus
		Exit Function
	Else
		frm1.txtPurGrpCd2.Value = arrRet(0)
		frm1.txtPurGrpNm2.Value = arrRet(1)
		frm1.txtPurGrpCd2.focus
	End If	
End Function  

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029                       
    Call ggoOper.LockField(Document, "N")     
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)

    Call InitVariables                        
	Call GetValue_ko441()

    Call SetDefaultVal
    Call SetToolbar("1000000000001111")		
    
    frm1.txtFrDt.focus 
	Set gActiveElement = document.activeElement
End Sub

'===================================  txtFrDt_DblClick()  ========================================
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

'===================================  FncPrint()  ========================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function
'===================================  FncFind()  ========================================
Function FncFind() 
    Call parent.FncFind(parent.C_SINGLE , False)                       
	Set gActiveElement = document.activeElement
End Function
'===================================  FncBtnPrint()  ========================================
Function FncBtnPrint() 
	Dim StrUrl
	Dim lngPos
	Dim intCnt
	dim var1,var2,var3,var4,var5,var6,var7
    Dim ObjName
    	
    If Not chkField(Document, "1") Then									
       Exit Function
    End If	

    IF ChkKeyField() = False Then 
		Exit Function
    End if

	with frm1
	     If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then
			Call DisplayMsgBox("17a003", "X","발행기간", "X")			
			Exit Function
		End if   
	End with
	
	On Error Resume Next             
	
	lngPos = 0
	
	var1 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,parent.gDateFormat,parent.gServerDateType)'uniCdate(frm1.txtFrDt.text)
	var2 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,parent.gDateFormat,parent.gServerDateType)'uniCdate(frm1.txtToDt.text)
	var3 = Trim("KR")
	var4 = FilterVar(Trim(UCase(frm1.txtBpCd1.value)), "" ,  "SNM") 
	var5 = FilterVar(Trim(UCase(frm1.txtBpCd2.value)), "" ,  "SNM")
	var6 = FilterVar(Trim(UCase(frm1.txtPurGrpCd1.value)), "" ,  "SNM") 
	var7 = FilterVar(Trim(UCase(frm1.txtPurGrpCd2.value)), "" ,  "SNM")  
	
'	if	var4="" then
'		var4="%"
'	end if
'	if	var6="" then
'		var6="%"
'	end if
'	
'	if var5="" then 
'	 var5="ZZZZZZZZZZ"
'	end if
'	
'	if var7="" then 
'	 var7="ZZZZZZZZZZ"
'	end if
        		
	For intCnt = 1 To 3
	    lngPos = InStr(lngPos + 1, GetUserPath, "/")
	Next
		
	strUrl = strUrl & "fr_dt|" & var1 & "|to_dt|" & var2 & "|country|" & var3 
	strUrl = strUrl & "|fr_bp|" & var4
	strUrl = strUrl & "|to_bp|" & var5
	strUrl = strUrl & "|fr_pur|" & var6
	strUrl = strUrl & "|to_pur|" & var7

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
	
    If lgPLCd<>"" Then
        strUrl = strUrl & "|FR_PLANT_CD|" & lgPLCd 
        strUrl = strUrl & "|TO_PLANT_CD|" & lgPLCd 
    Else
        strUrl = strUrl & "|FR_PLANT_CD|" & "" 
        strUrl = strUrl & "|TO_PLANT_CD|" & "ZZZZZZZZZZ" 
    End If
	

	ObjName = AskEBDocumentName("m3111oa4_ko441","ebr")
	Call FncEBRprint(EBAction, ObjName, strUrl)
	Call BtnDisabled(0)	
	Set gActiveElement = document.activeElement
End Function
'===================================  BtnPreview()  ========================================
Function BtnPreview() 
    If Not chkField(Document, "1") Then	
       Exit Function
    End If

    IF ChkKeyField() = False Then 
		Exit Function
    End if

	with frm1
	    If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,Parent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" then
			Call DisplayMsgBox("17a003", "X","발행기간", "X")
			Exit Function
		End if   
	End with

	dim var1,var2,var3,var4,var5,var6,var7
	
	dim strUrl
	dim arrParam, arrField, arrHeader
    Dim ObjName
		
	var1 = UniConvDateToYYYYMMDD(frm1.txtFrDt.Text,parent.gDateFormat,parent.gServerDateType) 'uniCdate(frm1.txtFrDt.text)
	var2 = UniConvDateToYYYYMMDD(frm1.txtToDt.Text,parent.gDateFormat,parent.gServerDateType) 'uniCdate(frm1.txtToDt.text)
	var3 = Trim("KR")
	var4 = FilterVar(Trim(UCase(frm1.txtBpCd1.value)), "" ,  "SNM") 
	var5 = FilterVar(Trim(UCase(frm1.txtBpCd2.value)), "" ,  "SNM")
	var6 = FilterVar(Trim(UCase(frm1.txtPurGrpCd1.value)), "" ,  "SNM") 
	var7 = FilterVar(Trim(UCase(frm1.txtPurGrpCd2.value)), "" ,  "SNM")  
	

	If frm1.txtBpCd1.value = "" then frm1.txtBpNm1.value = ""
	If frm1.txtBpCd2.value = "" then frm1.txtBpNm2.value = ""
	If frm1.txtPurGrpCd1.value = "" then frm1.txtPurGrpNm1.value = ""
	If frm1.txtPurGrpCd2.value = "" then frm1.txtPurGrpNm2.value = ""
	 
	
	strUrl = strUrl & "fr_dt|" & var1 & "|to_dt|" & var2 & "|country|" & var3 
	strUrl = strUrl & "|fr_bp|" & var4
	strUrl = strUrl & "|to_bp|" & var5
	strUrl = strUrl & "|fr_pur|" & var6
	strUrl = strUrl & "|to_pur|" & var7

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
	
    If lgPLCd<>"" Then
        strUrl = strUrl & "|FR_PLANT_CD|" & lgPLCd 
        strUrl = strUrl & "|TO_PLANT_CD|" & lgPLCd 
    Else
        strUrl = strUrl & "|FR_PLANT_CD|" & "" 
        strUrl = strUrl & "|TO_PLANT_CD|" & "ZZZZZZZZZZ" 
    End If
	

	ObjName = AskEBDocumentName("m3111oa4_ko441","ebr")
	Call FncEBRPreview(ObjName, strUrl)
	Call BtnDisabled(0)
End Function

'===================================  FncExit()  ========================================
Function FncExit()
    FncExit = True
	Set gActiveElement = document.activeElement
End Function

'==========================================  2.2.6 ChkKeyField()  =======================================
 Function ChkKeyField()
	
	Dim strDataCd, strDataNm
    Dim strWhere 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       
	
	ChkKeyField = true
	
	If Trim(frm1.txtBpCd1.value) <> "" Then
		strWhere = " BP_CD =  " & FilterVar(frm1.txtBpCd1.value, "''", "S") & "  "
	
		Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","공급처","X")
			frm1.txtBpCd1.focus 
			frm1.txtBpNm1.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtBpNm1.value = strDataNm(0)
	End If

	If Trim(frm1.txtBpCd2.value) <> "" Then
		strWhere = " BP_CD =  " & FilterVar(frm1.txtBpCd2.value, "''", "S") & "  "
	
		Call CommonQueryRs(" BP_NM "," B_BIZ_PARTNER ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","공급처","X")
			frm1.txtBpCd2.focus 
			frm1.txtBpNm2.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtBpNm2.value = strDataNm(0)
	End If

	If Trim(frm1.txtPurGrpCd1.value) <> "" Then
		strWhere = " PUR_GRP =  " & FilterVar(frm1.txtPurGrpCd1.value, "''", "S") & "  "
	
		Call CommonQueryRs(" PUR_GRP_NM "," B_PUR_GRP ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","구매그룹","X")
			frm1.txtPurGrpCd1.focus 
			frm1.txtPurGrpNm1.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPurGrpNm1.value = strDataNm(0)
	End If

	If Trim(frm1.txtPurGrpCd2.value) <> "" Then
		strWhere = " PUR_GRP =  " & FilterVar(frm1.txtPurGrpCd2.value, "''", "S") & "  "
	
		Call CommonQueryRs(" PUR_GRP_NM "," B_PUR_GRP ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("17a003","X","구매그룹","X")
			frm1.txtPurGrpCd2.focus 
			frm1.txtPurGrpNm2.value = ""
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtPurGrpNm2.value = strDataNm(0)
	End If

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>OFFER 대장</font></td>
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
								<TD CLASS="TD5" NOWRAP>OFFER작성일</TD>
								<TD CLASS="TD6" NOWRAP>
									<table cellspacing=0 cellpadding=0>
										<tr>
											<td>
												<script language =javascript src='./js/m3111oa4_fpDateTime1_txtFrDt.js'></script>
											</td>
											<td>~</td>
											<td>
												<script language =javascript src='./js/m3111oa4_fpDateTime2_txtToDt.js'></script>
											</td>
										<tr>
									</table>
								</TD>	
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>공급처</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd1" SIZE=10 MAXLENGTH=10 ALT="공급처" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBpCd1()">
													   <INPUT TYPE=TEXT NAME="txtBpNm1" SIZE=20 MAXLENGTH=18 ALT="공급처" tag="14"> ~ </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd2" SIZE=10 MAXLENGTH=10 ALT="공급처" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenBpCd2()">
													   <INPUT TYPE=TEXT NAME="txtBpNm2" SIZE=20 MAXLENGTH=18 ALT="공급처" tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>구매그룹</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrpCd1" SIZE=10 MAXLENGTH=10 ALT="구매그룹" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrpCd1()">
													   <INPUT TYPE=TEXT NAME="txtPurGrpNm1" SIZE=20 MAXLENGTH=18 ALT="구매그룹" tag="14"> ~ </TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPurGrpCd2" SIZE=10 MAXLENGTH=10 ALT="구매그룹" tag="11NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrpCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrpCd2()">
													   <INPUT TYPE=TEXT NAME="txtPurGrpNm2" SIZE=20 MAXLENGTH=18 ALT="구매그룹" tag="14"></TD>					   
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
					<TD><BUTTON NAME="btnRun" CLASS="CLSSBTN" ONCLICK="vbscript:BtnPreview()" Flag=1>미리보기</BUTTON>&nbsp;<BUTTON NAME="btnPrint" CLASS="CLSSBTN" ONCLICK="vbscript:FncBtnPrint()" Flag=1>인쇄</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" tabindex="-1" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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
