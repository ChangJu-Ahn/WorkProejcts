<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 일괄출하등록 
'*  3. Program ID           : S4111BA1
'*  4. Program Name         : 
'*  5. Program Desc         : 출하관리관리 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/07/02
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
' =======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>		            '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "S4111bb1.asp"
Const BIZ_PGM_JUMP_ID = "s4111ma6"

Const C_PopPlant		= 1			' 공장 
Const C_PopMovType		= 2			' 출하형태 
Const C_PopShipToParty	= 3			' 납품처 
Const C_PopSalesGrp		= 4			' 영업그룹 

Const C_PopTransMeth	= 5			' 운송방법 

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgBlnOpenedFlag
Dim lgBlnRegChecked
Dim lgBlnOpenPop			' Popup Window의 Open 여부 

Dim ToDateOfDB

ToDateOfDB = UNIConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat,parent.gDateFormat)

'========================================================================================================
Sub InitVariables()
End Sub

'========================================================================================================
Sub SetDefaultVal()
	If parent.gPlant <> "" And Trim(frm1.txtConPlant.value) = "" Then
		frm1.txtConPlant.value = parent.gPlant
		Call txtConPlant_OnChange()
	End If

	If parent.gSalesGrp <> "" And Trim(frm1.txtConSalesGrp.value) = "" Then
		frm1.txtConSalesGrp.value = parent.gSalesGrp
		Call txtConSalesGrp_OnChange()
	End If

	lgBlnRegChecked = True
	
	frm1.txtConFromDt.Text	= ToDateOfDB
	frm1.txtConToDt.Text	= ToDateOfDB
	frm1.txtPromiseDt.Text	= ToDateOfDB
	frm1.txtActualGIDt.Text	= ToDateOfDB
	
	Call chkGIFlag_OnClick
	
	frm1.txtConPlant.Focus
End Sub	

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "BA") %>
End Sub

'==========================================================================================================
Function JumpChgCheck()
	Call CookiePage(1)
	Call PgmJump(BIZ_PGM_JUMP_ID)
End Function

'==========================================================================================================
Function CookiePage(Byval pvKubun)

	On Error Resume Next
	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>
	Dim iStrTemp, iArrVal

	With frm1
		If pvKubun = 1 Then
			WriteCookie CookieSplit , .txtConPlant.value & Parent.gColSep & .txtPromiseDt.Text & Parent.gColSep & _
									  .txtConMovType.value & Parent.gColSep & .txtConShipToParty.value

		ElseIf pvKubun = 0 Then
			iStrTemp = ReadCookie(CookieSplit)
			
			If Trim(Replace(iStrTemp, parent.gColSep, "")) = "" Then Exit Function
			
			WriteCookie CookieSplit , ""
		End If
	End With
End Function

'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
	lgBlnOpenedFlag	 = True
End Sub
	
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'========================================================================================================
Function FncExit()
    FncExit = True
End Function

'========================================================================================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	With frm1
		Select Case pvIntWhere
			Case C_PopPlant		'공장 
				iArrParam(1) = "dbo.B_PLANT"									
				iArrParam(2) = Trim(.txtConPlant.value)				
				iArrParam(3) = ""										
				iArrParam(4) = ""										
				
				iArrField(0) = "ED15" & Parent.gColSep & "PLANT_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "PLANT_NM"
							
				iArrHeader(0) = .txtConPlant.alt						
				iArrHeader(1) = .txtConPlantNm.alt					
	
				.txtConPlant.focus

			Case C_PopMovType	'출하형태 
				iArrParam(1) = "dbo.B_MINOR MN "		
				iArrParam(2) = Trim(.txtConMovType.value)					
				iArrParam(3) = ""											
				iArrParam(4) = "MN.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND EXISTS (SELECT * FROM dbo.S_SO_TYPE_CONFIG ST WHERE	ST.MOV_TYPE = MN.MINOR_CD) "			
				
				iArrField(0) = "ED15" & Parent.gColSep & "MN.MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MN.MINOR_NM"
				
				iArrHeader(0) = .txtConMovType.alt							
				iArrHeader(1) = .txtConMovTypeNm.alt	
				
				frm1.txtConMovType.focus

			Case C_PopShipToParty	'납품처 
				iArrParam(1) = "dbo.B_BIZ_PARTNER BP INNER JOIN dbo.B_COUNTRY CT ON (CT.COUNTRY_CD = BP.CONTRY_CD)"								
				iArrParam(2) = Trim(.txtConShipToParty.value)			
				iArrParam(3) = ""											
				iArrParam(4) = "BP.BP_TYPE IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND EXISTS (SELECT * FROM B_BIZ_PARTNER_FTN BPF WHERE BPF.PARTNER_BP_CD = BP.BP_CD AND BPF.PARTNER_FTN = " & FilterVar("SSH", "''", "S") & ")"						
	
				iArrField(0) = "ED15" & Parent.gColSep & "BP.BP_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "BP.BP_NM"
				iArrField(2) = "ED10" & Parent.gColSep & "BP.CONTRY_CD"
				iArrField(3) = "ED20" & Parent.gColSep & "CT.COUNTRY_NM"
    
				iArrHeader(0) = .txtConShipToParty.alt					
				iArrHeader(1) = .txtConShipToPartyNm.alt					
				iArrHeader(2) = "국가"
				iArrHeader(3) = "국가명"

				.txtConShipToParty.focus
			
			' 영업그룹 
			Case C_PopSalesGrp												
				iArrParam(1) = "dbo.B_SALES_GRP"
				iArrParam(2) = Trim(.txtConSalesGrp.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
					
				iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
				iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
			    iArrHeader(0) = .txtConSalesGrp.Alt
			    iArrHeader(1) = .txtConSalesGrpNm.Alt
				    
			    .txtConSalesGrp.focus

		End Select
	End With
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
	If iArrRet(0) <> "" Then
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

' 입력 관련 Popup
'=========================================
Function OpenPopUp(Byval pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True

	With frm1
		Select Case pvIntWhere
			Case C_PopTransMeth	'운송방법 
				iArrParam(1) = "dbo.B_MINOR"
				iArrParam(2) = Trim(.txtTransMeth.value)
				iArrParam(3) = ""											
				iArrParam(4) = "MAJOR_CD = " & FilterVar("B9009", "''", "S") & ""
				
				iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"
							
				iArrHeader(0) = .txtTransMeth.alt						
				iArrHeader(1) = .txtTransMethNm.alt						

				.txtTransMeth.focus
		End Select
	End With
	
	iArrParam(0) = iArrHeader(0)							' 팝업 Title
	iArrParam(5) = iArrHeader(0)							' 조회조건 명칭 

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) <> "" Then
		OpenPopUp = SetPopUp(iArrRet,pvIntWhere)
	End If	
	
End Function

'=========================================
Function SetConPopUp(ByVal pvArrRet,ByVal pvIntWhere)
	SetConPopUp = False

	With frm1
		Select Case pvIntWhere
			Case C_PopPlant
				.txtConPlant.value = pvArrRet(0)
				.txtConPlantNm.value = pvArrRet(1) 

			Case C_PopMovType
				.txtConMovType.value = pvArrRet(0)
				.txtConMovTypeNm.value = pvArrRet(1) 

			Case C_PopShipToParty
				.txtConShipToParty.value = pvArrRet(0)
				.txtConShipToPartyNm.value = pvArrRet(1) 

			Case C_PopSalesGrp
				.txtConSalesGrp.value = pvArrRet(0)
				.txtConSalesGrpNm.value = pvArrRet(1) 
		End Select
	End With

	SetConPopUp = True
End Function

'========================================
Function SetPopUp(ByVal pvArrRet,ByVal pvIntWhere)

	SetPopup = False

	With frm1
		Select Case pvIntWhere
			Case C_PopTransMeth
				.txtTransMeth.value = pvArrRet(0)
				.txtTransMethNm.value = pvArrRet(1) 
		End Select
	End With

	SetPopup = True
End Function

'	Description : 코드값에 해당하는 명을 Display한다.
'====================================================================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)
	On Error Resume Next

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(5), iArrTemp

	GetCodeName = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, parent.gColSep)
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		
		If pvIntWhere = C_PopTransMeth Then
			GetCodeName = SetPopup(iArrRs, pvIntWhere)
		Else
			GetCodeName = SetConPopup(iArrRs, pvIntWhere)
		End If
		
	Else
		' 관련 Popup Display
		If err.number = 0 Then
			If lgBlnOpenedFlag Then
				If pvIntWhere = C_PopTransMeth Then
					GetCodeName = OpenPopup(pvIntWhere)
				Else
					GetCodeName = OpenConPopup(pvIntWhere)
				End If
			End If
		Else
			MsgBox Err.description, vbInformation,Parent.gLogoName
			Err.Clear
		End If
	End if
End Function

'	Description : 출하형태에 대한 Description Fetch
'====================================================================================================
Function GetMovTypeInfo()
	On Error Resume Next

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(5), iArrTemp

	GetMovTypeInfo = False

	iStrSelectList = " MN.MINOR_CD, MN.MINOR_NM "
	iStrFromList   = " dbo.B_MINOR MN "
	iStrWhereList  = " MN.MINOR_CD =  " & FilterVar(frm1.txtConMovType.value, "''", "S") & "" & _
					 " AND MN.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " " & _
					 " AND EXISTS (SELECT * FROM dbo.S_SO_TYPE_CONFIG ST WHERE	ST.MOV_TYPE = MN.MINOR_CD) "
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, parent.gColSep)
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		
		GetMovTypeInfo = SetConPopup(iArrRs, C_PopMovType)
		
	Else
		' 관련 Popup Display
		If err.number = 0 Then
			GetMovTypeInfo = OpenConPopup(C_PopMovType)
		Else
			MsgBox Err.description, vbInformation,Parent.gLogoName
			Err.Clear
		End If
	End if
End Function

'=======================================================================================================
Function ExeReflect() 
	Call BtnDisabled(1)
	Dim iStrVal

	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Or Not chkField(Document, "2") Then
		Call BtnDisabled(0)
		Exit Function
	End If

	With frm1
		If Not ValidDateCheck(.txtConFromDt, .txtConToDt) Then
			Call BtnDisabled(0)
			Exit Function
		End If

'		If Not ValidDateCheck(.txtConToDt, .txtPromiseDt) Then
'			Call BtnDisabled(0)
'			Exit Function
'		End If

'		If .chkGIFlag.checked Then
'			If Not ValidDateCheck(.txtPromiseDt, .txtActualGiDt) Then
'				Call BtnDisabled(0)
'				Exit Function
'			End If
'		End If

		iStrVal = BIZ_PGM_ID & "?txtMode="		& Parent.UID_M0006
		iStrVal = iStrVal & "&txtConPlant="		& .txtConPlant.value
		iStrVal = iStrVal & "&txtConFromDt="	& .txtConFromDt.Text
		iStrVal = iStrVal & "&txtConToDt="		& .txtConToDt.Text
		iStrVal = iStrVal & "&txtConMovType="	& .txtConMovType.value
		iStrVal = iStrVal & "&txtConShipToParty="	& .txtConShipToParty.value
		iStrVal = iStrVal & "&txtConSalesGrp="	& .txtConSalesGrp.value
		iStrVal = iStrVal & "&txtPromiseDt="	& .txtPromiseDt.Text
		iStrVal = iStrVal & "&txtTransMeth="	& .txtTransMeth.value
		
		' 작업유형 
		If .rdoWorkTypeReg.checked Then
			iStrVal = iStrVal & "&txtWorkType=C"
		Else
			iStrVal = iStrVal & "&txtWorkType=D"
		End If
		
		' 후속작업여부(출고처리)
		If .chkGIFlag.checked Then
			iStrVal = iStrVal & "&txtGiFlag=Y"
			iStrVal = iStrVal & "&txtActualGiDt="	& .txtActualGiDt.Text
			
			' 매출채권 
			If .chkArFlag.checked Then
				iStrVal = iStrVal & "&txtArFlag=Y"
			Else
				iStrVal = iStrVal & "&txtArFlag=N"
			End If
			
			' 세금계산서 
			If .chkVatFlag.checked Then
				iStrVal = iStrVal & "&txtVatFlag=Y"
			Else
				iStrVal = iStrVal & "&txtVatFlag=N"
			End If
		Else
			iStrVal = iStrVal & "&txtGiFlag=N"
		End If
		
		iStrVal = iStrVal & "&txtUserId="		& Parent.gUsrID
	End With

	If LayerShowHide(1) = False then
		Call BtnDisabled(0)
		Exit Function 
	End if

	Call RunMyBizASP(MyBizASP, iStrVal)	                                        '☜: 비지니스 ASP 를 가동 

	ExeReflect = True                                                           '⊙: Processing is NG
End Function

'=======================================================================================================
Function ExeReflectOk()				            '☆: 저장 성공후 실행 로직 
	Call DisplayMsgBox("990000","X","X","X")
End Function

'=======================================================================================================
Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
    Call DisplayMsgBox("800161","X","X","X")
End Function

'========================================
Sub rdoWorkTypeReg_OnClick()
	If Not lgBlnRegChecked Then
		lgBlnRegChecked = True
		Call ggoOper.SetReqAttr(frm1.txtPromiseDt,"N")
		Call ggoOper.SetReqAttr(frm1.txtTransMeth,"D")
		Call ggoOper.SetReqAttr(frm1.chkGIFlag,"D")
		frm1.btnTransMeth.disabled = False
	End If
End Sub

'========================================
Sub rdoWorkTypeDel_OnClick()
	If lgBlnRegChecked Then
		lgBlnRegChecked = False
		Call ggoOper.SetReqAttr(frm1.txtPromiseDt,"Q")
		Call ggoOper.SetReqAttr(frm1.txtTransMeth,"Q")
		Call ggoOper.SetReqAttr(frm1.chkGIFlag,"Q")
		
		frm1.btnTransMeth.disabled = True
		frm1.chkGIFlag.checked = False
		Call chkGIFlag_OnClick
	End If
End Sub

'   Event Desc : 공장 
'==========================================================================================
Function txtConPlant_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtConPlant.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("PT", "''", "S") & "", C_PopPlant) Then
				.txtConPlant.value = ""
				.txtConPlantNm.value = ""
				.txtConPlant.focus
			End If
			txtConPlant_OnChange = False
		Else
			.txtConPlantNm.value = ""
		End If
	End With
End Function

'   Event Desc : 출하형태 
'==========================================================================================
Function txtConMovType_OnChange()
	With frm1
		If Trim(.txtConMovType.value) <> "" Then
			If Not GetMovTypeInfo Then
				.txtConMovType.value = ""
				.txtConMovTypeNm.value = ""
				.txtConMovType.focus
			End If
			txtConMovType_OnChange = False
		Else
			.txtConMovTypeNm.value = ""
		End If
	End With
End Function

'   Event Desc : 납품처 
'==========================================================================================
Function txtConShipToParty_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtConShipToParty.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("SSH", "''", "S") & "", "default", "default", "default", "" & FilterVar("BF", "''", "S") & "", C_PopShipToParty) Then
				.txtConShipToParty.value = ""
				.txtConShipToPartyNm.value = ""
				.txtConShipToParty.focus
			End If
			txtConShipToParty_OnChange = False
		Else
			.txtConShipToPartyNm.value = ""
		End If
	End With
End Function

'   Event Desc : 영업그룹 
'==========================================================================================
Function txtConSalesGrp_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtConSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				.txtConSalesGrp.value = ""
				.txtConSalesGrpNm.value = ""
				.txtConSalesGrp.focus
			End If
			txtConSalesGrp_OnChange = False
		Else
			.txtConSalesGrpNm.value = ""
		End If
	End With
End Function

'   Event Desc : 운송방법 
'==========================================================================================
Function txtTransMeth_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtTransMeth.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "" & FilterVar("B9009", "''", "S") & "", "default", "default", "default", "" & FilterVar("MJ", "''", "S") & "", C_PopTransMeth) Then
				.txtTransMeth.value = ""
				.txtTransMethNm.value = ""
				.txtTransMeth.focus
			End If
			txtTransMeth_OnChange = False
		Else
			.txtTransMethNm.value = ""
		End If
	End With
End Function

'========================================
Sub chkGIFlag_OnClick()
	With frm1
		If .chkGIFlag.checked Then
			.chkArFlag.disabled = False
			.chkVatFlag.disabled = False
			Call ggoOper.SetReqAttr(.txtActualGiDt,"N")
		Else
			.chkArFlag.disabled = True
			.chkVatFlag.disabled = True
			Call ggoOper.SetReqAttr(.txtActualGiDt,"Q")
		End If
	End With
End Sub

'========================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConFromDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtConFromDt.focus
	End If
End Sub

'========================================
Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConToDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtConToDt.focus
	End If
End Sub

'========================================
Sub txtPromiseDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtPromiseDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtPromiseDt.focus
	End If
End Sub

'========================================
Sub txtActualGiDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtActualGiDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtActualGiDt.focus
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
 
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE WIDTH=100% CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>일괄출하등록</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* HEIGHT="right">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD CLASS="Tab11" VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>작업유형</TD>
								    <TD CLASS=TD6><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoWorkType" TAG="11X" VALUE="Y" CHECKED ID="rdoWorkTypeReg"><LABEL FOR="rdoWorkTypeReg">등록</LABEL>&nbsp;
								                  <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoWorkType" TAG="11X" VALUE="N" ID="rdoWorkTypeDel"><LABEL FOR="rdoWorkTypeDel">삭제</LABEL></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6><INPUT NAME="txtConPlant" TYPE="Text" Alt="공장" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopPlant">&nbsp;<INPUT NAME="txtConPlantNm" TYPE="Text" MAXLENGTH="20" Alt="공장명" SIZE=25 tag="14"></TD>									
									<TD CLASS="TD5" NOWRAP>출고예정일</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s4111ba1_fpDateTime1_txtConFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s4111ba1_fpDateTime2_txtConToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>출하형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConMovType" TYPE="Text" MAXLENGTH="3" SIZE=10 Alt="출하형태" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConMovType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopMovType">&nbsp;<INPUT NAME="txtConMovTypeNm" TYPE="Text" Alt="출하형태명" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>납품처</TD>
									<TD CLASS=TD6><INPUT NAME="txtConShipToParty" TYPE="Text" Alt="납품처" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConShipToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopShipToParty">&nbsp;<INPUT NAME="txtConShipToPartyNm" TYPE="Text" MAXLENGTH="20" Alt="납품처명" SIZE=25 tag="14"></TD>									
								</TR>
								<TR>	
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6><INPUT NAME="txtConSalesGrp" TYPE="Text" Alt="영업그룹" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSalesGrp">&nbsp;<INPUT NAME="txtConSalesGrpNm" TYPE="Text" Alt="영업그룹명" SIZE=25 tag="14"></TD>									
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>출고예정일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4111ba1_fpDateTime1_txtPromiseDt.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>운송방법</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtTransMeth" TYPE="Text" MAXLENGTH="5" SIZE=10 Alt="운송방법" tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTransMeth" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPopUp C_PopTransMeth">&nbsp;<INPUT NAME="txtTransMethNm" TYPE="Text" Alt="운송방법명" SIZE=25 tag="24"></TD> 									
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>후속작업여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=CHECKBOX NAME="chkGIFlag" tag="21" Class="Check"><LABEL ID="lblArFlag" FOR="chkArFlag">출고처리</LABEL>&nbsp;&nbsp;
										<INPUT TYPE=CHECKBOX NAME="chkArFlag" tag="21" Class="Check"><LABEL ID="lblArFlag" FOR="chkArFlag">매출채권</LABEL>&nbsp;&nbsp;
										<INPUT TYPE=CHECKBOX NAME="chkVatFlag" tag="21" Class="Check"><LABEL ID="lblVatFlag" FOR="chkVatFlag">세금계산서</LABEL>
									</TD>
									<TD CLASS=TD5 NOWRAP>실제출고일</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s4111ba1_fpDateTime1_txtActualGIDt.js'></script></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD VALIGN=TOP>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>실행</BUTTON></TD>
					<TD WIDTH=* Align=Right><a href = "VBSCRIPT:JumpChgCheck()">일괄출고처리</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
		        </TR>
		    </TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
