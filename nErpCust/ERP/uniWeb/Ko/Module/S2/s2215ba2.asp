<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 품목판매계획 일별 배분 
'*  3. Program ID           : S2215BA2
'*  4. Program Name         : 
'*  5. Program Desc         : 판매계획관리 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/01/15
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%
Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "S","NOCOOKIE","QB")

Call SubOpenDB(lgObjConn)

Dim lgLngUseStep, lgLngProcessByPlant

lgStrSql = "SELECT USE_STEP, PROCESS_BY_PLANT FROM dbo.S_SP_CONFIG WHERE SP_TYPE = " & FilterVar("E", "''", "S") & " " 

If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") Then
	lgLngUseStep = CLng(lgObjRs("USE_STEP"))
	lgLngProcessByPlant = CLng(lgObjRs("PROCESS_BY_PLANT"))
Else
    'If data not exists
    lgLngUseStep = -1
    lgLngProcessByPlant = -1
End If
lgObjRs.Close
lgObjConn.Close
Set lgObjRs = Nothing
Set lgObjConn = Nothing
%>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>		            '☜: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "S2215bb2.asp"
CONST BIZ_JUMP_ID_S2216MA1 = "S2216MA1"			' 공장별일별품목판매계획조정 

Const C_PopSalesGrp	= 1
Const C_PopPlantCd	= 1

<% If (lgLngProcessByPlant And 1024) > 0 Then %>
Const C_StrGrpFlag = "N"
<% Else %>
Const C_StrGrpFlag = "Y"
<% End If%>

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop          
Dim lgBlnOpenedFlag
Dim lgBlnCfmChecked

'========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

'========================================================================================================
Sub SetDefaultVal()
	' 공장별로 확정 처리할 경우 
	<% If (lgLngProcessByPlant And 1024) > 0 Then %>
	If parent.gPlant <> "" And Trim(frm1.txtPlantCd.value) = "" Then
		frm1.txtPlantCd.value = parent.gPlant
		frm1.txtPlantCd.focus 
		Call txtPlantCd_OnChange()
	End If
	<% Else %>
	If parent.gSalesGrp <> "" And Trim(frm1.txtSalesGrp.value) = "" Then
		frm1.txtSalesGrp.value = parent.gSalesGrp
		frm1.txtSalesGrp.focus 
		Call txtSalesGrp_OnChange()
	End If
	<% End If%>
	
    lgBlnCfmChecked = True
End Sub	

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("Q","H","NOCOOKIE", "BA") %>
End Sub

'==========================================================================================================
Function CookiePage(Byval pvKubun)

	On Error Resume Next
	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>
	Dim iStrTemp, iArrVal
	
	With frm1
		If pvKubun = 1 Then
			<% If (lgLngProcessByPlant And 1024) > 0 Then %>
			WriteCookie CookieSplit , "" & Parent.gColSep & "" & Parent.gColSep & _
									  .txtPlantCd.value & Parent.gColSep & .txtPlantNm.value & Parent.gColSep & _
									  .txtCfmFrSPPeriod.value& Parent.gColSep & .txtCfmFrSPPeriodDesc.value & Parent.gColSep & _
									  .txtFcToSPPeriod.value & Parent.gColSep & .txtFcToSPPeriodDesc.value
			<% Else %>
			WriteCookie CookieSplit , .txtSalesGrp.value & Parent.gColSep & .txtSalesGrpNm.value & Parent.gColSep & _
									  "" & Parent.gColSep & "" & Parent.gColSep & _
									  .txtCfmFrSPPeriod.value& Parent.gColSep & .txtCfmFrSPPeriodDesc.value & Parent.gColSep & _
									  .txtFcToSPPeriod.value & Parent.gColSep & .txtFcToSPPeriodDesc.value
			<% End If %>									  
		' Load시 
		ElseIf pvKubun = 0 Then
			iStrTemp = ReadCookie(CookieSplit)
			If Trim(Replace(iStrTemp, Parent.gColSep, "")) = "" then Exit Function
			iArrVal = Split(iStrTemp, Parent.gColSep)
			<% If (lgLngProcessByPlant And 1024) > 0 Then %>
			.txtPlantCd.value = iArrVal(2)
			.txtPlantNm.value = iArrVal(3)
			<% Else %>
			.txtSalesGrp.value	 = iArrVal(0)
			.txtSalesGrpNm.value = iArrVal(1)
			<% End If %>	
			WriteCookie CookieSplit , ""
			Call GetCfmPeriod(0)
		End If
	End With
End Function

'==========================================================================================================
Function JumpChgCheck(byVal pvStrJumpPgmId)
	Call CookiePage(1)
	Call PgmJump(pvStrJumpPgmId)
End Function

'========================================================================================================
Sub MakeKeyStream(pOpt)
   
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   If pOpt = "Q" Then
      lgKeyStream = Frm1.txtWarrentNo.Value & Parent.gColSep       'You Must append one character(Parent.gColSep)
   Else
      lgKeyStream = Frm1.txtMajorCd.Value & Parent.gColSep         'You Must append one character(Parent.gColSep)
   End If   

   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        
	
'========================================================================================================
Sub Form_Load()
    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

	Call CookiePage(0)
	Call InitVariables                                                     '⊙: Setup the Spread sheet
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'⊙: 버튼 툴바 제어 
	<%If lgLngUseStep = -1 Then%>
	'판매계획환경설정 정보가 없습니다.
	Call DisplayMsgBox("202403", "X", "", "")
	Call BtnDisabled(1)
	' 공장별 일별 품목 판매계획조정을 사용하지 않을 경우 
	<%ElseIf (lgLngUseStep And 4096) = 0 Then %>
	Call DisplayMsgBox("202415", "X", "", "")
	Call BtnDisabled(1)
	<%End If%>

	lgBlnOpenedflag = True
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
	
	Select Case pvIntWhere
	<% If (lgLngProcessByPlant And 1024) > 0 Then %>
	Case C_PopSalesGrp												
		iArrParam(1) = "B_PLANT"						<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtPlantCd.value)		<%' Code Condition%>
		iArrParam(3) = ""								<%' Name Cindition%>
		iArrParam(4) = ""								<%' Where Condition%>
		iArrParam(5) = "공장"						<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & Parent.gColSep & "PLANT_CD"						<%' Field명(0)%>
		iArrField(1) = "ED30" & Parent.gColSep & "PLANT_NM"					<%' Field명(1)%>
    
	    iArrHeader(0) = "공장"						<%' Header명(0)%>
	    iArrHeader(1) = "공장명"					<%' Header명(1)%>

		frm1.txtPlantCd.focus 
	<% Else %>
	Case C_PopSalesGrp												
		iArrParam(1) = "B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtSalesGrp.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = "영업그룹"
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
	    iArrHeader(0) = "영업그룹"
	    iArrHeader(1) = "영업그룹명"

		frm1.txtSalesGrp.focus 
	<% End If%>
	End Select
 
	iArrParam(0) = iArrParam(5)							<%' 팝업 명칭 %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

'========================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		<% If (lgLngProcessByPlant And 1024) > 0 Then %>
		Case C_PopPlantCd
			.txtPlantCd.value = pvArrRet(0) 
			.txtPlantNm.value = pvArrRet(1)
			If .rdoWorkTypeDistr.checked Then
				Call GetCfmPeriod(0)
			Else
				Call GetCancelPeriod()
			End If
		<% Else %>	
		Case C_PopSalesGrp
			.txtSalesGrp.value = pvArrRet(0) 
			.txtSalesGrpNm.value = pvArrRet(1)
			If .rdoWorkTypeDistr.checked Then
				Call GetCfmPeriod(0)
			Else
				Call GetCancelPeriod()
			End If
		<% End If%>
		End Select
	End With
	
	SetConPopup = True

End Function

'========================================================================================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(5), iArrTemp
	
	GetCodeName = False
	
	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		iArrRs(5) = iArrTemp(3)				' 계획기간 순번 
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		If err.number = 0 Then
			If lgBlnOpenedFlag Then
				GetCodeName = OpenConPopup(pvIntWhere)
			End If
		Else
			MsgBox err.Description
		End If
	End if
End Function

<%'======================================   GetCfmPeriod()  =====================================
'	Description : 확정기간 Fetch
'====================================================================================================
%>
Function GetCfmPeriod(ByVal pvIntSpPeriodSeq)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrCode
	Dim iStrRs
	Dim iArrRs
	
	GetCfmPeriod = False

	If C_StrGrpFlag = "Y" Then
		iStrCode = Trim(frm1.txtSalesGrp.value)
	Else
		iStrCode = Trim(frm1.txtPlantCd.value)
	End If

	If iStrCode = "" Then Exit Function
	
	iStrSelectList = " * "
	iStrFromList = "  dbo.ufn_s_GetCfmPeriod(" & FilterVar("S2215BA2", "''", "S") & ",  " & FilterVar(iStrCode, "''", "S") & ", " & FilterVar("1", "''", "S") & " , " & FilterVar("E", "''", "S") & " ,  " & FilterVar(C_StrGrpFlag, "''", "S") & ", " & pvIntSpPeriodSeq & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrRs = Split(iStrRs, Chr(11))
		With frm1
			.txtCfmFrSPPeriod.value = iArrRs(1)
			.txtCfmFrSPPeriodDesc.value = iArrRs(2)
			.txtCfmToSPPeriod.value = iArrRs(3)
			.txtCfmToSPPeriodDesc.value = iArrRs(4)
			.txtFcToSPPeriod.value = iArrRs(5)
			.txtFcToSPPeriodDesc.value = iArrRs(6)
		End With
		
		GetCfmPeriod = True
	Else
		With frm1
			.txtCfmFrSPPeriod.value = ""
			.txtCfmFrSPPeriodDesc.value = ""
			.txtCfmToSPPeriod.value = ""
			.txtCfmToSPPeriodDesc.value = ""
			.txtFcToSPPeriod.value = ""
			.txtFcToSPPeriodDesc.value = ""
		End With
	End if
End Function

<%'======================================   GetCancelPeriod()  =====================================
'	Description : 확정기간 Fetch
'====================================================================================================
%>
Function GetCancelPeriod()

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrCode
	Dim iStrRs
	Dim iArrRs
	
	GetCancelPeriod = False

	If C_StrGrpFlag = "Y" Then
		iStrCode = Trim(frm1.txtSalesGrp.value)
	Else
		iStrCode = Trim(frm1.txtPlantCd.value)
	End If
	
	If iStrCode = "" Then Exit Function

	iStrSelectList = " * "
	iStrFromList = "  dbo.ufn_s_GetCancelPeriod(" & FilterVar("S2215BA2", "''", "S") & ",  " & FilterVar(iStrCode, "''", "S") & ", " & FilterVar("1", "''", "S") & " , " & FilterVar("E", "''", "S") & " ,  " & FilterVar(C_StrGrpFlag, "''", "S") & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrRs = Split(iStrRs, Chr(11))
		With frm1
			.txtCfmFrSPPeriod.value = iArrRs(1)
			.txtCfmFrSPPeriodDesc.value = iArrRs(2)
			.txtCfmToSPPeriod.value = ""
			.txtCfmToSPPeriodDesc.value = ""
			.txtFcToSPPeriod.value = ""
			.txtFcToSPPeriodDesc.value = ""
		End With
		
		GetCancelPeriod = True
	Else
		With frm1
			.txtCfmFrSPPeriod.value = ""
			.txtCfmFrSPPeriodDesc.value = ""
			.txtCfmToSPPeriod.value = ""
			.txtCfmToSPPeriodDesc.value = ""
			.txtFcToSPPeriod.value = ""
			.txtFcToSPPeriodDesc.value = ""
		End With
	End if
End Function

'=======================================================================================================
Function ExeReflect() 
	Call BtnDisabled(1)
	Dim iStrVal

	ExeReflect = False                                                          '⊙: Processing is NG
    
	On Error Resume Next                                                   '☜: Protect system from crashing

	If Not chkField(Document, "1") Then
		Call BtnDisabled(0)
		Exit Function
	End If

	If DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X") = vbNo Then
		Call BtnDisabled(0)
		Exit Function
	End If

	With frm1
		iStrVal = BIZ_PGM_ID & "?txtMode="		& Parent.UID_M0006
		If .rdoWorkTypeDistr.checked Then
			iStrVal = iStrVal     & "&txtWorkType=Y"		' 확정 

			If Trim(.txtCfmToSPPeriod.value) = "" Or Trim(.txtFcToSPPeriod.value) = "" Then
				Call BtnDisabled(0)
				Call DisplayMsgBox("970029","X",.txtCfmFrSPPeriod.alt,"X")
				Exit function
			End If
		Else
			iStrVal = iStrVal     & "&txtWorkType=N"		' 취소 
		End If
		If C_StrGrpFlag = "Y" Then
			iStrVal = iStrVal     & "&txtSalesGrp="	& .txtSalesGrp.value
			iStrVal = iStrVal     & "&txtGrpFlag=Y"
		Else
			iStrVal = iStrVal     & "&txtSalesGrp="	& .txtPlantCd.value
			iStrVal = iStrVal     & "&txtGrpFlag=N"
		End If
		iStrVal = iStrVal     & "&txtFrSpPeriod="	& .txtCfmFrSPPeriod.value
		iStrVal = iStrVal     & "&txtToSpPeriod=" & .txtCfmToSPPeriod.value
		iStrVal = iStrVal     & "&txtFcSpPeriod="	& .txtFcToSPPeriod.value
		
		iStrVal = iStrVal & "&txtUserId=" & Parent.gUsrID
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

Function ExeReflectNo()				            '☆: 실행된 자료가 없습니다 
    Call DisplayMsgBox("800161","X","X","X")

End Function

'========================================================================================================
' 공장별로 확정 처리할 경우 
<% If (lgLngProcessByPlant And 1024) > 0 Then %>
' OnChange - 공장 
Function txtPlantCd_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtPlantCd.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("PT", "''", "S") & "", C_PopPlantCd) Then
				.txtPlantCd.value = ""
				.txtPlantNm.value = ""
				.txtPlantCd.focus
			End If
			txtPlantCd_OnChange = False
		Else
			.txtPlantNm.value = ""
		End If
	End With
End Function
<% Else %>
' OnChange - 영업그룹 
Function txtSalesGrp_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtSalesGrp.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", "default", "" & FilterVar("SG", "''", "S") & "", C_PopSalesGrp) Then
				.txtSalesGrp.value = ""
				.txtSalesGrpNm.value = ""
				.txtSalesGrp.focus
			End If
			txtSalesGrp_OnChange = False
		Else
			.txtSalesGrpNm.value = ""
		End If
	End With
End Function
<% End If %>

' _OnClick
Sub rdoWorkTypeDistr_OnClick()
	If Not lgBlnCfmChecked Then
		lgBlnCfmChecked = True
		frm1.btnExe.value = "배분"
		Call GetCfmPeriod(0)
	End If
End Sub

Sub rdoWorkTypeCancel_OnClick()
	If lgBlnCfmChecked Then
		lgBlnCfmChecked = False
		frm1.btnExe.value = "취소"
		Call GetCancelPeriod()
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
 
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB4" CELLSPACING=0 CELLPADDING=0>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>품목판매계획일별배분</font></td>
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
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=20>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS=TD5 NOWRAP>작업유형</TD>
							    <TD CLASS=TD6 NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoWorkType" TAG="11X" VALUE="Y" CHECKED ID="rdoWorkTypeDistr"><LABEL FOR="rdoWorkTypeDistr">배분</LABEL>&nbsp;
							                         <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoWorkType" TAG="11X" VALUE="N" ID="rdoWorkTypeCancel"><LABEL FOR="rdoWorkTypeCancel">취소</LABEL></TD>
							</TR>
							<TR>
								<% If (lgLngProcessByPlant And 1024) > 0 Then %>
								<TD CLASS=TD5 NOWRAP>공장</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPlantCd" ALT="공장" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopPlantCd">&nbsp;<INPUT NAME="txtPlantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								<% Else %>
								<TD CLASS=TD5 NOWRAP>영업그룹</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrp" ALT="영업그룹" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSalesGrp">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								<% End If %>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>확정기간</TD>
								<TD CLASS="TD6"><INPUT NAME="txtCfmFrSPPeriod" ALT="확정기간" TYPE="Text" MAXLENGTH=8 SiZE=10 tag="14XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCfmFrSPPeriod" align=top TYPE="BUTTON" >&nbsp;<INPUT NAME="txtCfmFrSPPeriodDesc" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14">&nbsp;~&nbsp;
												<INPUT NAME="txtCfmToSPPeriod" ALT="확정기간" TYPE="Text" MAXLENGTH=8 SiZE=10 tag="14XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCfmToSPPeriod" align=top TYPE="BUTTON" >&nbsp;<INPUT NAME="txtCfmToSPPeriodDesc" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>예시기간</TD>
								<TD CLASS="TD6"><INPUT NAME="txtFcToSPPeriod" ALT="예시기간" TYPE="Text" MAXLENGTH=8 SiZE=10 tag="14XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnFcToSPPeriod" align=top TYPE="BUTTON" >&nbsp;<INPUT NAME="txtFcToSPPeriodDesc" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
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
		<TD>
		    <TABLE <%=LR_SPACE_TYPE_30%>>
		        <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnExe" CLASS="CLSMBTN" onclick="ExeReflect()" Flag=1>배분</BUTTON></TD>
					<TD WIDTH=* ALIGN="right">
						<% If (lgLngUseStep And 4096) > 0 Then %>
						<a href = "vbscript:JumpChgCheck(BIZ_JUMP_ID_S2216MA1)">공장별일별품목판매계획조정</a>
						<% End If %></TD>
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
