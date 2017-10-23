<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : ����ä���ϰ���� 
'*  3. Program ID           : S5111BA1
'*  4. Program Name         : 
'*  5. Program Desc         : ����ä�ǰ��� 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/06/30
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seongbae
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
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
<SCRIPT LANGUAGE=VBSCRIPT>		            '��: indicates that All variables must be declared in advance

Const BIZ_PGM_ID = "S5111bb1.asp"
Const BIZ_PGM_JUMP_ID = "s5114ma2"

Const C_PopMovType		= 1			' �������� 
Const C_PopShipToParty	= 2			' ��ǰó 
Const C_PopSalesGrp		= 3			' �����׷� 

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgBlnOpenedFlag
Dim lgBlnOpenPop			' Popup Window�� Open ���� 
Dim lgBlnRegChecked

Dim ToDateOfDB

ToDateOfDB = UNIConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat,parent.gDateFormat)

'=========================================
Sub InitVariables()
End Sub

'=========================================
Sub SetDefaultVal()
	With frm1
		If parent.gSalesGrp <> "" And Trim(.txtConSalesGrp.value) = "" Then
			.txtConSalesGrp.value = parent.gSalesGrp
			Call txtConSalesGrp_OnChange()
		End If

		.txtConFromDt.Text	= ToDateOfDB
		.txtConToDt.Text	= ToDateOfDB
		.txtBillDt.Text		= ToDateOfDB

		lgBlnRegChecked = True
		
		If GetTaxBillNoMgmtMeth = "M" Then
			.chkVatFlag.disabled = True	
		End If
		
		.txtConFromDt.Focus
	End With
End Sub	

'===========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"-->
	<% Call LoadInfTB19029A("I","*","NOCOOKIE", "BA") %>
End Sub

'===========================================
Function JumpChgCheck()
	Call PgmJump(BIZ_PGM_JUMP_ID)
End Function

'=========================================
Sub Form_Load()
    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
	Call SetDefaultVal
	Call SetToolbar("1000000000000111")										'��: ��ư ���� ���� 
	lgBlnOpenedFlag	 = True
End Sub
	
'=========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'=========================================
Function FncPrint()
	Call Parent.FncPrint()                                                      '��: Protect system from crashing
End Function

'=========================================
Function FncExcel() 
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'=========================================
Function FncFind() 
	Call Parent.FncFind(Parent.C_SINGLE, False)
End Function

'=========================================
Function FncExit()
    FncExit = True
End Function

'=========================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True
	
	With frm1
		Select Case pvIntWhere
			Case C_PopMovType	'�������� 
				iArrParam(0) = .txtConMovType.alt							
				iArrParam(1) = "dbo.B_MINOR MN "		
				iArrParam(2) = Trim(.txtConMovType.value)					
				iArrParam(3) = ""											
				iArrParam(4) = "MN.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND EXISTS (SELECT * FROM dbo.S_SO_TYPE_CONFIG ST WHERE	ST.MOV_TYPE = MN.MINOR_CD) "			
				
				iArrField(0) = "ED15" & Parent.gColSep & "MN.MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MN.MINOR_NM"
				
				iArrHeader(0) = .txtConMovType.alt							
				iArrHeader(1) = .txtConMovTypeNm.alt	
				
				frm1.txtConMovType.focus

			Case C_PopShipToParty	'��ǰó 
				iArrParam(1) = "dbo.B_BIZ_PARTNER BP INNER JOIN dbo.B_COUNTRY CT ON (CT.COUNTRY_CD = BP.CONTRY_CD)"								
				iArrParam(2) = Trim(.txtConShipToParty.value)			
				iArrParam(3) = ""											
				
				' ��ǰó Popup
				If lgBlnRegChecked Then
					iArrParam(0) = .txtConShipToParty.alt					
					iArrParam(4) = "BP.BP_TYPE IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND BP.usage_flag = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM B_BIZ_PARTNER_FTN BPF WHERE BPF.PARTNER_BP_CD = BP.BP_CD AND BPF.PARTNER_FTN = " & FilterVar("SSH", "''", "S") & ")" 

					iArrHeader(0) = .txtConShipToParty.alt					
					iArrHeader(1) = .txtConShipToPartyNm.alt					
				' �ֹ�ó Popup
				Else
					iArrParam(0) = "�ֹ�ó"
					iArrParam(4) = "BP.bp_type IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND BP.usage_flag = " & FilterVar("Y", "''", "S") & " "		' Where Condition
					
					iArrHeader(0) = "�ֹ�ó"
					iArrHeader(1) = "�ֹ�ó��"
				End If
				iArrHeader(2) = "����"
				iArrHeader(3) = "������"

				iArrField(0) = "ED15" & Parent.gColSep & "BP.BP_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "BP.BP_NM"
				iArrField(2) = "ED10" & Parent.gColSep & "BP.CONTRY_CD"
				iArrField(3) = "ED20" & Parent.gColSep & "CT.COUNTRY_NM"

				.txtConShipToParty.focus
			
			' �����׷� 
			Case C_PopSalesGrp												
			    iArrParam(0) = .txtConSalesGrp.Alt
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
 
	iArrParam(5) = iArrParam(0)							<%' �˾� ��Ī %> 

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False
	
	If iArrRet(0) <> "" Then
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function

'=========================================
Function SetConPopUp(ByVal pvArrRet,ByVal pvIntWhere)
	SetConPopUp = False

	With frm1
		Select Case pvIntWhere
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

'	Description : �ڵ尪�� �ش��ϴ� ���� Display�Ѵ�.
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
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' ���� Popup Display
		If err.number = 0 Then
			If lgBlnOpenedFlag Then
				GetCodeName = OpenConPopup(pvIntWhere)
			End If
		Else
			MsgBox Err.description, vbInformation,Parent.gLogoName
			Err.Clear
		End If
	End if
End Function

' ���ݰ�꼭 ��ȣ ���� ��� Fetch
'========================================
Function GetTaxBillNoMgmtMeth()
	On Error Resume Next
	
	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(5), iArrTemp
	
	iStrSelectList = " TOP 1 MINOR_CD "
	iStrFromList = " dbo.B_CONFIGURATION "
	iStrWhereList = "MAJOR_CD = " & FilterVar("S5001", "''", "S") & " AND SEQ_NO = 1 AND REFERENCE IS NOT NULL "
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, parent.gColSep)
		GetTaxBillNoMgmtMeth = iArrTemp(1)
	Else
		GetTaxBillNoMgmtMeth = "M"		' �������� ���� �������� ó�� 
		
		If Err.number <> 0 Then
			MsgBox Err.description, vbInformation,Parent.gLogoName
			Err.Clear
		End If
	End if
End Function

'	Description : �������¿� ���� Description Fetch
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
		' ���� Popup Display
		If err.number = 0 Then
			GetMovTypeInfo = OpenConPopup(C_PopMovType)
		Else
			MsgBox Err.description, vbInformation,Parent.gLogoName
			Err.Clear
		End If
	End if
End Function

'========================================
Function ExeReflect() 
	Call BtnDisabled(1)
	Dim iStrVal

	ExeReflect = False                                                          '��: Processing is NG
    
	On Error Resume Next                                                   '��: Protect system from crashing

	If Not chkField(Document, "1") Or Not chkField(Document, "2") Then
		Call BtnDisabled(0)
		Exit Function
	End If

	With frm1
		If Not ValidDateCheck(.txtConFromDt, .txtConToDt) Then
			Call BtnDisabled(0)
			Exit Function
		End If

		If Not ValidDateCheck(.txtConToDt, .txtBillDt) Then
			Call BtnDisabled(0)
			Exit Function
		End If
	
		If DisplayMsgBox("900018",Parent.VB_YES_NO,"X","X") = vbNo Then
			Call BtnDisabled(0)
			Exit Function
		End If

		iStrVal = BIZ_PGM_ID & "?txtMode="		& Parent.UID_M0006
		iStrVal = iStrVal & "&txtConFromDt="	& .txtConFromDt.Text
		iStrVal = iStrVal & "&txtConToDt="		& .txtConToDt.Text
		iStrVal = iStrVal & "&txtConMovType="	& .txtConMovType.value
		iStrVal = iStrVal & "&txtConShipToParty="	& .txtConShipToParty.value
		iStrVal = iStrVal & "&txtConSalesGrp="	& .txtConSalesGrp.value
		iStrVal = iStrVal & "&txtBillDt="		& .txtBillDt.Text
		iStrVal = iStrVal & "&txtUserId="		& Parent.gUsrID

		' ��� 
		If .rdoWorkTypeReg.checked Then
			iStrVal = iStrVal & "&txtWorkType=C"

			' ����ä��Ȯ�� 
			If .chkArFlag.checked Then
				iStrVal = iStrVal & "&txtArFlag=Y"
			Else
				iStrVal = iStrVal & "&txtArFlag=N"
			End If
				
			' ���ݰ�꼭 
			If .chkVatFlag.checked Then
				iStrVal = iStrVal & "&txtVatFlag=Y"
			Else
				iStrVal = iStrVal & "&txtVatFlag=N"
			End If
		' ���� 
		Else
			iStrVal = iStrVal & "&txtWorkType=D"
			iStrVal = iStrVal & "&txtArFlag=N"
		End If

	End With

	If LayerShowHide(1) = False then
		Call BtnDisabled(0)
		Exit Function 
	End if

	Call RunMyBizASP(MyBizASP, iStrVal)	                                        '��: �����Ͻ� ASP �� ���� 

	ExeReflect = True                                                           '��: Processing is NG
End Function

'========================================
Function ExeReflectOk()				            '��: ���� ������ ���� ���� 
	Call DisplayMsgBox("990000","X","X","X")
End Function

'========================================
Function ExeReflectNo()				            '��: ����� �ڷᰡ �����ϴ� 
    Call DisplayMsgBox("800161","X","X","X")
End Function

'========================================
Sub rdoWorkTypeReg_OnClick()
	If Not lgBlnRegChecked Then
		lgBlnRegChecked = True
		idDateTitle.innerHTML = "���������"
		idBpTitle.innerHTML = "��ǰó"
		Call ggoOper.SetReqAttr(frm1.txtConMovType,"D")
		frm1.btnMovType.disabled = False

		Call ggoOper.SetReqAttr(frm1.txtBillDt,"N")
		Call ggoOper.SetReqAttr(frm1.chkArFlag,"D")
		Call ggoOper.SetReqAttr(frm1.chkVatFlag,"D")
	End If
End Sub

'========================================
Sub rdoWorkTypeDel_OnClick()
	If lgBlnRegChecked Then
		lgBlnRegChecked = False
		idDateTitle.innerHTML = "����ä����"
		idBpTitle.innerHTML = "�ֹ�ó"
		Call ggoOper.SetReqAttr(frm1.txtConMovType,"Q")
		frm1.btnMovType.disabled = True

		Call ggoOper.SetReqAttr(frm1.txtBillDt,"Q")
		Call ggoOper.SetReqAttr(frm1.chkArFlag,"Q")
		Call ggoOper.SetReqAttr(frm1.chkVatFlag,"Q")
		
	End If
End Sub

'   Event Desc : �������� 
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

'   Event Desc : ��ǰó 
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

'   Event Desc : �����׷� 
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
Sub txtBillDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBillDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtBillDt.focus
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>����ä���ϰ����</font></td>
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
									<TD CLASS=TD5 NOWRAP>�۾�����</TD>
								    <TD CLASS=TD6><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoWorkType" TAG="11X" VALUE="Y" CHECKED ID="rdoWorkTypeReg"><LABEL FOR="rdoWorkTypeReg">���</LABEL>&nbsp;
								                  <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoWorkType" TAG="11X" VALUE="N" ID="rdoWorkTypeDel"><LABEL FOR="rdoWorkTypeDel">����</LABEL></TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" ID="idDateTitle" NOWRAP>���������</TD>
									<TD CLASS="TD6" NOWRAP>
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s5111ba1_fpDateTime1_txtConFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s5111ba1_fpDateTime2_txtConToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD CLASS=TD5 NOWRAP>��������</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtConMovType" TYPE="Text" MAXLENGTH="3" SIZE=10 Alt="��������" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMovType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopMovType">&nbsp;<INPUT NAME="txtConMovTypeNm" TYPE="Text" Alt="�������¸�" SIZE=25 tag="14"></TD>
								</TR>
									<TD CLASS=TD5 ID="idBpTitle" NOWRAP>��ǰó</TD>
									<TD CLASS=TD6><INPUT NAME="txtConShipToParty" TYPE="Text" Alt="��ǰó" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnShipToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopShipToParty">&nbsp;<INPUT NAME="txtConShipToPartyNm" TYPE="Text" MAXLENGTH="20" Alt="��ǰó��" SIZE=25 tag="14"></TD>									
									<TD CLASS=TD5 NOWRAP>�����׷�</TD>
									<TD CLASS=TD6><INPUT NAME="txtConSalesGrp" TYPE="Text" Alt="�����׷�" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSalesGrp">&nbsp;<INPUT NAME="txtConSalesGrpNm" TYPE="Text" Alt="�����׷��" SIZE=25 tag="14"></TD>
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
									<TD CLASS=TD5 NOWRAP>����ä����</TD>
									<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s5111ba1_fpDateTime3_txtBillDt.js'></script>
									<TD CLASS=TD5 NOWRAP>�ļ��۾�����</TD>
									<TD CLASS=TD6 NOWRAP>
										<INPUT TYPE=CHECKBOX NAME="chkArFlag" tag="21" Class="Check"><LABEL ID="lblArFlag" FOR="chkArFlag">Ȯ��</LABEL>&nbsp;&nbsp;
										<INPUT TYPE=CHECKBOX NAME="chkVatFlag" tag="21" Class="Check"><LABEL ID="lblVatFlag" FOR="chkVatFlag">���ݰ�꼭</LABEL>
									</TD>
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
					<TD><BUTTON NAME="btnExe" CLASS="CLSSBTN" onclick="ExeReflect()" Flag=1>����</BUTTON></TD>
					<TD WIDTH=* Align=Right><a href = "VBSCRIPT:JumpChgCheck()">����ä���ϰ�Ȯ��</a></TD>
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
