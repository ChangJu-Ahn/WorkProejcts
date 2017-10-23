<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : S3113PA1
'*  4. Program Name         : CTP Check
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/09/27
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Cho Song-Hyon
'* 10. Modifier (Last)      : Cho Song-Hyon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit					

Dim arrParent
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)


<!-- #Include file="../../inc/lgvariables.inc" -->

	Const BIZ_PGM_ID = "s3113pb1.asp"			

	Const C_SoNo = 0

	Const CTPQuery = "CTPQuery"
	Const CTPAccept = "CTPAccept"
	Const CTPModify = "CTPModify"
	Const CTPCancel = "CTPCancel"

	Const Div1 = "create"
	Const Div2 = "combination"
	Const Div3 = "division"

	Dim arrReturn					

	Dim gblnWinEvent				
	Dim intRowCnt

	Dim arrParam
	Dim arrSoNo
	Dim arrGridCount	
	
	arrParam = arrParent(1)
	arrSoNo = arrParent(2)
	arrGridCount = arrParent(3)
	
'================================================================================================================
Function InitVariables()
	lgIntGrpCount = 0										
	lgStrPrevKey = ""										
		
	<% '------ Coding part ------ %>
	gblnWinEvent = False

	intRowCnt = 1

	frm1.txtCtpSeq.value = 0
	frm1.txtSoNo.value = arrSoNo(C_SoNo)
	frm1.txtItemCode.value = arrParam(0,0)
	frm1.txtItemName.value = arrParam(0,1)
	frm1.txtSoSeq.value = arrParam(0,2)
	frm1.txtReqDate.value = arrParam(0,3)
	frm1.txtPlantCd.value = arrParam(0,4)
	frm1.txtReqQty.value = UNIFormatNumber(arrParam(0,5),PopupParent.ggQty.DecPoint,-2,0,PopupParent.ggQty.RndPolicy,PopupParent.ggQty.RndUnit)
	frm1.txtTrackingNO.value = arrParam(0,6)
	frm1.txtAPSHost.value = arrParam(0,7)
	frm1.txtAPSPort.value = arrParam(0,8)
	frm1.txtCTPTimes.value = arrParam(0,9)

	Self.Returnvalue = ""

	frm1.txtTodayDate.value = EndDate

	frm1.btnCTPSave.Disabled = True
	frm1.btnCTPCancel.Disabled = True
	frm1.btnSave.Disabled = True

	frm1.txtExitFlag.value = ""

End Function
	
'================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'================================================================================================================
Function CancelClick()

Dim Answer
	
	With frm1
		<% ' CTP Accept/Modify를 수행한 경우 %> 
		If Len(Trim(.txtAfterChangeDate.value)) Then

			Answer = DisplayMsgBox("900016", VB_YES_NO, "X", "X")                
			If Answer = vbNo Then
				Exit Function
			End If

			Call DbCTPCancel()
			Exit Function

		End If

		Call CancelClickOK()

	End With

End Function

'================================================================================================================
Function CancelClickOK()

	If Trim(Self.Returnvalue) = "" Then Self.Returnvalue = "Cancel"
	Self.Close()
		
End Function

'================================================================================================================
Function DisplayFlag(strDispaly)

	Select Case strDispaly
	Case Div1
		divCreate.style.display = ""
		divCom.style.display = "None"
		divDiv.style.display = "None"

		frm1.txtReqDate.readOnly = True
		frm1.txtReqQty.readOnly = True
		frm1.txtAccDate_All.readOnly = True
		frm1.txtAccQty_All.readOnly = True
		frm1.txtAccDate_Sub1.readOnly = True
		frm1.txtAccQty_Sub1.readOnly = True
		frm1.txtAccDate_Sub2.readOnly = True
		frm1.txtAccQty_Sub2.readOnly = True

	Case Div2
		divCom.style.display = ""
		divCreate.style.display = "None"
		divDiv.style.display = "None"

		frm1.txtAccDate_All_Com.readOnly = True
		frm1.txtAccQty_All_Com.readOnly = True

	Case Div3
		divDiv.style.display = ""
		divCreate.style.display = "None"
		divCom.style.display = "None"

		frm1.txtAccDate_Sub1_Div.readOnly = True
		frm1.txtAccQty_Sub1_Div.readOnly = True
		frm1.txtAccDate_Sub2_Div.readOnly = True
		frm1.txtAccQty_Sub2_Div.readOnly = True

	End Select

	'---Button Value / Disable Control
	If UNICDBL(frm1.txtCtpSeq.value) > 0 Then
		frm1.btnCTPSave.Value = "CTP Modify"
		frm1.btnCTPSave.Disabled = False
		frm1.btnCTPCancel.Disabled = False
	Else
		frm1.btnCTPSave.Value = "CTP Accept"
		frm1.btnCTPSave.Disabled = False
		frm1.btnCTPCancel.Disabled = True
	End If

End Function

'================================================================================================================

Function CTPKaraCalc()

	Dim strReqQty, strAcctQty

	strReqQty = Round(UNICDbl(frm1.txtReqQty.value) * 0.6)
	strAcctQty = UNICDbl(frm1.txtReqQty.value) - UNICDbl(strReqQty)

	<% '요청한 일자(납기일)보다 무조건 +10일 %>
	frm1.txtAccDate_All.value = UnIDateAdd("d",10,frm1.txtReqDate.value, PopupParent.gDateFormat)
	frm1.txtAccQty_All.value = UNIFormatNumber(frm1.txtReqQty.value, PopupParent.ggQty.DecPoint, -2, 0,PopupParent.ggQty.RndPolicy,PopupParent.ggQty.RndUnit)

	frm1.txtAccDate_Sub1.value = frm1.txtReqDate.value
	frm1.txtAccQty_Sub1.value = UNIFormatNumber(strReqQty, PopupParent.ggQty.DecPoint, -2, 0,PopupParent.ggQty.RndPolicy,PopupParent.ggQty.RndUnit)

	frm1.txtAccDate_Sub2.value = UnIDateAdd("d",10,frm1.txtReqDate.value, PopupParent.gDateFormat)
	frm1.txtAccQty_Sub2.value = UNIFormatNumber(strAcctQty, PopupParent.ggQty.DecPoint, -2, 0,PopupParent.ggQty.RndPolicy,PopupParent.ggQty.RndUnit)

End Function

'================================================================================================================
Function CTPFixedDateTime()

	CTPFixedDateTime = False

	'Order Entry Date
	'=If Len(Trim(frm1.txtTodayDate.value)) Then frm1.txtTodayDate.value = FormatDateTime(Trim(frm1.txtTodayDate.value) & #23:59:59#,vbGeneralDate)
	If Len(Trim(frm1.txtTodayDate.value)) Then frm1.txtTodayDate.value = Trim(frm1.txtTodayDate.value) & #23:59:59#

	'Order Request Date
	If Len(Trim(frm1.txtReqDate.value)) Then frm1.txtReqDate.value = Trim(frm1.txtReqDate.value) & #23:59:59#

	'Order Promise Date
	If Len(Trim(frm1.txtAccDate_All.value)) Then frm1.txtAccDate_All.value = Trim(frm1.txtAccDate_All.value) & #23:59:59#
	If Len(Trim(frm1.txtAccDate_Sub1.value)) Then frm1.txtAccDate_Sub1.value = Trim(frm1.txtAccDate_Sub1.value) & #23:59:59#
	If Len(Trim(frm1.txtAccDate_Sub2.value)) Then frm1.txtAccDate_Sub2.value = Trim(frm1.txtAccDate_Sub2.value) & #23:59:59#

	CTPFixedDateTime = True

End Function

'================================================================================================================

Sub Form_Load()

	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")
	Call InitVariables
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call DbQuery()

End Sub

'================================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

	'If Trim(frm1.btnCTPSave.Value) = "CTP Modify" Then
	If UNICDBL(frm1.txtCtpSeq.value) > 0 Then
		CancelClickOK()
		Exit Sub
	End If
		
	If Trim(frm1.txtExitFlag.value) = "" And Len(Trim(frm1.txtAfterChangeDate.value)) Then Call DbCTPCancel()

End Sub
	
'================================================================================================================
Sub btnCTPSave_OnClick()
	Call DbCTPSave()
End Sub

'================================================================================================================
Sub btnCTPCancel_OnClick()
	Call DbCTPCancel()
End Sub

'================================================================================================================
Sub btnSave_OnClick()
	Dim Answer
	
    
	Answer = DisplayMsgBox("900018", VB_YES_NO, "X", "X")	
	If Answer = VBNO Then Exit Sub

	Call DbSave()

End Sub

'================================================================================================================
Function DbQuery() 
    
    Err.Clear                                                               
    
    DbQuery = False                                                         

	If LayerShowHide(1) = False Then
		Exit Function
	End If
	    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001							
    strVal = strVal & "&txtSoNo=" & Trim(frm1.txtSoNo.value)
    strVal = strVal & "&txtSoSeq=" & Trim(frm1.txtSoSeq.value)
    strVal = strVal & "&txtItemCode=" & Trim(frm1.txtItemCode.value)        
    strVal = strVal & "&txtTrackingNO=" & Trim(frm1.txtTrackingNO.value)
    strVal = strVal & "&txtAPSHost=" & Trim(frm1.txtAPSHost.value)
    strVal = strVal & "&txtAPSPort=" & Trim(frm1.txtAPSPort.value)

	Call RunMyBizASP(MyBizASP, strVal)										

    DbQuery = True                                                          

End Function

'================================================================================================================
Function DbSave() 

    Err.Clear																

	DbSave = False															

	If frm1.rdoSelect_All.checked = True Then
		frm1.txtRadioFlg.value = frm1.rdoSelect_All.value
	ElseIf frm1.rdoSelect_Sub.checked = True Then
		frm1.txtRadioFlg.value = frm1.rdoSelect_Sub.value 
	End If

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	frm1.txtMode.value = PopupParent.UID_M0002											<%'☜: 비지니스 처리 ASP 의 상태 %>
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
    DbSave = True                                                           
    
End Function

'================================================================================================================
Function DbSaveOk()															
	MsgBox "CTP 저장이 완료되었습니다.", vbInformation, "<%=gLogoName%>"
	Self.Returnvalue = "Save"
	Self.Close()
End Function

'================================================================================================================
Function DbCTPQuery() 
   
    Err.Clear                                                               

	'If CTPFixedDateTime = False Then Exit Function

    DbCTPQuery = False                                                         

	Call LayerShowHide(1)
	    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & CTPQuery							
    strVal = strVal & "&txtReqDate=" & Trim(frm1.txtReqDate.value)			
    strVal = strVal & "&txtReqQty=" & Trim(frm1.txtReqQty.value)
    strVal = strVal & "&txtSoNo=" & Trim(frm1.txtSoNo.value)
    strVal = strVal & "&txtSoSeq=" & Trim(frm1.txtSoSeq.value)
    strVal = strVal & "&txtItemCode=" & Trim(frm1.txtItemCode.value)        
    strVal = strVal & "&txtTodayDate=" & Trim(frm1.txtTodayDate.value)        
    strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
    strVal = strVal & "&txtTrackingNO=" & Trim(frm1.txtTrackingNO.value)
    strVal = strVal & "&txtAPSHost=" & Trim(frm1.txtAPSHost.value)
    strVal = strVal & "&txtAPSPort=" & Trim(frm1.txtAPSPort.value)

	Call RunMyBizASP(MyBizASP, strVal)										

    DbCTPQuery = True                                                       

End Function


'================================================================================================================
Function DbCTPSave() 

    Err.Clear																	

	DbCTPSave = False															

	If frm1.rdoSelect_All.checked = True Then
		frm1.txtRadioFlg.value = frm1.rdoSelect_All.value
	ElseIf frm1.rdoSelect_Sub.checked = True Then
		frm1.txtRadioFlg.value = frm1.rdoSelect_Sub.value
	End If

	Call LayerShowHide(1)

    Dim strVal

	With frm1

		If UNICDbl(.txtCtpSeq.value) > 0 Then
			.txtMode.value		= CTPModify										<%'☜: 비지니스 처리 ASP 의 상태 %>
		Else
			.txtMode.value		= CTPAccept										<%'☜: 비지니스 처리 ASP 의 상태 %>
		End If

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
	End With

    DbCTPSave = True                                                           
    
End Function

'================================================================================================================
Function DbCTPSaveOk()															

	With frm1
		<% ' [저장전(OLD) 실제가능한 일자] 와 [저장후(NEW) 실제가능한 일자] 가 틀릴경우 %> 
		If CDate(.txtBeforeChangeDate.value) > CDate(.txtAfterChangeDate.value) Or CDate(.txtBeforeChangeDate.value) < CDate(.txtAfterChangeDate.value) Then

			If frm1.rdoSelect_All.checked = True Then
				.txtAccDate_All.value	= .txtAfterChangeDate.value
				'.txtAccDate_All.style.color = vbRed
				.txtAccDate_All.style.color = vbBlue
			ElseIf frm1.rdoSelect_Sub.checked = True Then
				.txtAccDate_Sub2.value	= .txtAfterChangeDate.value
				'.txtAccDate_Sub2.style.color = vbRed
				.txtAccDate_Sub2.style.color = vbBlue
			End If
		
		End If
	End With

	frm1.btnCTPSave.Disabled = True
	frm1.btnCTPCancel.Disabled = False
	frm1.btnSave.Disabled = False

End Function

'================================================================================================================
Function DbCTPCancel()

    Err.Clear																			

	If frm1.rdoSelect_All.checked = True Then
		frm1.txtRadioFlg.value = frm1.rdoSelect_All.value
	ElseIf frm1.rdoSelect_Sub.checked = True Then
		frm1.txtRadioFlg.value = frm1.rdoSelect_Sub.value
	End If
    
	Call LayerShowHide(1)
	    
    Dim strVal

	strVal = BIZ_PGM_ID & "?txtMode=" & CTPCancel										
	strVal = strVal & "&txtSoNo=" & Trim(frm1.txtSoNo.value)							
	strVal = strVal & "&txtSoSeq=" & Trim(frm1.txtSoSeq.value)
	strVal = strVal & "&txtItemCode=" & Trim(frm1.txtItemCode.value)
	strVal = strVal & "&txtPlantCd=" & Trim(frm1.txtPlantCd.value)
	strVal = strVal & "&txtAPSHost=" & Trim(frm1.txtAPSHost.value)
	strVal = strVal & "&txtAPSPort=" & Trim(frm1.txtAPSPort.value)
	strVal = strVal & "&txtRadioFlg=" & Trim(frm1.txtRadioFlg.value)

	Call RunMyBizASP(MyBizASP, strVal)													

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE CELLSPACING=0 CLASS="basicTB">
		<TR>
			<TD HEIGHT=5>&nbsp;<% ' 상위 여백 %></TD>
		</TR>
		<TR HEIGHT=23>
			<TD WIDTH=100%>
				<TABLE CELLSPACING=0 CELLPADDING=0 WIDTH=100% border=0>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD CLASS="CLSMTABP">
							<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
								<TR>
									<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>회답납기</font></td>
									<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
							    </TR>
							</TABLE>
						</TD>
						<TD WIDTH=500>&nbsp;</TD>
						<TD WIDTH=*>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD HEIGHT=40>
				<FIELDSET STYLE="margin-left:10px; margin-right:10px;">
					<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
						<TR>
							<TD CLASS=TD5>수주번호</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSoNo" SIZE=20 MAXLENGTH=18 TAG="14XXXU">&nbsp;-&nbsp;<INPUT TYPE=TEXT NAME="txtSoSeq" SIZE=5 TAG="14"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5>품목</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtItemCode" SIZE=20 MAXLENGTH=18 TAG="14XXXU"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5>&nbsp;</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtItemName" SIZE=40 TAG="14"></TD>
						</TR>
					</TABLE>
				</FIELDSET> 
			</TD>
		</TR>
		<TR>
			<TD HEIGHT=5>&nbsp;<% ' 상위 여백 %></TD>
		</TR>
		<TR>
			<TD WIDTH="100%" HEIGHT="*" valign="top">

				<!-- 신규입력 -->
				<DIV ID="divCreate" STYLE="FLOAT: left; HEIGHT: 160px; WIDTH: 450px; margin-left:5px; margin-right:5px;" SCROLL=no STYLE="Display=NONE;">
				<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
				  <TR>
					<TD width="100%" height="100%" align="center">
					  <TABLE border="1" width="100%" height="55" cellpadding="1" cellspacing="1" bordercolor="#C0C0C0">
						<TR>
						  <TD width="50%" height="25" bgcolor="#FFFF00" align="center"><font color="#C0C0C0"><b>납기요청일</b></font></TD>
						  <TD width="50%" height="25" bgcolor="#FFFF00" align="center"><font color="#C0C0C0"><b>납기요청수량</b></font></TD>
						</TR>
						<TR>
						  <TD width="50%" height="30" align="center"><input TYPE="TEXT" NAME="txtReqDate" SIZE="15" MAXLENGTH="10" TAG="11X1" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:CENTER"></TD>
						  <TD width="50%" height="30" align="center"><input TYPE="TEXT" NAME="txtReqQty" SIZE="15" MAXLENGTH="15" TAG="11X3" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:CENTER"></TD>
						</TR>
					  </TABLE>
					</TD>
				  </TR>
				  <TR>
					<TD width="100%" height="15" bgcolor="#0000FF" bordercolor="#C0C0C0" style="border-style: double; border-width: 5">
					</TD>
				  </TR>
				  <TR>
					<TD width="100%" height="100%" align="center">
					  <TABLE border="1" width="100%" height="85" cellpadding="1" cellspacing="1" bordercolor="#C0C0C0">
						<TR>
						  <TD width="33%" height="25" bgcolor="#FFFF00" align="center"><font color="#C0C0C0"><b>납기가능일</b></font></TD>
						  <TD width="33%" height="25" bgcolor="#FFFF00" align="center"><font color="#C0C0C0"><b>납기가능수량</b></font></TD>
						  <TD width="34%" height="25" bgcolor="#FFFF00" align="center"><font color="#C0C0C0"><b>선택</b></font></TD>
						</TR>
						<TR>
						  <TD width="33%" height="30" align="center"><input TYPE="TEXT" NAME="txtAccDate_All" SIZE="15" MAXLENGTH="10" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:CENTER"></TD>
						  <TD width="33%" height="30" align="center"><input TYPE="TEXT" NAME="txtAccQty_All" SIZE="15" MAXLENGTH="15" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:RIGHT"></TD>
						  <TD width="34%" height="30" align="center"><input TYPE="radio" CLASS="RADIO" name="rdoSelect" id="rdoSelect_All" value="A" tag="11" checked><Label for=rdoSelect_All>통합</Label></TD>
						</TR>
						<TR>
						  <TD width="33%" height="30" align="center"><input TYPE="TEXT" NAME="txtAccDate_Sub1" SIZE="15" MAXLENGTH="10" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:CENTER"></TD>
						  <TD width="33%" height="30" align="center"><input TYPE="TEXT" NAME="txtAccQty_Sub1" SIZE="15" MAXLENGTH="15" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:RIGHT"></TD>
						  <TD width="34%" height="60" rowspan="2" align="center"><input type="radio" CLASS="RADIO" name="rdoSelect" id="rdoSelect_Sub" value="S" tag="11"><Label for=rdoSelect_Sub>분할</Label></TD>
						</TR>
						<TR>
						  <TD width="33%" height="30" align="center"><input TYPE="TEXT" NAME="txtAccDate_Sub2" SIZE="15" MAXLENGTH="10" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:CENTER"></TD>
						  <TD width="33%" height="30" align="center"><input TYPE="TEXT" NAME="txtAccQty_Sub2" SIZE="15" MAXLENGTH="15" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:RIGHT"></TD>
						</TR>
					  </TABLE>
					</TD>
				  </TR>
				</TABLE>
				</DIV>

				<!-- 통합수정 -->
				<DIV ID="divCom" STYLE="FLOAT: left; HEIGHT: 80px; WIDTH: 450px; margin-left:5px; margin-right:5px;" SCROLL=no STYLE="Display=NONE;">
				<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
				  <TR>
					<TD width="100%" height="100%" align="center">
					  <TABLE border="1" width="100%" height="60" cellpadding="1" cellspacing="1" bordercolor="#C0C0C0">
						<TR>
						  <TD width="50%" height="25" bgcolor="#FFFF00" align="center"><font color="#C0C0C0"><b>납기가능일</b></font></TD>
						  <TD width="50%" height="25" bgcolor="#FFFF00" align="center"><font color="#C0C0C0"><b>납기가능수량</b></font></TD>
						</TR>
						<TR>
						  <TD width="50%" height="35" align="center"><input TYPE="TEXT" NAME="txtAccDate_All_Com" SIZE="15" MAXLENGTH="10" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:CENTER"></TD>
						  <TD width="50%" height="35" align="center"><input TYPE="TEXT" NAME="txtAccQty_All_Com" SIZE="15" MAXLENGTH="15" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:RIGHT"></TD>
						</TR>
					  </TABLE>
					</TD>
				  </TR>
				</TABLE>
				</DIV>

				<!-- 분할수정 -->
				<DIV ID="divDiv" STYLE="FLOAT: left; HEIGHT: 115px; WIDTH: 450px; margin-left:5px; margin-right:5px;" SCROLL=no STYLE="Display=NONE;">
				<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
				  <TR>
					<TD width="100%" height="100%" align="center">
					  <TABLE border="1" width="100%" height="95" cellpadding="1" cellspacing="1" bordercolor="#C0C0C0">
						<TR>
						  <TD width="50%" height="25" bgcolor="#FFFF00" align="center"><font color="#C0C0C0"><b>납기가능일</b></font></TD>
						  <TD width="50%" height="25" bgcolor="#FFFF00" align="center"><font color="#C0C0C0"><b>납기가능수량</b></font></TD>
						</TR>
						<TR>
						  <TD width="50%" height="35" align="center"><input TYPE="TEXT" NAME="txtAccDate_Sub1_Div" SIZE="15" MAXLENGTH="10" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:CENTER"></TD>
						  <TD width="50%" height="35" align="center"><input TYPE="TEXT" NAME="txtAccQty_Sub1_Div" SIZE="15" MAXLENGTH="15" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:RIGHT"></TD>
						</TR>
						<TR>
						  <TD width="50%" height="35" align="center"><input TYPE="TEXT" NAME="txtAccDate_Sub2_Div" SIZE="15" MAXLENGTH="10" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:CENTER"></TD>
						  <TD width="50%" height="35" align="center"><input TYPE="TEXT" NAME="txtAccQty_Sub2_Div" SIZE="15" MAXLENGTH="15" TAG="11" STYLE="BORDER-BOTTOM: 0px solid;BORDER-TOP: 0px solid;BORDER-RIGHT: 0px solid;BORDER-LEFT: 0px solid; TEXT-ALIGN:RIGHT"></TD>
						</TR>
					  </TABLE>
					</TD>
				  </TR>
				</TABLE>
				</DIV>

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
						<TD><BUTTON NAME="btnCTPSave" CLASS="CLSSBTN">CTP Accept</BUTTON>&nbsp;
							<BUTTON NAME="btnCTPCancel" CLASS="CLSSBTN">CTP Cancel</BUTTON>&nbsp;
							<BUTTON NAME="btnSave" CLASS="CLSSBTN">Save</BUTTON>
						</TD>
						<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="s3113pb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>
	<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
	<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
	<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
	<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
	<INPUT TYPE=HIDDEN NAME="txtRadioFlg" tag="24">
	<INPUT TYPE=HIDDEN NAME="txtCtpSeq" tag="24">
	<INPUT TYPE=HIDDEN NAME="txtPlantCd" tag="24">
	<INPUT TYPE=HIDDEN NAME="txtTodayDate" tag="24">
	<INPUT TYPE=HIDDEN NAME="txtBeforeChangeDate" tag="24">	
	<INPUT TYPE=HIDDEN NAME="txtAfterChangeDate" tag="24">	
	<INPUT TYPE=HIDDEN NAME="txtTrackingNO" tag="24">	
	<INPUT TYPE=HIDDEN NAME="txtAPSHost" tag="24">	
	<INPUT TYPE=HIDDEN NAME="txtAPSPort" tag="24">	
	<INPUT TYPE=HIDDEN NAME="txtCTPTimes" tag="24">	
	<INPUT TYPE=HIDDEN NAME="txtCtpCDFlag" tag="24">
	
	<INPUT TYPE=HIDDEN NAME="txtExitFlag" tag="24">	
	
	</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>