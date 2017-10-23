<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : scm																		*
'*  2. Function Name        :																			*
'*  3. Program ID           : U2124PA1.asp																*
'*  4. Program Name         : 명세서발행번호팝업																*
'*  5. Program Desc         : 명세서발행번호팝업																*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2004/08/11																*
'*  8. Modified date(Last)  : 2004/08/11																*
'*  9. Modifier (First)     : nhg												*
'* 10. Modifier (Last)      : nhg																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit		

'================================================================================================================================
Const BIZ_PGM_ID 		= "u2124pb1.asp"

'================================================================================================================================
Dim C_DLVY_No
Dim C_DLVY_TIME
Dim C_TRANS_CO
Dim C_VEHICLE
Dim C_DRIVER
Dim C_TEL_NO1
Dim C_TEL_NO2
Dim C_DLVY_PLACE
Dim C_REMARK

'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================

'================================================================================================================================
Const C_MaxKey          = 28                                           '☆: key count of SpreadSheet
Dim IsOpenPop
Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam
Dim EndDate, StartDate
'================================================================================================================================    
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName
arrParam= arrParent(1)

EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
'================================================================================================================================
Function InitVariables()
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                        'Indicates that current mode is Create mode
    lgSortKey        = 1
	lgKeyStream		 = ""
	frm1.vspdData.MaxRows = 0	
	
	IsOpenPop = False
	ReDim arrReturn(0)
	Self.Returnvalue = arrReturn
End Function
'================================================================================================================================
Sub SetDefaultVal()
	
	Dim iCodeArr
		
	Err.Clear
	
	With frm1
		
		.hdnBpCd.value 	= arrParam(0)
		
		'.txtBpCd.value		=  PopupParent.gPlant
		'.txtBpNm.value		=  PopupParent.gPlantNm
	End With
	
	CALL SetBPCD()
	
End Sub

'================================================================================================================================
Sub SetBPCD()
	Dim lgF0 
 
	'If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = " & FilterVar(popupparent.gUsrId, "", "S"), lgF0) = False Then
	If 	CommonQueryRs2by2(" BP_NM ", " B_BIZ_PARTNER ", " BP_CD = '" & Trim(frm1.hdnBpCd.value)  & "'" , lgF0) = False Then
		Exit Sub
	End If

	lgF0 = Split(lgF0, Chr(11))
	'frm1.txtBpCd.value = popupparent.gUsrId
	frm1.txtBpCd.value = Trim(frm1.hdnBpCd.value)
	frm1.txtBpNm.value = lgF0(1)

End Sub


'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("I", "*", "NOCOOKIE", "RA") %>
End Sub
'================================================================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	'------------------------------------------
	' Grid 1 - Operation Spread Setting
	'------------------------------------------
	With frm1.vspdData 
			
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20050420", ,PopupParent.gAllowDragDropSpread
		frm1.vspdData.OperationMode = 3
		
		.ReDraw = false
					
		.MaxCols = C_REMARK + 1    
		.MaxRows = 0    
			
		Call GetSpreadColumnPos()

		ggoSpread.SSSetEdit 		C_DLVY_No,		"명세서발행번호",14
		ggoSpread.SSSetEdit 		C_DLVY_TIME,	"납입시간"		,14
		ggoSpread.SSSetEdit 		C_TRANS_CO,		"운송회사"		,14
		ggoSpread.SSSetEdit 		C_VEHICLE,		"차량번호"		,14
		ggoSpread.SSSetEdit 		C_DRIVER,		"운전자"		,10
		ggoSpread.SSSetEdit 		C_TEL_NO1,		"전화번호1"		,10
		ggoSpread.SSSetEdit 		C_TEL_NO2,		"전화번호2"		,10
		ggoSpread.SSSetEdit 		C_DLVY_PLACE,	"납품장소"		,10
		ggoSpread.SSSetEdit 		C_REMARK,		"비고"			,20
			
		Call ggoSpread.SSSetColHidden( .MaxCols, .MaxCols, True)
			
						
		Call SetSpreadLock()
						
		.ReDraw = true    
    
	End With
	   
End Sub
'================================================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'================================================================================================================================
Sub InitSpreadPosVariables()

	C_DLVY_No		= 1
	C_DLVY_TIME		= 2
	C_TRANS_CO		= 3
	C_VEHICLE		= 4
	C_DRIVER		= 5
	C_TEL_NO1		= 6
	C_TEL_NO2		= 7
	C_DLVY_PLACE	= 8
	C_REMARK		= 9
	
End Sub
'================================================================================================================================
Sub GetSpreadColumnPos()
      
    Dim iCurColumnPos
    
 	ggoSpread.Source = frm1.vspdData
		
	Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_DLVY_No		= iCurColumnPos(1)
	C_DLVY_TIME		= iCurColumnPos(2)
	C_TRANS_CO		= iCurColumnPos(3)
	C_VEHICLE		= iCurColumnPos(4)
	C_DRIVER		= iCurColumnPos(5)
	C_TEL_NO1		= iCurColumnPos(6)
	C_TEL_NO2		= iCurColumnPos(7)
	C_DLVY_PLACE	= iCurColumnPos(8)
	C_REMARK		= iCurColumnPos(9)
	
End Sub    
'================================================================================================================================
Function OKClick()
	Dim intRowCnt
	Dim intColCnt
	Dim intSelCnt

	If frm1.vspdData.MaxRows > 0 Then
		intSelCnt = 0
		Redim arrReturn(0)
		
		frm1.vspdData.Row = frm1.vspdData.ActiveRow

		If frm1.vspdData.SelModeSelected = True Then
			frm1.vspdData.Col = C_DLVY_NO
			arrReturn(0) = frm1.vspdData.Text
		End If

		Self.Returnvalue = arrReturn
	End If		
	Self.Close()
End Function	
'================================================================================================================================
Function CancelClick()
	Redim arrReturn(1)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'================================================================================================================================
Function OpenSortPopup()

	On Error Resume Next
	
End Function

'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "납품처팝업"
	arrParam(1) = "(			SELECT	DISTINCT B.PLANT_CD FROM M_SCM_PLAN_PUR_RCPT A, M_PUR_ORD_DTL B, M_PUR_ORD_HDR C "
	arrParam(1) = arrParam(1) & "WHERE	A.PO_NO = B.PO_NO AND A.PO_SEQ_NO = B.PO_SEQ_NO AND A.SPLIT_SEQ_NO = 0 "
	arrParam(1) = arrParam(1) & "AND	A.PO_NO = C.PO_NO AND C.BP_CD = '" & frm1.hdnBpCd.value & "') A, B_PLANT B"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = "A.PLANT_CD = B.PLANT_CD"			
	arrParam(5) = "납품처"			
	
    arrField(0) = "A.PLANT_CD"	
    arrField(1) = "B.PLANT_NM"	
    
    arrHeader(0) = "납품처"		
    arrHeader(1) = "납품처명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else	
		frm1.txtPlantCd.value = arrRet(0)
		frm1.txtPlantNm.value = arrRet(1)
		frm1.txtPlantCd.focus	
		Set gActiveElement = document.activeElement
	End If	
	
End Function

'================================================================================================================================
Sub Form_Load()
	
	Call LoadInfTB19029															'⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)	                                           
	Call ggoOper.LockField(Document, "N")										'⊙: Lock  Suitable  Field 
	Call InitVariables														    '⊙: Initializes local global variables
	
	Call SetDefaultVal	
		
	Call InitSpreadSheet()
		
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")

	Call FncQuery()
End Sub
'================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    
End Sub
'================================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or Frm1.vspdData.MaxRows = 0 Then 
	     Exit Sub
	End If

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub
'================================================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub

'================================================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        
	
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData
        
	Call InitVariables												
	
	If CheckRunningBizProcess = True Then Exit Function
    If DbQuery = False Then Exit Function

    FncQuery = True									
        
End Function
'================================================================================================================================
Function DbQuery()
	
	Dim strVal
	
	Err.Clear															'☜: Protect system from crashing

	DbQuery = False														'⊙: Processing is NG

    If LayerShowHide(1) = False Then Exit Function
    
    Call MakeKeyStream()
    
	strVal = BIZ_PGM_ID & "?txtMode="	& PopupParent.UID_M0001
	strVal = strVal & "&txtKeyStream="  & lgKeyStream
	strVal = strVal & "&lgStrPrevKey="  & lgPageNo
	
	Call RunMyBizASP(MyBizASP, strVal)									'☜: 비지니스 ASP 를 가동 

	DbQuery = True														'⊙: Processing is NG
End Function
'================================================================================================================================
Function DbQueryOk()														<%'☆: 조회 성공후 실행로직 %>

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtPlantCd.focus
	End If

End Function
'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream()

	With frm1
		lgKeyStream = lgKeyStream & UCase(Trim(.txtBPCd.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.txtPlantCd.value))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.txtFrPoDt.text))  & PopupParent.gColSep
		lgKeyStream = lgKeyStream & UCase(Trim(.txtToPoDt.text))  & PopupParent.gColSep
		
	End With
			 
End Sub    

'================================================================================================================================
Sub txtFrPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_Keypress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 27 Then
		Call CancelClick()
	Elseif KeyAscii = 13 Then
		Call FncQuery()
	End if
End Sub
'================================================================================================================================
Sub txtFrPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtFrPoDt.Focus
	End if
End Sub
'================================================================================================================================
Sub txtToPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToPoDt.Action = 7
		Call SetFocusToDocument("P")	
        frm1.txtToPoDt.Focus
	End if
End Sub
'================================================================================================================================

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS="TD5" NOWRAP>업체</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtBpCd" SIZE=10 MAXLENGTH=10 ALT="업체" tag="14xxxU">
											   <INPUT TYPE=TEXT AlT="업체명" ID="txtBpNm" tag="14X"></TD>
						<TD CLASS="TD5" NOWRAP>명세서발행번호</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=18 MAXLENGTH=18 ALT="명세서발행번호" tag="1XxXXU"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>납품예정일</td>
						<TD CLASS="TD6" NOWRAP>
							<table cellspacing=0 cellpadding=0>
								<tr>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="납품예정일" NAME="txtFrPoDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
									<td NOWRAP>~</td>
									<td NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT ALT="납품예정일" NAME="txtToPoDt" CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD tag="11N1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</td>
								<tr>
							</table>
						</TD>
						<TD CLASS="TD5" NOWRAP>
						<TD CLASS="TD6" NOWRAP>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* valign=top>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% ID=vspdData> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
					</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
                    <TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnFrPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToPoDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnBpCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>                                                                                                                                                                                                                                          