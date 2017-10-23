<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 																			*
'*  3. Program ID           : S3112RA3																	*
'*  4. Program Name         : ATP수행																	*
'*  5. Program Desc         :																			*
'*  6. Comproxy List        : 															*
'*  7. Modified date(First) : 2001/01/07																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Kim Hyungsuk																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 2001/01/07 : 화면 design													*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		<% '☆: 해당 위치에 따라 달라짐, 상대 경로 %>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
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

Const BIZ_PGM_ID = "s3112rb3.asp"						

Dim C_SchdDlvyDt								
Dim C_PromiseGIDt			
Dim C_ComfirmSOQty		
Dim C_ComfirmBonusQty		
Dim C_ComfirmBaseQty		
Dim C_ComfirmBaseBonusQty
Dim C_ProductFlg			
Dim C_ATPFlg	
			
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim arrReturn					
Dim arrParam	
Dim gblnWinEvent			


'========================================================================================================
Sub initSpreadPosVariables()
	C_SchdDlvyDt			= 1								
	C_PromiseGIDt			= 2
	C_ComfirmSOQty			= 3
	C_ComfirmBonusQty		= 4
	C_ComfirmBaseQty		= 5
	C_ComfirmBaseBonusQty	= 6
	C_ProductFlg			= 7
	C_ATPFlg				= 8
End Sub

'========================================================================================================
Function InitVariables()
	lgIntFlgMode = PopupParent.OPMD_CMODE								
	lgIntGrpCount = 0										
	lgStrPrevKey = ""	
	gblnWinEvent = False
	Self.Returnvalue = ""
End Function
	
'========================================================================================================
Sub SetDefaultVal()
	arrParam = arrParent(1)
	frm1.txtSONo.value = arrParam(0)
	frm1.txtSOSeq.value = arrParam(1)
	frm1.btnCompleteDelevery.disabled = True
	frm1.btnPartailDelevery.disabled = True  
End Sub

'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
		
	ggoSpread.Spreadinit "V20021214",,PopupParent.gAllowDragDropSpread
	frm1.vspdData.ReDraw = False
	frm1.vspdData.MaxCols = C_ATPFlg + 1
	frm1.vspdData.MaxRows = 0
		
	Call GetSpreadColumnPos("A")	 		

	ggoSpread.SSSetDate		C_SchdDlvyDt, "가능납기일",12,2,PopupParent.gDateFormat
	ggoSpread.SSSetDate		C_PromiseGIDt, "출고예정일",12,2,PopupParent.gDateFormat
    ggoSpread.SSSetFloat	C_ComfirmSOQty,"확정수주량" ,15,PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
    ggoSpread.SSSetFloat	C_ComfirmBonusQty,"확정덤수량" ,15,PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
    ggoSpread.SSSetFloat	C_ComfirmBaseQty,"확정수주량(재고단위)" ,20,PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
    ggoSpread.SSSetFloat	C_ComfirmBaseBonusQty,"확정덤수량(재고단위)" ,20,PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
    ggoSpread.SSSetEdit		C_ProductFlg, "추가생산", 10, 0
	ggoSpread.SSSetEdit		C_ATPFlg, "ATP구분", 10, 0
		
	Call ggoSpread.SSSetColHidden(C_ATPFlg, C_ATPFlg, True)
	Call ggoSpread.SSSetColHidden(frm1.vspdData.MaxCols, frm1.vspdData.MaxCols, True)				'☜: 공통콘트롤 사용 Hidden Column
		 		
	SetSpreadLock "", 0, -1, ""

	frm1.vspdData.ReDraw = True
End Sub

'========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_SchdDlvyDt			= iCurColumnPos(1)   
			C_PromiseGIDt			= iCurColumnPos(2)   
			C_ComfirmSOQty			= iCurColumnPos(3)   
			C_ComfirmBonusQty		= iCurColumnPos(4)
			C_ComfirmBaseQty		= iCurColumnPos(5)   
			C_ComfirmBaseBonusQty	= iCurColumnPos(6) 
			C_ProductFlg			= iCurColumnPos(7)  
			C_ATPFlg				= iCurColumnPos(8)  		
    End Select    
End Sub


'========================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
	ggoSpread.Source = frm1.vspdData
			
	frm1.vspdData.ReDraw = False
			
	ggoSpread.SpreadLock C_SchdDlvyDt, lRow, -1
	ggoSpread.SpreadLock C_PromiseGIDt, lRow, -1
	ggoSpread.SpreadLock C_ComfirmSOQty, lRow, -1
	ggoSpread.SpreadLock C_ComfirmBonusQty, lRow, -1
	ggoSpread.SpreadLock C_ComfirmBaseQty, lRow, -1
	ggoSpread.SpreadLock C_ComfirmBaseBonusQty, lRow, -1
	ggoSpread.SpreadLock C_ProductFlg, lRow, -1
	ggoSpread.SpreadLock C_ATPFlg, lRow, -1
			
	frm1.vspdData.ReDraw = True
End Sub
	
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
Function ATPOk()
	Dim IntRetCD

	Err.Clear												

	Call DisplayMsgBox("183114", "X", "X", "X")  '☜ 바뀐부분 
	Call CancelClick()
End Function

'========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029											<%  %>
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")						<%  %>
	Call InitVariables
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call FncQuery()
End Sub

'========================================================================================================
Sub btnCompleteDelevery_OnClick()
	Err.Clear															

	frm1.txtInsrtUserId.value = PopupParent.gUsrID

	If LayerShowHide(1) = False Then
		Exit Sub
	End If

	Dim strVal

	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & "COMPLETE"		
		strVal = strVal & "&txtSONo=" & Trim(frm1.txtSONo.value)			
		strVal = strVal & "&txtSOSeq=" & Trim(frm1.txtSOSeq.value)			
		strVal = strVal & "&txtSchdDlvyDt=" & Trim(frm1.txtAvalSchdDlvyDt.text)			
		strVal = strVal & "&txtPromiseGIDt=" & Trim(frm1.txtAvalGIDt.text)
		strVal = strVal & "&C_ComfirmSOQty=" & Trim(frm1.txtSOQty.text)
		strVal = strVal & "&C_ComfirmBonusQty=" & Trim(frm1.txtBonusQty.text)
		strVal = strVal & "&C_ComfirmBaseQty=" & Trim(frm1.txtHBaseQty.value)
		strVal = strVal & "&C_ComfirmBaseBonusQty=" & Trim(frm1.txtHBonusBaseQty.value)
		strVal = strVal & "&C_ATPFlg=" & Trim(frm1.txtHATPFlag.value)
		strVal = strVal & "&txtInsrtUserId=" & Trim(frm1.txtInsrtUserId.value)
	Else 
		Exit Sub
	End If

	Call RunMyBizASP(MyBizASP, strVal)									
End Sub

'========================================================================================================
Sub btnPartailDelevery_OnClick()
	Dim lRow
	Dim lGrpCnt
	Dim strVal
	Dim intInsrtCnt
	    
	If LayerShowHide(1) = False Then
		Exit Sub
	End If

    frm1.txtMode.value = "PARTIAL"
	frm1.txtUpdtUserId.value = PopupParent.gUsrID
	frm1.txtInsrtUserId.value = PopupParent.gUsrID

	lGrpCnt = 1

	strVal = ""

	With frm1
		For lRow = 1 To .vspdData.MaxRows
			.vspdData.Row = lRow
			.vspdData.Col = 0

			strVal = strVal & lRow & PopupParent.gColSep						

			.vspdData.Col = C_SchdDlvyDt							
			strVal = strVal & Trim(.vspdData.Text) & PopupParent.gColSep

			.vspdData.Col = C_PromiseGIDt							
			strVal = strVal & Trim(.vspdData.Text) & PopupParent.gColSep

			.vspdData.Col = C_ComfirmSOQty							
			strVal = strVal & Trim(.vspdData.Text) & PopupParent.gColSep

			.vspdData.Col = C_ComfirmBonusQty					
			strVal = strVal & Trim(.vspdData.Text) & PopupParent.gColSep
								
			.vspdData.Col = C_ComfirmBaseQty						
			strVal = strVal & Trim(.vspdData.Text) & PopupParent.gColSep

			.vspdData.Col = C_ComfirmBaseBonusQty					
			strVal = strVal & Trim(.vspdData.Text) & PopupParent.gColSep

			.vspdData.Col = C_ATPFlg								
			strVal = strVal & Trim(.vspdData.Text) & PopupParent.gRowSep
								
			lGrpCnt = lGrpCnt + 1
		Next

		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strVal
	End With
				
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										

End Sub

'========================================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
	With frm1.vspdData
		If Row >= NewRow Then
			Exit Sub
		End If

		If NewRow = .MaxRows Then
			If lgStrPrevKey <> "" Then							
				DbQuery
			End If
		End If
	End With
End Sub
	
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgStrPrevKey <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub
	
'========================================================================================================
Sub FncSplitColumn()        
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)    
End Sub

'========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	
	Call SetQuerySpreadColor(1)    

End Sub

'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
   On Error Resume Next
   If KeyAscii = 27 Then
	  Call CancelClick()
   End If
End Function

'========================================================================================================
Function FncQuery()
	Dim IntRetCD

	FncQuery = False													

	Err.Clear	
		
	Call ggoOper.ClearField(Document, "2")								
	Call InitVariables	
	
	If Not chkField(Document, "1") Then							
		Exit Function
	End If
	
	Call DbQuery()														

	FncQuery = True														
End Function

'========================================================================================================
Function DbQuery()
	Err.Clear															

	DbQuery = False														

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
		strVal = strVal & "&txtSONo=" & Trim(frm1.txtHSONo.value)			
		strVal = strVal & "&txtSOSeq=" & Trim(frm1.txtHSOSeq.value)
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
		strVal = strVal & "&txtSONo=" & Trim(frm1.txtSONo.value)			
		strVal = strVal & "&txtSOSeq=" & Trim(frm1.txtSOSeq.value)
	End If

	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True														
End Function

'========================================================================================================
Function DbQueryOk()													
	Dim lRow
	lgIntFlgMode = PopupParent.OPMD_UMODE											
	With frm1
			
		If .txtAvalGIDt.text = "" Then
			.btnCompleteDelevery.disabled = True
		Else
			.btnCompleteDelevery.disabled = False
		End If
			
		If .vspdData.MaxRows = 0 then
			.btnPartailDelevery.disabled = True
		Else
			.btnPartailDelevery.disabled = False
		End If
		
		'헤더빨강	
		If UniConvDateToYYYYMMDD(.txtAvalSchdDlvyDt.Text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtSchdDlvyDt.text,PopupParent.gDateFormat,"") Then 
			.txtAvalSchdDlvyDt.forecolor = &HFF&
		End If
		
		If UniConvDateToYYYYMMDD(.txtAvalGIDt.Text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtSchdDlvyDt.text,PopupParent.gDateFormat,"") Then 
			.txtAvalGIDt.forecolor = &HFF&
		End If	
		
		'그리드 빨강			
		For lRow = 1 To .vspdData.MaxRows          
				
			.vspdData.Row = lRow
				
			.vspdData.Col = C_SchdDlvyDt
	
			If UniConvDateToYYYYMMDD(.vspdData.Text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtSchdDlvyDt.text,PopupParent.gDateFormat,"") Then 
				Call sprRedComColor(C_SchdDlvyDt,lRow,lRow)
			End If
				
			.vspdData.Col = C_PromiseGIDt
	
			If UniConvDateToYYYYMMDD(.vspdData.Text,PopupParent.gDateFormat,"") > UniConvDateToYYYYMMDD(.txtSchdDlvyDt.text,PopupParent.gDateFormat,"") Then 
				Call sprRedComColor(C_PromiseGIDt,lRow,lRow)
			End If
		Next     		
		
	End With		
	
End Function


'========================================================================================================
Sub sprRedComColor(ByVal Col, ByVal Row, ByVal Row2)
    With frm1
		.vspdData.Col = Col
		.vspdData.Col2 = Col
		.vspdData.Row = Row
		.vspdData.Row2 = Row2
		.vspdData.ForeColor = vbRed
    End With    
End Sub


'========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("0000111111")
	
    gMouseClickStatus = "SPC"    
    
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	
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

'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
    Dim intIndex

	If Col = C_VatIncFlagNm Then
 
		With frm1.vspdData
		  .Row = Row
		  .Col = Col
		  intIndex = .Value
		  
		  .Col = C_VatIncFlag
		  .Value = intIndex+1
		End With
		
		Call vspdData_Change(C_VatIncFlag , Row)
		
	End If
End Sub

'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub  
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>수주번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSONo" ALT="수주번호" SIZE=20 TYPE="Text" MAXLENGTH="18" tag="14XXXU"></TD>
						<TD CLASS=TD5 NOWRAP>수주순번</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSOSeq" ALT="수주순번" SIZE=5 TYPE="Text" MAXLENGTH="2" tag="14"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>품목</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItem" ALT="품목" TYPE="Text" MAXLENGTH=18 SIZE=20 TAG="24XXXU"></TD>
						<TD CLASS=TD5 NOWRAP>품목명</TD>
						<TD CLASS=TD6 NOWRAP><INPUT NAME="txtItemNm" ALT="품목명" TYPE="Text" MAXLENGTH=50 SIZE=20 TAG="24"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>공장</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlant" SIZE=10 MAXLENGTH=4 TAG="24XXXU" ALT="공장">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 TAG="24"></TD>
						<TD CLASS=TD5 NOWRAP>Tracking No</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=20 MAXLENGTH=25 TAG="24XXXU" ALT="Tracking No"></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>수주단위</TD>
						<TD CLASS=TD6><INPUT NAME="txtSOUnit" ALT="수주단위" TYPE=TEXT MAXLENGTH=5 SIZE=10 TAG="24XXXU" TABINDEX=-1></TD>
						<TD CLASS=TD5 NOWRAP>요청납기일</TD>						
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3112ra3_fpDateTime1_txtSchdDlvyDt.js'></script></TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>수주수량</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3112ra3_fpDoubleSingle1_txtSOQty.js'></script>
						<TD CLASS=TD5 NOWRAP>덤수량</TD>
						<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3112ra3_fpDoubleSingle1_txtBonusQty.js'></script>														
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_60%>>
				<TR>	
					<TD CLASS=TD5 NOWRAP>일괄가능납기일</TD>						
					<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/s3112ra3_fpDateTime1_txtAvalSchdDlvyDt.js'></script>
										<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="24X" VALUE="F" NAME="chkPPFlg" ID="chkPPFlg">
										<LABEL FOR="chkPPFlg">추가생산</LABEL>
										</TD>
					<TD CLASS=TD5 NOWRAP>일괄출고예정일</TD>						
					<TD CLASS=TD6 NOWRAP>
						<script language =javascript src='./js/s3112ra3_fpDateTime1_txtAvalGIDt.js'></script>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<BUTTON NAME="btnCompleteDelevery" CLASS="CLSSBTN">일괄납품</BUTTON></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>분할납기가능일</TD>						
					<TD CLASS=TD6 NOWRAP></OBJECT></TD>
					<TD CLASS=TDT NOWRAP></TD>						
					<TD CLASS=TD6 NOWRAP>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<BUTTON NAME="btnPartailDelevery" CLASS="CLSSBTN">분할납품</BUTTON>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<script language =javascript src='./js/s3112ra3_vaSpread_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP>     <BUTTON NAME="btnATP" CLASS="CLSMBTN" ONCLICK="vbscript:FncQuery()">ATP 재수행</BUTTON></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME></TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtMaxSeq" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSONo" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHSOSeq" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHBaseQty" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHBonusBaseQty" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHAtpFlag" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>