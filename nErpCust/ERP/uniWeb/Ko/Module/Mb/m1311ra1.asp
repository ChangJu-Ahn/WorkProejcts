<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m1311ra1.asp																*
'*  4. Program Name         : Bom 정보 																	*
'*  5. Program Desc         : Bom 정보 Ref																*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2000/03/21																*
'*  8. Modified date(Last)  : 2003/06/12																*
'*  9. Modifier (First)     : Park Jin Uk																*
'* 10. Modifier (Last)      : Kim Jin Ha																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/21 : 화면 design												*
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE="VBScript">

Option Explicit					

<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================

Const BIZ_PGM_QRY_ID = "m1311rb1.asp"			<% '☆: 비지니스 로직 ASP명 %>
Const C_MaxKey          = 10        

Dim C_ChdItemCd
Dim C_ChdItemNm
Dim C_ParItemQty
Dim C_ParItemUnit
Dim C_ChdItemQty
Dim C_ChdItemUnit
Dim C_Paytype
Dim C_Loss
Dim C_ValidFrDt
Dim C_ValidToDt

Dim IsOpenPop					
Dim arrReturn					
Dim arrParent
Dim arrParam
'================================================================================================================================
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName



	
'================================================================================================================================
Function InitVariables()
		
	Dim temp 
	Dim strYear,strMonth,strDay
		
	lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>
		
	frm1.txtPlantCd.value 	= arrParam(1)
	frm1.txtItemCd.value 	= arrParam(2)
	'=======================================
	'200704 KSJ 추가(BOM적용유효일추가)
	frm1.txtFrDt.Text    	= arrParam(3)
	frm1.txtToDt.Text   	= arrParam(4)
	'=======================================
	frm1.txtBomNo.Value		= arrParam(5)
	
	if arrParam(0) = PopupParent.OPMD_UMODE then
		Call ggoOper.LockField(Document, "Q")
		Call FncQuery()
	else
		Call ggoOper.LockField(Document,"N")
	End if
		
	Redim arrReturn(0, 0)

		
	Self.Returnvalue = arrReturn
	if frm1.txtPlantCd.Value <> "" And frm1.txtItemCd.Value <> "" then
		FncQuery()
	End if
		
End Function
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "RA") %>
End Sub
'================================================================================================================================
Sub InitSpreadPosVariables()
	C_ChdItemCd 	= 1
	C_ChdItemNm 	= 2
	C_ParItemQty	= 3 
	C_ParItemUnit	= 4
	C_ChdItemQty  	= 5
	C_ChdItemUnit	= 6
	C_Paytype		= 7
	C_Loss		 	= 8
	C_ValidFrDt		= 9
	C_ValidToDt		= 10
End Sub
'================================================================================================================================
Sub InitSpreadSheet()
	
	Call InitSpreadPosVariables()
	
	With frm1.vspdData
		
		ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20030612",, PopupParent.gAllowDragDropSpread
        .ReDraw = false
	
		.OperationMode 	= 5												<%'multiSelection Mode%>
		.MaxCols = C_ValidToDt+1												'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = C_ValidToDt+1:    .ColHidden = True
		.MaxRows = 0
		
		Call GetSpreadColumnPos("A")
		
		ggoSpread.Spreadinit
		ggoSpread.SSSetEdit 	C_ChdItemCd,	"자품목", 15,,,18,2
		ggoSpread.SSSetEdit 	C_ChdItemNm,	"자품목명", 25
		SetSpreadFloat		 	C_ParItemQty,	"모품목수량", 15,1,6
		ggoSpread.SSSetEdit 	C_ParItemUnit,	"모품목단위", 15,,,3,2
		SetSpreadFloat		 	C_ChdItemQty,	"자품목소요수량", 15,1,6
		ggoSpread.SSSetEdit 	C_ChdItemUnit,	"자품목단위",15,,,3,2
		ggoSpread.SSSetEdit 	C_Paytype,		"지급구분",10
		SetSpreadFloat		 	C_Loss,			"Loss율", 15,1,5
		ggoSpread.SSSetDate 	C_ValidFrDt,	"시작유효일", 15, 2, popupParent.gDateFormat
		ggoSpread.SSSetDate 	C_ValidFrDt,	"시작유효일", 15, 2, popupParent.gServerDateFormat
		ggoSpread.SSSetDate 	C_ValidToDt,	"종료유효일", 15, 2, popupParent.gDateFormat
		
		Call SetSpreadLock 
    
		.ReDraw = true
	
    End With
End Sub
'================================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
     
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_ChdItemCd  	= iCurColumnPos(1)
			C_ChdItemNm 	= iCurColumnPos(2)
			C_ParItemQty 	= iCurColumnPos(3)
			C_ParItemUnit	= iCurColumnPos(4) 
			C_ChdItemQty 	= iCurColumnPos(5)
			C_ChdItemUnit 	= iCurColumnPos(6)
			C_Paytype  		= iCurColumnPos(7)
			C_Loss			= iCurColumnPos(8)
			C_ValidFrDt		= iCurColumnPos(9)
			C_ValidToDt		= iCurColumnPos(10)
    End Select    
End Sub
'================================================================================================================================
Sub SetSpreadLock()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub
'================================================================================================================================
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow
	with frm1
	If .vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(.vspdData.SelModeSelCount+1, .vspdData.MaxCols - 1)

		For intRowCnt = 0 To .vspdData.MaxRows - 1

			.vspdData.Row = intRowCnt + 1

			If .vspdData.SelModeSelected Then
				For intColCnt = 0 To .vspdData.MaxCols - 1
					.vspdData.Col = intColCnt + 1
					arrReturn(intInsRow, intColCnt) = .vspdData.Text
				Next

				intInsRow = intInsRow + 1

			End IF
		Next
			
		arrReturn(.vspdData.SelModeSelCount,1) = Trim(frm1.txtPlantCd.Value)
		arrReturn(.vspdData.SelModeSelCount,2) = Trim(frm1.txtItemCd.Value)
		arrReturn(.vspdData.SelModeSelCount,3) = ""
		arrReturn(.vspdData.SelModeSelCount+1,1) = Trim(frm1.txtPlantNm.Value)
		arrReturn(.vspdData.SelModeSelCount+1,2) = Trim(frm1.txtItemNm.Value)
		arrReturn(.vspdData.SelModeSelCount+1,3) = ""
		arrReturn(.vspdData.SelModeSelCount+1,4) = Trim(frm1.txtBomNo.Value)
			
	End if			
	end with
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'================================================================================================================================
Function CancelClick()
	Self.Close()
End Function
'================================================================================================================================
Function OpenItem()
	Dim arrRet
	Dim arrParam(5), arrField(2)
	Dim iCalledAspName
	Dim IntRetCD
	
	If IsOpenPop = True Then Exit Function
	if UCase(frm1.txtItemCd.ClassName) = UCase(PopupParent.UCN_PROTECTED) then Exit Function
	
	if Trim(frm1.txtPlantCd.Value) = "" then		
		Call DisplayMsgBox("17A002","X" , "공장","X")
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if 
	
	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)	' Plant Code
	arrParam(1) = Trim(frm1.txtItemCd.value)		' Item Code
	arrParam(2) = "12!MO"	' “12!MO"로 변경 -- 품목계정 구분자 조달구분 
	arrParam(3) = "20!M"
	arrParam(4) = ""		'-- 날짜 
	arrParam(5) = ""		'-- 자유(b_item_by_plant a, b_item b: and 부터 시작)
	
	arrField(0) = 1 ' -- 품목코드 
	arrField(1) = 2 ' -- 품목명 
	arrField(2) = 3 ' -- Spec		
	
	
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	
	arrRet = window.showModalDialog(iCalledAspName, Array(window.PopupParent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")				
    
	'arrRet = window.showModalDialog("../../comasp/B1B11PA3.asp", Array(arrParam, arrField, arrHeader), _
	'	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtItemCd.Value  = arrRet(0)		
		frm1.txtItemNm.Value  = arrRet(1)	
		frm1.txtItemCd.focus
		Set gActiveElement = document.activeElement
	End If	
End Function
'================================================================================================================================
Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
    
    Call AppendNumberPlace("6","5","4")
    	
	Call InitSpreadSheet()
	Call InitVariables
	exit sub
	
	
End Sub
'================================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
	   Exit Sub
	End If
	   	    
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey
			lgSortKey = 1
		End If
		Exit Sub
	End If

	Call SetPopupMenuItemInf("0001111111")
End Sub
'================================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'================================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'================================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'================================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub
'================================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'================================================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
		Exit Function
	End If
    
	With frm1.vspdData 
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
End Function
'================================================================================================================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			If DBQuery = False Then
				Exit Sub
			End If
		End If
    End if
    
End Sub
'================================================================================================================================
 Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

    If Not chkField(Document, "1") Then									
       Exit Function
    End If
    
    if Trim(frm1.txtFrDt.Text) = "" or Trim(frm1.txtToDt.Text) = "" then
		Call DisplayMsgBox("17A002","X" , "적용유효일","X")
		Set gActiveElement = document.activeElement	
		Exit Function 
	end if

    'Call ggoOper.ClearField(Document, "2")								

    ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData       	
    
    Call DbQuery														
       
    FncQuery = True														
        
End Function	
'================================================================================================================================
Function DbQuery() 

    Dim strVal
    Err.Clear                                                               
    
    DbQuery = False                                                         
    
    If LayerShowHide(1) = False then
       Exit Function 
    End if
    
    With frm1
    
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
			strVal = strVal & "&txtItemCd=" & .hdnItemCd.value
			strVal = strVal & "&txtFrDt=" & .hdnFrDt.value
			strVal = strVal & "&txtToDt=" & .hdnToDt.value
		Else
			strVal = BIZ_PGM_QRY_ID & "?txtMode=" & PopupParent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
			strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
			strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.Text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.Text)
		End if 
    End with
	Call RunMyBizASP(MyBizASP, strVal)										
		
    DbQuery = True                                                          

End Function	

Sub txtFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtFrDt.focus
	End If
End Sub


Sub txtFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtToDt.focus
	End If
End Sub


Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13 Then Call FncQuery
End Sub

'================================================================================================================================
Function DBQueryOK()
	lgIntFlgMode = PopupParent.OPMD_CMODE
	Frm1.vspdData.Focus
End Function
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
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="14NXXU">&nbsp;&nbsp;&nbsp;&nbsp;
											   <INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 tag="14x"></TD>
						<TD CLASS="TD5" NOWRAP>모품목</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="모품목" NAME="txtItemCd" SIZE=10 MAXLENGTH=18 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">
											   <INPUT TYPE=TEXT ALT="모품목" NAME="txtItemNm" SIZE=20 tag="14x"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>적용유효일</TD>
						<TD CLASS="TD6">
									<script language=JavaScript>
										ExternalWrite('<OBJECT ALT=적용유효일 NAME="txtFrDt" classid=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD Title="FPDATETIME" tag="23N1"></OBJECT>');
									</script>&nbsp;~&nbsp;
									<script language=JavaScript>
										ExternalWrite('<OBJECT ALT=적용유효일 NAME="txtToDt" classid=<%=gCLSIDFPDT%> id=fpDateTime1 CLASS=FPDTYYYYMMDD Title="FPDATETIME" tag="23N1"></OBJECT>');
									</script>
								</TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_60%>>
				<TR>
					<TD CLASS="TD5" NOWRAP>BOM Type</TD>
					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="BOM Type" NAME="txtBomno" SIZE=34 tag="24X">
					<TD CLASS="TD5" NOWRAP></TD>
					<TD CLASS="TD6" NOWRAP></TD>
				</TR>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnFrDt" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnToDt" tag="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
