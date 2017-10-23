<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1.  Module Name			: Production
'*  2.  Function Name		: 
'*  3.  Program ID			: p2215ma1.asp
'*  4.  Program Name		: List MPS Requirement
'*  5.  Program Desc		: List MPS Requirement which is made by sales order or(and) sales plan
'*  6.  Business ASP List	: List MPS Requirement
'*  7.  Modified date(First):
'*  8.  Modified date(Last)	: 
'*  9.  Modifier (First)	:
'* 10. Modifier (Last)		: Jung Yu Kyung
'* 11. Comment				: 
'* 12. History              : Tracking No 9자리에서 25자리로 변경(2003.03.03)
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

'===========================================================================================================
Const BIZ_PGM_ID = "p2215mb1.asp"
'===========================================================================================================
Dim C_ItemCode
Dim C_ItemNm
Dim C_ItemSpec
Dim C_TrackingNo
Dim C_ReqrdQty
Dim C_IssuedQty
Dim C_BasicUnit
Dim C_ReqrdDt
Dim C_ReqType
Dim C_ReqStatus
Dim C_SONo
Dim C_SOSeq
Dim C_ItemGroupCd
Dim C_ItemGroupNm

Const C_SHEETMAXROWS = 30

Dim StartDate
Dim LastDate

StartDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
LastDate =  UNIDateAdd("m",1,StartDate,parent.gDateFormat)

'===========================================================================================================
<!-- #Include file="../../inc/lgVariables.inc" -->	

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim IsOpenPop          
'===========================================================================================================
Sub initSpreadPosVariables()  
	C_ItemCode      =  1
	C_ItemNm        =  2
	C_ItemSpec		=  3
	C_TrackingNo    =  4
	C_ReqrdQty		=  5
	C_IssuedQty		=  6
	C_BasicUnit		=  7
	C_ReqrdDt 		=  8
    C_ReqType 		=  9
    C_ReqStatus		=  10
    C_SONo 			=  11
    C_SOSeq 		=  12
    C_ItemGroupCd	=  13
    C_ItemGroupNm	=  14
End Sub

'===========================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0

    lgStrPrevKey1 = ""
    lgStrPrevKey2 = ""
    lgLngCurRows = 0
    lgSortKey    = 1
End Sub

'===========================================================================================================
Sub SetDefaultVal()
	frm1.txtFromReqrdDt.text	= StartDate
	frm1.txtToReqrdDt.text	= LastDate
End Sub

'===========================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call LoadInfTB19029A("Q", "P", "NOCOOKIE", "MA") %>
End Sub

'===========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	
	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    

	.ReDraw = false
		
	.MaxCols = C_ItemGroupNm + 1 
	.MaxRows = 0
    
    Call AppendNumberPlace("6", "6", "0")
	
    Call GetSpreadColumnPos("A")

	ggoSpread.SSSetEdit 	C_ItemCode,		"품목",		18
	ggoSpread.SSSetEdit		C_ItemNm,		"품목명",	25
	ggoSpread.SSSetEdit		C_ItemSpec,		"규격",	25
	ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.", 25
	ggoSpread.SSSetFloat	C_ReqrdQty,		"계획수량", 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetFloat	C_IssuedQty,	"출고수량", 15,Parent.ggQtyNo,ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetEdit		C_BasicUnit,	"단위", 7                                  
	ggoSpread.SSSetDate		C_ReqrdDt,		"필요일", 11, 2, gDateFormat	
	ggoSpread.SSSetEdit		C_ReqType,		"생성구분", 10
	ggoSpread.SSSetEdit		C_ReqStatus,	"Status", 12
	ggoSpread.SSSetEdit		C_SONo,			"수주번호", 18
	ggoSpread.SSSetFloat	C_SOSeq,		"수주순번", 8,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec
	ggoSpread.SSSetEdit 	C_ItemGroupCd,	"품목그룹",		15
	ggoSpread.SSSetEdit		C_ItemGroupNm,	"품목그룹명",	30
	
	Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	ggoSpread.SSSetSplit2(1)
	
	.ReDraw = true
	
	Call SetSpreadLock 
    
    End With
    
End Sub
'===========================================================================================================
Sub SetSpreadLock()
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
        
End Sub
'===========================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_ItemCode		= iCurColumnPos(1)
			C_ItemNm		= iCurColumnPos(2)
			C_ItemSpec		= iCurColumnPos(3)
			C_TrackingNo	= iCurColumnPos(4)
			C_ReqrdQty		= iCurColumnPos(5)
			C_IssuedQty		= iCurColumnPos(6)
			C_BasicUnit		= iCurColumnPos(7)
			C_ReqrdDt		= iCurColumnPos(8)
			C_ReqType		= iCurColumnPos(9)
			C_ReqStatus		= iCurColumnPos(10)
			C_SONo			= iCurColumnPos(11)
			C_SOSeq			= iCurColumnPos(12)
			C_ItemGroupCd	= iCurColumnPos(13)
			C_ItemGroupNm	= iCurColumnPos(14)
			
    End Select    

End Sub
'===========================================================================================================
Sub InitComboBox()
    Call SetCombo(frm1.cboReqStatus, "AC", "Accepted")
    Call SetCombo(frm1.cboReqStatus, "RQ", "Requested")
End Sub
'===========================================================================================================
Function OpenItemInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6)
	Dim iCalledAspName, IntRetCD
		
	If IsOpenPop = True Then Exit Function
	
	If frm1.txtPlantCd.value = "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = strCode
	arrParam(2) = "12!MO"			' Combo Set Data:"1020!MP" -- 품목계정 구분자 조달구분 
	arrParam(3) = ""
	
	arrField(0) = 1 '"ITEM_CD"
	arrField(1) = 2 '"ITEM_NM"
    
	iCalledAspName = AskPRAspName("B1B11PA3")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "B1B11PA3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam, arrField), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetItemInfo(arrRet)
	End If	

End Function
'===========================================================================================================
Function OpenPlant()

    Dim arrRet
    Dim arrParam(5), arrField(6), arrHeader(6)

    If IsOpenPop = True Then Exit Function

    IsOpenPop = True

    arrParam(0) = "공장팝업"	
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
	Exit Function
    Else
	Call SetPlant(arrRet)
    End If	

End Function
'===========================================================================================================
Function OpenTrackingInfo()

	Dim iCalledAspName, IntRetCD
	
	If frm1.txtPlantCd.value= "" Then
		Call DisplayMsgBox("971012","X", "공장","X")
		frm1.txtPlantCd.focus 
		Set gActiveElement = document.activeElement
		IsOpenPop = False 
		Exit Function
	End If

	Dim arrRet
	Dim arrParam(4)

	If IsOpenPop = True Or UCase(frm1.txtTrackingNo.className) = "PROTECTED" Then Exit Function

	IsOpenPop = True

	arrParam(0) = Trim(frm1.txtPlantCd.value)
	arrParam(1) = Trim(frm1.txtTrackingNo.value)
	arrParam(2) = Trim(frm1.txtItemCd.value)
	
	iCalledAspName = AskPRAspName("P4600PA1")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "P4600PA1", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName, Array(Window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTrackingNo(arrRet)
	End If
	
End Function
'===========================================================================================================
Function OpenItemGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtItemGroupCd.className) = UCase(parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "품목그룹팝업"
	arrParam(1) = "B_ITEM_GROUP"
	arrParam(2) = Trim(UCase(frm1.txtItemGroupCd.Value))
	arrParam(3) = ""
	arrParam(4) = "DEL_FLG = " & FilterVar("N", "''", "S") & " "
	arrParam(5) = "품목그룹"
	 
	arrField(0) = "ITEM_GROUP_CD"
	arrField(1) = "ITEM_GROUP_NM"
	    
	arrHeader(0) = "품목그룹"
	arrHeader(1) = "품목그룹명"
	    
	arrRet = window.showModalDialog("../../ComAsp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	IsOpenPop = False
	 
	If arrRet(0) <> "" Then
		Call SetItemGroup(arrRet)
	End If 
	
	Call SetFocusToDocument("M")
	frm1.txtItemGroupCd.focus
 
End Function
'===========================================================================================================
Function SetItemInfo(ByRef arrRet)
    With frm1
		.txtItemCd.value = arrRet(0)
		.txtItemNm.value = arrRet(1)
		.txtItemCd.focus
		Set gActiveElement = document.activeElement
    End With
End Function
'===========================================================================================================
Function SetPlant(ByRef arrRet)
    frm1.txtPlantCd.Value = arrRet(0)
    frm1.txtPlantNm.value = arrRet(1)
    frm1.txtPlantCd.focus
    Set gActiveElement = document.activeElement
End Function
'===========================================================================================================
Function SetTrackingNo(ByRef arrRet)
	frm1.txtTrackingNo.Value = arrRet(0)
	frm1.txtTrackingNo.focus
	Set gActiveElement = document.activeElement
End Function
'===========================================================================================================
Function SetItemGroup(byval arrRet)
	frm1.txtItemGroupCd.Value    = arrRet(0)  
	frm1.txtItemGroupNm.Value    = arrRet(1)  
End Function

'===========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")

    Call InitSpreadSheet
	Call SetDefaultVal
    Call InitVariables
    Call InitComboBox
    Call SetToolBar("11000000000011")

    If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Parent.gPlant
		frm1.txtPlantNm.value = Parent.gPlantNm
	
		frm1.txtItemCd.focus 
	Else
		frm1.txtPlantCd.focus
	End If   
	
	Set gActiveElement = document.activeElement
  
End Sub


'===========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	
    Call SetPopupMenuItemInf("0000111111")

	gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows = 0 Then
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
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       Exit Sub
    End If
    
End Sub


'===========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    
End Sub


'===========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub


'===========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'===========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True

End Sub


'===========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
   
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey1 <> "" Then
			
			Call DisableToolBar(Parent.TBC_QUERY)
            If DBQuery = False Then 
               Call RestoreToolBar()
               Exit Sub
            End If 
		End If
    End if
End Sub


'===========================================================================================================
Sub txtFromReqrdDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromReqrdDt.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtFromReqrdDt.Focus
    End If
End Sub


'===========================================================================================================
Sub txtToReqrdDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToReqrdDt.Action = 7
        Call SetFocusToDocument("M")
		frm1.txtToReqrdDt.Focus
    End If
End Sub


'===========================================================================================================
Sub txtFromReqrdDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub


'===========================================================================================================
Sub txtToReqrdDt_KeyDown(keycode, shift)
	If keycode = 13 Then
		Call MainQuery()
	End If
End Sub


'===========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub


'===========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet

    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()

End Sub


'===========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False
    
    Err.Clear

    If frm1.txtPlantCd.value = "" Then
		frm1.txtPlantNm.value = ""
	End If
	
	If frm1.txtItemCd.value = "" Then
		frm1.txtItemNm.value = ""
	End If

    Call ggoOper.ClearField(Document, "2")  

    Call InitVariables
    
    If Not chkField(Document, "1") Then
       Exit Function
    End If
    
    If ValidDateCheck(frm1.txtFromReqrdDt, frm1.txtToReqrdDt)  = False Then		
		Exit Function
	End If
    
    If DbQuery = False Then
		Exit Function
	End If
       
    FncQuery = True
    
End Function


'===========================================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)
End Function


'===========================================================================================================
Function FncPrint() 
    Call parent.fncPrint()
End Function


'===========================================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI) 
End Function


'===========================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    Call LayerShowHide(1)
    
    Err.Clear

    Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(.hPlantCd.value)
		strVal = strVal & "&cboReqStatus=" & Trim(.hReqStatus.value)
		strVal = strVal & "&txtFromReqrdDt=" & Trim(.hFromReqrdDt.value)
		strVal = strVal & "&txtToReqrdDt=" & Trim(.hToReqrdDt.value)
		strVal = strVal & "&txtItemCd=" & Trim(.hItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.hTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.hItemGroupCd.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
		strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
		strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		strVal = strVal & "&lgIntFlgMode=" & lgIntFlgMode
		strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		strVal = strVal & "&cboReqStatus=" & Trim(.cboReqStatus.value)
		strVal = strVal & "&txtFromReqrdDt=" & Trim(.txtFromReqrdDt.Text)
		strVal = strVal & "&txtToReqrdDt=" & Trim(.txtToReqrdDt.Text)
		strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)
		strVal = strVal & "&txtItemGroupCd=" & Trim(.txtItemGroupCd.value)
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
	
	End IF	

	Call RunMyBizASP(MyBizASP, strVal)

    End With
    
    DbQuery = True
    
End Function


'===========================================================================================================
Function DbQueryOk()
	
	Call SetToolBar("11000000000111")
    
    lgIntFlgMode = Parent.OPMD_UMODE
	lgBlnFlgChgValue = False    
    
    Call ggoOper.LockField(Document, "Q")
	frm1.vspdData.focus
End Function


'===========================================================================================================
Function FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
    
End Function


Function FncExit()
	FncExit = True
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE  <%=LR_SPACE_TYPE_00%>>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>MPS요청조회</font></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6 NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()" OnMouseOver="vbscript:PopupMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>필요일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/p2215ma1_I318591934_txtFromReqrdDt.js'></script>
										&nbsp;~&nbsp;
										<script language =javascript src='./js/p2215ma1_I957299232_txtToReqrdDt.js'></script>
									</TD>																						
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=18 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemInfo frm1.txtItemCd.value">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>Tracking No.</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=25 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No."><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingInfo"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>품목그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtItemGroupCd" SIZE=15 MAXLENGTH=10 tag="11XXXU"  ALT="품목그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemGroupCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenItemGroup()">&nbsp;<INPUT TYPE=TEXT NAME="txtItemGroupNm" SIZE=30 MAXLENGTH=40 tag="14" ALT="품목그룹명"></TD>
									<TD CLASS=TD5 NOWRAP>Status</TD>
									<TD CLASS=TD6 NOWRAP><SELECT NAME="cboReqStatus" tag="11X" STYLE="WIDTH: 115px;" ALT="Status"><OPTION value=""></OPTION></SELECT></TD>
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
							<TD HEIGHT="100%" colspan=4>
								<script language =javascript src='./js/p2215ma1_I193368500_vspdData.js'></script>
							</TD>
						</TR>						
					</TABLE>
					</TD>
				</TR>
			</TABLE>
		</DIV>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm "  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="hPlantCd" tag="24"><INPUT TYPE=HIDDEN NAME="hReqStatus" tag="24">
<INPUT TYPE=HIDDEN NAME="hFromReqrdDt" tag="24"><INPUT TYPE=HIDDEN NAME="hToReqrdDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemCd" tag="24"><INPUT TYPE=HIDDEN NAME="hTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="hItemGroupCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
