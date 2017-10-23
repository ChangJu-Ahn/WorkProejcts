<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 출하관리																	*
'*  3. Program ID           : I1311xa1_ko441     														*
'*  4. Program Name         : MES출고참조																*
'*  5. Program Desc         : 국내출고_반품등록을 위한 MES출고참조 	                                    *
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2008/02/21																*
'*  8. Modified date(Last)  : 2008/02/21																*
'*  9. Modifier (First)     : HAN cheol 																*
'* 10. Modifier (Last)      : HAN cheol     															*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              :                                       									*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>건설중인자산번호팝업</TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit

'========================================
Dim arrParent, strBuildAsstNo
ArrParent = window.dialogArguments
Set PopupParent  = arrParent(0)
    strBuildAsstNo	= arrParent(1)

top.document.title = PopupParent.gActivePRAspName


Const BIZ_PGM_ID = "a7102pb1.asp"

Dim C_AcctCd  
Dim C_AcctNm    
Dim C_BuildAsstNo   
Dim C_DrAmt    
Dim C_CrAmt     
Dim C_BalAmt    


Dim IsOpenPop      ' Popup


'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->

Dim arrReturn
Dim gblnWinEvent
Dim lgStrAllocInvFlag		' 재고할당 사용여부 

'========================================
Function InitVariables()
	lgIntGrpCount = 0										
	lgStrPrevKey = ""
	lgSortKey = 1										
	lgIntFlgMode = PopupParent.OPMD_CMODE										
		
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function

'========================================
Sub SetDefaultVal()

	txtBuildAsstNo.value = arrParent(1)

End Sub

'=========================================
Sub initSpreadPosVariables()
	C_AcctCd   		= 1
	C_AcctNm		= 2 
	C_BuildAsstNo   = 3
	C_DrAmt    		= 4
	C_CrAmt     	= 5
	C_BalAmt    	= 6
End Sub

'=====================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "PA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    
	
	With ggoSpread

		.Source = vspdData
		.Spreadinit "V20030901",,PopupParent.gAllowDragDropSpread    

		vspdData.MaxCols = C_BalAmt + 1
		vspdData.MaxRows = 0

		vspdData.ReDraw = False

		Call GetSpreadColumnPos("A")

		Call AppendNumberPlace("7","5","0")

		ggoSpread.SSSetEdit  C_AcctCd,		"계정코드",			10
		ggoSpread.SSSetEdit  C_AcctNm,		"계정명",			18		
		ggoSpread.SSSetEdit  C_BuildAsstNo,	"건설중인자산번호",		14
		ggoSpread.SSSetFloat C_DrAmt,		"발생금액",	14, PopupParent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_CrAmt,		"반제금액",	14, PopupParent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
		ggoSpread.SSSetFloat C_BalAmt,		"잔액",	14, PopupParent.ggAmtOfMoneyNo, ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gComNum1000,PopupParent.gComNumDec
		
    	Call .SSSetColHidden(vspdData.MaxCols,vspdData.MaxCols,True)

		
		vspdData.ReDraw = True
	End With

	Call SetSpreadLock

End Sub

'========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
	vspdData.OperationMode = 5	'Multi Select Mode
End Sub

'========================================
Function OKClick()
	on error resume next
	
	Dim intColCnt, intRowCnt, intInsRow, iLngSelectedRows



	With vspdData
		iLngSelectedRows = .SelModeSelCount
		' 전체선택시 
		If iLngSelectedRows = -1 Then
			iLngSelectedRows = .MaxRows
		End If

		If iLngSelectedRows > 0 Then 
			intInsRow = 0

			Redim arrReturn(iLngSelectedRows, .MaxCols)

			For intRowCnt = 1 To .MaxRows

				.Row = intRowCnt

				If .SelModeSelected Then

					.Col = C_BuildAsstNo    : arrReturn(intInsRow, 0) = .Text
'msgbox .text
					intInsRow = intInsRow + 1

				End IF

			Next
		End if			
	End With

	Self.Returnvalue = arrReturn
	Self.Close()
End Function

'========================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function




Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_AcctCd   		= iCurColumnPos(1)
			C_AcctNm   		= iCurColumnPos(2)			
			C_BuildAsstNo   = iCurColumnPos(3)
			C_DrAmt    		= iCurColumnPos(4)
			C_CrAmt     	= iCurColumnPos(5)
			C_BalAmt    	= iCurColumnPos(6)
    End Select    
End Sub

'========================================
Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
	Call ggoOper.LockField(Document, "N")						<% '⊙: Lock  Suitable  Field %>
	Call InitSpreadSheet()
	Call SetDefaultVal
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call InitVariables
	Call DbQuery()

End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If vspdData.ActiveRow > 0 Then	Call OKClick
End Function

'==========================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	
	gMouseClickStatus = "SPC"					'SpreadSheet 대상명이 vspdData일경우 
	Set gActiveSpdSheet = vspdData
    Call SetPopupMenuItemInf("0000111111")

    If vspdData.MaxRows <= 0 Then Exit Sub
   	    
    If Row = 0 Then
		vspdData.OperationMode = 0
        ggoSpread.Source = vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
	Else
		vspdData.OperationMode = 5
    End If
End Sub

'========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'========================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

'========================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	If OldLeft <> NewLeft Then Exit Sub

	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If CheckRunningBizProcess Then Exit Sub
		If lgStrPrevKey <> "" Then Call DbQuery
	End If		 

End Sub


'=====================================================
Function FncQuery() 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

	If Not chkField(Document, "1") Then Exit Function


	Call ggoOper.ClearField(Document, "2")
	Call InitVariables

    Call DbQuery

    FncQuery = True																
        
End Function

'=====================================================
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'=====================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=====================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'=====================================================
Function DbQuery()
	Err.Clear															

	DbQuery = False														

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal

		strVal = BIZ_PGM_ID & "?txtMode="           & PopupParent.UID_M0001	& _
    							"&txtBuildAsstNo="  & Trim(txtBuildAsstNo.value)		& _
    							"&txtMaxRows="        & vspdData.MaxRows			& _
    							"&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
'		strVal = BIZ_PGM_ID & "?txtMode="           & PopupParent.UID_M0001		& _
'								"&txtBuildAsstNo="        & Trim(txtBuildAsstNo.value)		& _
'								'"&txtBuildAsstNoNm="        & Trim(txtBuildAsstNoNm.value)		& _
'								"&txtMaxRows="        & vspdData.MaxRows			& _
'								"&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex

	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True														
End Function

'=====================================================
Function DbQueryOk()
	If vspdData.MaxRows > 0 Then
		If lgIntFlgMode <> PopupParent.OPMD_UMODE Then
			lgIntFlgMode = PopupParent.OPMD_UMODE
			vspdData.Row = 1
			vspdData.SelModeSelected = True		
		End If
		vspdData.Focus
	Else
		Call SetFocusToDocument("P")

	End If

End Function




</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">

<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR >
						<TD CLASS="TD5" NOWRAP>건설중인자산번호</TD>
						<TD CLASS="TD6" NOWRAP COLSPAN=3><INPUT TYPE="Text" NAME="txtBuildAsstNo" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="건설중인자산번호"></TD>
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
					<TD HEIGHT="100%">
						<script language =javascript src='./js/a7102pa1_OBJECT1_vspdData.js'></script>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX ="-1"></IFRAME></TD>
	</TR>
</TABLE>
<!--
<INPUT TYPE=HIDDEN NAME="txtHRetFlag" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHMovType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHSoType" tag="14">
<INPUT TYPE=HIDDEN NAME="txtHPlantCd" tag="14">

<INPUT TYPE=HIDDEN NAME="HFromDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HShipToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="HSlCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HItemCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="HSoNo" tag="24">
-->
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
