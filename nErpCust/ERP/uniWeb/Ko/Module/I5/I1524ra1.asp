<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : I1524RA1.ASP
'*  4. Program Name         : VMI 입고현황 참조화면 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/13
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2003/01/13
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit																

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID = "i1524rb1.asp"

Dim C_BpCd
Dim C_BpNm
Dim C_ItemCode    
Dim C_ItemName
Dim C_ItemUnit
Dim C_Qty
Dim C_VMISlCd
Dim C_SlCd
Dim C_TrackingNo
Dim C_LotNo
Dim C_LotSubNo
Dim C_ItemSpec

Dim arrReturn
Dim arrParent
Dim arrPlantCd
Dim arrPlantNm
Dim arrRcptNo

arrParent		= window.dialogArguments

set PopupParent = arrParent(0)

arrPlantCd		= arrParent(1)
arrPlantNm		= arrParent(2)
arrRcptNo		= arrParent(3)

top.document.title = PopupParent.gActivePRAspName

Dim lgOldRow
Dim gblnWinEvent
Dim strReturn

Dim IsOpenPop          

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
	lgStrPrevKeyIndex	= ""
	lgLngCurRows		= 0
	lgSortKey			= 1
    Self.Returnvalue	= Array("")  
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtPlantCd.value 	= arrPlantCd
	frm1.txtPlantNm.value	= arrPlantNm	
	frm1.txtRcptNo.value 	= arrRcptNo
End Sub 

'========================================================================================
' Function Name : LoadInfTB19029
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "I","NOCOOKIE","RA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030428", , PopupParent.gAllowDragDropSpread

	With  frm1.vspdData
		.ReDraw = false
	    .OperationMode = 3
	    .MaxCols = C_ItemSpec+1														
	    .MaxRows = 0

		Call GetSpreadColumnPos("A")
		
	    ggoSpread.SSSetEdit	 C_BpCd,			"공급처",		10
	    ggoSpread.SSSetEdit  C_BpNm,			"공급처명",		25
	    ggoSpread.SSSetEdit  C_ItemCode,		"품목",			18
	    ggoSpread.SSSetEdit  C_ItemName,		"품목명",		25
	    ggoSpread.SSSetEdit  C_ItemUnit,		"단위",			10
	    ggoSpread.SSSetFloat C_Qty,				"수량",			15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	    ggoSpread.SSSetEdit	 C_VMISlCd,			"VMI창고",		10
	    ggoSpread.SSSetEdit  C_SlCd,			"입고창고",		10 
	    ggoSpread.SSSetEdit  C_TrackingNo,		"Tracking No.", 20
	    ggoSpread.SSSetEdit  C_LotNo,			"Lot No.",		10
	    ggoSpread.SSSetEdit  C_LotSubNo,		"Lot 순번",		4
	    ggoSpread.SSSetEdit  C_ItemSpec,		"규격",			17
	    
	    Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		    		
  		ggoSpread.SSSetSplit2(2)  
		
		.ReDraw = true
    End With
End Sub

'================================== 2.2.4 InitSpreadPosVariables() ==================================================
Sub InitSpreadPosVariables()
	C_BpCd			= 1
	C_BpNm			= 2
	C_ItemCode		= 3										
	C_ItemName		= 4
	C_ItemUnit		= 5
	C_Qty			= 6
	C_VMISlCd		= 7
	C_SlCd			= 8
	C_TrackingNo	= 9
	C_LotNo			= 10
	C_LotSubNo		= 11
	C_ItemSpec		= 12

End Sub

'================================== 2.2.4 GetSpreadColumnPos() ==================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
		C_BpCd			= iCurColumnPos(1)
	    C_BpNm			= iCurColumnPos(2)
		C_ItemCode		= iCurColumnPos(3)
    	C_ItemName		= iCurColumnPos(4)
	    C_ItemUnit		= iCurColumnPos(5)
	    C_Qty			= iCurColumnPos(6)
	    C_VMISlCd		= iCurColumnPos(7)
	    C_SlCd			= iCurColumnPos(8)
	    C_TrackingNo	= iCurColumnPos(9)
	    C_LotNo			= iCurColumnPos(10)
	    C_LotSubNo		= iCurColumnPos(11) 		
	    C_ItemSpec		= iCurColumnPos(12)
	End Select
End Sub

'=========================================  2.3.1 OKClick()  ========================================
Function OKClick()		
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()		
	Self.Close()
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")                                         
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    

    Call InitSpreadSheet
    Call InitVariables                                                    
    Call SetDefaultVal()
    
    If DbQuery = False Then	Exit Sub
End Sub

'******************************  3.2.1 Object Tag 처리  *********************************************
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Function
	
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then
		If lgStrPrevKeyIndex <> "" Then
			If DbQuery = False Then Exit Sub
		End if
	End if
End Sub

'==========================================================================================
'   Event Name : vspdData_KeyPress(KeyAscii)
'==========================================================================================
Function vspdData_KeyPress(KeyAscii)
	On error Resume Next
	If KeyAscii = 27 Then 
		Call CancelClick()
	End if
	
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	Elseif KeyAscii = 27 Then
		Call CancelClick()
	End IF
End Function

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 

    FncQuery = False                                                       
    Err.Clear                                                              
    '-----------------------
    'Erase contents area
    '-----------------------
	ggoSpread.source = frm1.vspddata
	ggoSpread.ClearSpreadData 
    Call InitVariables 														
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery() = False Then Exit Function
    FncQuery = True															
    
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    
    Call LayerShowHide(1)
    DbQuery = False
    Err.Clear                                                           
    Dim strVal
    
    strVal = BIZ_PGM_ID	&	"?txtPlantCd="			& Trim(arrPlantCd)		& _
							"&txtPlantNm="			& Trim(arrPlantNm)		& _
							"&txtRcptNo="			& Trim(arrRcptNo)		& _
							"&txtMaxRows="			& frm1.vspdData.MaxRows	& _
							"&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex
    
    Call RunMyBizASP(MyBizASP, strVal)									
    
	DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()														
    '-----------------------
    'Reset variables area
    '-----------------------
	frm1.vspdData.Focus
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="GET">


<TABLE CELLSPACING=0 CLASS="basicTB">
	<TR><TD HEIGHT=40>
		<FIELDSET>
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
			<TR>
				<TD CLASS="TD5" NOWRAP>공장</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=6 MAXLENGTH=4 tag="14XXXU" ALT="공장" TABINDEX="-1">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=30 tag="14" TABINDEX="-1"></TD>
				<TD CLASS="TD5" NOWRAP>구매입고번호</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtRcptNo" SIZE=20 MAXLENGTH=18 tag="14XXXU" ALT="구매입고번호" TABINDEX="-1"></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>수불번호</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtItemDocumentNo" SIZE=20 MAXLENGTH=18 tag="14XXXU" ALT="수불번호" TABINDEX="-1"></TD>
				<TD CLASS="TD5" NOWRAP>수불구분</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtTrnsType" SIZE=20 MAXLENGTH=10 tag="14XXXU" ALT="수불구분" TABINDEX="-1"></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>수불일자</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtDocumentDt" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="수불일자" TABINDEX="-1"></TD>
				<TD CLASS="TD5" NOWRAP></TD>
				<TD CLASS="TD6" NOWRAP></TD>
			</TR>
		</TABLE>
		</FIELDSET>
		</TD>
	</TR>
	<TR><TD HEIGHT=100%>
		<script language =javascript src='./js/i1524ra1_OBJECT1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
	
				<TD WIDTH=70% NOWRAP></TD>
				<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
				<TD WIDTH=10>&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


