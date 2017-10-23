<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : I1523PA1.ASP
'*  4. Program Name         : VMI 현재고현황 팝업 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/13
'*  8. Modified date(Last)  : 2003/04/25
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Ahn Jung Je
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

Const BIZ_PGM_ID = "i1523pb1.asp"

Dim C_ItemCode    
Dim C_ItemName
Dim C_GoodQty
Dim C_ItemUnit
Dim C_TrackingNo
Dim C_LotNo
Dim C_LotSubNo
Dim C_ItemSpec
Dim C_LotFlg
Dim C_TrackingFlg    
Dim C_RecvInspFlg    

Dim arrReturn
Dim arrParent
Dim arrPlantCd
Dim arrPlantNm
Dim arrSlCd
Dim arrSlNm
Dim arrBpCd
Dim arrBpNm
Dim arrItemCd
Dim arrItemNm

arrParent		= window.dialogArguments

set PopupParent = arrParent(0)

arrPlantCd		= arrParent(1)
arrPlantNm		= arrParent(2)
arrSlCd			= arrParent(3)
arrSlNm			= arrParent(4)
arrBpCd			= arrParent(5)
arrBpNm			= arrParent(6)
arrItemCd		= arrParent(7)	
arrItemNm		= arrParent(8)

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
	Redim arrReturn(0)
    Self.Returnvalue	= arrReturn  
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "I","NOCOOKIE","PA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtPlantCd.value	= arrPlantCd
	frm1.txtPlantNm.value	= arrPlantNm
	frm1.txtSlCd.value		= arrSlCd
	frm1.txtSlNm.value		= arrSlNm
	frm1.txtBpCd.value		= arrBpCd
	frm1.txtBpNm.value		= arrBpNm
	frm1.txtItemCd.value	= arrItemCd
	frm1.txtItemNm.value	= arrItemNm
End Sub 

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030409"

	With  frm1.vspdData
		.ReDraw = false
	    .OperationMode = 3
	    .MaxCols = C_RecvInspFlg+1												
	    .MaxRows = 0

		Call GetSpreadColumnPos("A")
		
	    ggoSpread.SSSetEdit		C_ItemCode,		"품목",			18
	    ggoSpread.SSSetEdit		C_ItemName,		"품목명",		25
	    ggoSpread.SSSetFloat	C_GoodQty,		"재고수량",		15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetEdit		C_ItemUnit,		"단위",			6
		ggoSpread.SSSetEdit		C_TrackingNo,	"Tracking No.",	20
		ggoSpread.SSSetEdit		C_LotNo,		"Lot No.",		10
		ggoSpread.SSSetEdit		C_LotSubNo,		"순번",			6
	    ggoSpread.SSSetEdit		C_ItemSpec,		"규격",			25
	    ggoSpread.SSSetEdit		C_LotFlg,		"LotFlg",			5
	    ggoSpread.SSSetEdit		C_TrackingFlg,	"TrackingFlg",		5
	    ggoSpread.SSSetEdit		C_RecvInspFlg,	"TrackingFlg",		5

	    Call ggoSpread.SSSetColHidden(C_LotFlg, .MaxCols, True)
	    
	    ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetSplit(2)
		
		.ReDraw = true
    End With
End Sub

'========================================================================================================
' Name : InitSpreadPosVariables()	
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ItemCode		= 1										
	C_ItemName		= 2
	C_GoodQty		= 3
	C_ItemUnit		= 4
	C_TrackingNo	= 5
	C_LotNo			= 6
	C_LotSubNo		= 7
	C_ItemSpec		= 8
	C_LotFlg		= 9
	C_TrackingFlg   = 10 
	C_RecvInspFlg	= 11
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
		C_ItemCode		= iCurColumnPos(1)
    	C_ItemName		= iCurColumnPos(2)
	    C_GoodQty		= iCurColumnPos(3)
	    C_ItemUnit		= iCurColumnPos(4)
	    C_TrackingNo	= iCurColumnPos(5)
	    C_LotNo			= iCurColumnPos(6)
	    C_LotSubNo		= iCurColumnPos(7)
	    C_ItemSpec		= iCurColumnPos(8)
		C_LotFlg		= iCurColumnPos(9)
		C_TrackingFlg   = iCurColumnPos(10)
		C_RecvInspFlg	= iCurColumnPos(11)
	End Select
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
With frm1.vspdData 
  
	If .ActiveRow > 0 Then 
		Redim arrReturn(.MaxCols - 1)
  
		.Row = .ActiveRow
     
		.Col = C_ItemCode
		arrReturn(0) = .Text
		.Col = C_ItemName
		arrReturn(1) = .Text
		.Col = C_GoodQty
		arrReturn(2) = .Text
		.Col = C_ItemUnit
		arrReturn(3) = .Text
		.Col = C_TrackingNo
		arrReturn(4) = .Text
		.Col = C_LotNo
		arrReturn(5) = .Text
		.Col = C_LotSubNo
		arrReturn(6) = .Text
		.Col = C_ItemSpec
		arrReturn(7) = .Text 
		.Col = C_LotFlg
		arrReturn(8) = .Text 
		.Col = C_TrackingFlg
		arrReturn(9) = .Text 
		.Col = C_RecvInspFlg
		arrReturn(10) = .Text 
		  
		Self.Returnvalue = arrReturn
	End If
End With
Self.Close()

End Function

'=========================================  2.3.2 CancelClick()  ========================================
'=	Name : CancelClick()																				=
'========================================================================================================
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
    Call InitSpreadSheet()
    
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
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and lgStrPrevKeyIndex <> "" Then
		If DbQuery = False Then Exit Sub
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
	Call InitVariables() 												
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
    
    strVal = BIZ_PGM_ID	&	"?txtPlantCd="			& Trim(frm1.txtPlantCd.value)	& _
							"&txtSlCd="				& Trim(frm1.txtSlCd.value)		& _
							"&txtBpCd="				& Trim(frm1.txtBpCd.value)		& _
							"&txtItemCd="			& Trim(frm1.txtItemCd.value)	& _
							"&txtItemNm="			& Trim(frm1.txtItemNm.value)	& _
							"&txtMaxRows="			& frm1.vspdData.MaxRows			& _
							"&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex
    
    Call RunMyBizASP(MyBizASP, strVal)									
    
	DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()													
    Call ggoOper.LockField(Document, "Q")								
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
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="14XXXU" ALT="공장" TABINDEX="-1">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 MAXLENGTH=40 tag="14XXXU" ALT="공장명" TABINDEX="-1"></TD>
				<TD CLASS="TD5" NOWRAP>VMI 창고</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtSlCd" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="창고" TABINDEX="-1">&nbsp;<INPUT TYPE=TEXT NAME="txtSlNm" SIZE=20 MAXLENGTH=40 tag="14XXXU" ALT="창고명" TABINDEX="-1"></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>공급처</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtBpCd" SIZE=10 MAXLENGTH=10 tag="14XXXU" ALT="공급처" TABINDEX="-1">&nbsp;<INPUT TYPE=TEXT NAME="txtBpNm" SIZE=20 MAXLENGTH=40 tag="14XXXU" ALT="공급처명" TABINDEX="-1"></TD>
				<TD CLASS="TD5" NOWRAP>품목</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="품목" TABINDEX="-1">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=18 MAXLENGTH=40 tag="11XXXU" ALT="품목명" TABINDEX="-1"></TD>
			</TR>
		</TABLE>
		</FIELDSET>
		</TD>
	</TR>
	<TR><TD HEIGHT=100%>
		<script language =javascript src='./js/i1523pa1_OBJECT1_vspdData.js'></script>
	</TD></TR>
	<TR><TD HEIGHT=20>
	
		<TABLE CLASS="basicTB" CELLSPACING=0>
			<TR>
				<TD WIDTH=10>&nbsp;</TD>
	
				<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="okclick()"    ></IMG>
						                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG></TD>
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


