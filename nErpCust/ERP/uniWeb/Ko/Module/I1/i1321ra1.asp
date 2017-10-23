<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : I1321RA1.ASP
'*  4. Program Name         : 사급품 출고예정정보 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/11/30
'*  8. Modified date(Last)  : 2003/06/03
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2002/11/30
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE>사급품 출고예정정보</TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBScript">

Option Explicit																	

<!-- #Include file="../../inc/lgvariables.inc" -->

Const BIZ_PGM_ID = "i1321rb1.asp"

Dim C_ItemCode    
Dim C_ItemName    
Dim C_TrackingNo  
Dim C_ReqmtQty	 
Dim C_ReqmtUnit	 
Dim C_IssueQty	 
Dim C_ToOnhandQty
Dim C_RequestQty
Dim C_FrOnhandQty
Dim C_BaseUnit	 
Dim C_ItemSpec	 

Dim arrReturn
Dim arrParent
Dim arrPlantCd
Dim arrPlantNm
Dim arrSlFrCd
Dim arrSlFrNm
Dim arrSlToCd
Dim arrSlToNm
Dim arrBpCd
Dim arrBpNm

arrParent		= window.dialogArguments

set PopupParent = arrParent(0)

arrPlantCd		= arrParent(1)
arrPlantNm		= arrParent(2)
arrSlFrCd       = arrParent(3)
arrSlFrNm       = arrParent(4)
arrSlToCd		= arrParent(5)
arrSlToNm		= arrParent(6)

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
	Redim arrReturn(0, 0)
	arrReturn(0,0)		= ""
	Self.Returnvalue	= arrReturn 
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtPlantCd.value 	= arrPlantCd
	frm1.txtPlantNm.value	= arrPlantNm	
	frm1.txtSLFrCd.value 	= arrSlFrCd
	frm1.txtSLFrNm.value	= arrSlFrNm
	frm1.txtSLToCd.value 	= arrSlToCd
	frm1.txtSLToNm.value	= arrSlToNm
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
	ggoSpread.Spreadinit "V20030423", , PopupParent.gAllowDragDropSpread

	With  frm1.vspdData
		.ReDraw = false
	    .OperationMode = 5
	    .MaxCols = C_ItemSpec+1												
	    .MaxRows = 0

		Call GetSpreadColumnPos("A")
		
	    ggoSpread.SSSetEdit  C_ItemCode,	"품목",				18, 0, -1, 18, 2
	    ggoSpread.SSSetEdit  C_ItemName,	"품목명",			25, 0, -1, 50
	    ggoSpread.SSSetEdit  C_TrackingNo,	"Tracking No.",		20, 0, -1, 25, 2
	    ggoSpread.SSSetFloat C_ReqmtQty,	"필요수량",			15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	    ggoSpread.SSSetEdit  C_ReqmtUnit,	"필요단위",			10, 0, -1, 3
	    ggoSpread.SSSetFloat C_IssueQty,	"출고수량",			15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	    ggoSpread.SSSetFloat C_ToOnhandQty,	"공급처재고수량",	15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	    ggoSpread.SSSetFloat C_RequestQty,	"출고예정수량",		15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	    ggoSpread.SSSetFloat C_FrOnhandQty,	"재고수량",			15, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
	    ggoSpread.SSSetEdit  C_BaseUnit,	"재고단위",			10, 0, -1, 3
	    ggoSpread.SSSetEdit  C_ItemSpec,	"규격",				20, 0, -1, 50
	    
	    Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	
	    ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetSplit2(2)
		
		.ReDraw = true
    End With
    
End Sub

'================================== 2.2.4 InitSpreadPosVariables() ==================================================
Sub InitSpreadPosVariables()
	C_ItemCode     = 1										
	C_ItemName     = 2
	C_TrackingNo   = 3
	C_ReqmtQty	   = 4
	C_ReqmtUnit	   = 5
	C_IssueQty	   = 6
	C_ToOnhandQty  = 7
	C_RequestQty   = 8
	C_FrOnhandQty  = 9
	C_BaseUnit	   = 10
	C_ItemSpec	   = 11
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = frm1.vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		
		C_ItemCode     = iCurColumnPos(1)
    	C_ItemName     = iCurColumnPos(2)
	    C_TrackingNo   = iCurColumnPos(3)
	    C_ReqmtQty	   = iCurColumnPos(4)
	    C_ReqmtUnit	   = iCurColumnPos(5)
	    C_IssueQty	   = iCurColumnPos(6)
	    C_ToOnhandQty  = iCurColumnPos(7)
	    C_RequestQty   = iCurColumnPos(8)
	    C_FrOnhandQty  = iCurColumnPos(9)
	    C_BaseUnit	   = iCurColumnPos(10)
	    C_ItemSpec	   = iCurColumnPos(11)
	End Select
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
	Dim intInsRow, intColCnt
	Dim i
    Dim ret

    With frm1.vspdData
        
        If.IsBlockSelected Or .SelModeSelCount Then
			
			ReDim arrReturn(.SelModeSelCount - 1, .MaxCols - 1)
            intInsRow = 0
            For i = 0 To .SelModeSelCount - 1
            
                ret = .GetMultiSelItem(ret)
            
                If ret = -1 Then Exit For
            
                .Row = ret
                For intColCnt = 0 To .MaxCols - 1
                    .Col = intColCnt + 1
                    arrReturn(intInsRow, intColCnt) = .Text
                Next
                intInsRow = intInsRow + 1
            Next
        End If
    End With
    
	Self.Returnvalue = arrReturn
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
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    

    Call InitSpreadSheet
    Call InitVariables                                                 
    Call SetDefaultVal()
    
    If DbQuery = False Then Exit Sub
	
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	
	gMouseClickStatus = "SPC"   
   
	Set gActiveSpdSheet = frm1.vspdData
   
	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	
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
End Sub

'========================================================================================
' Function Name : vspdData_DblClick
'========================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
   
	If Row <= 0 Then Exit Sub
	If frm1.vspdData.MaxRows = 0 Then Exit Sub

	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			Call OKClick
		End If
	End If
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub 

'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub 

'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.SaveSpreadColumnInf()
End Sub 

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
   ggoSpread.Source = gActiveSpdSheet
   Call ggoSpread.RestoreSpreadInf()
   Call InitSpreadSheet
   Call ggoSpread.ReOrderingSpreadData
End Sub 

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
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 

    FncQuery = False                                                   
    Err.Clear                                                          
 	
 	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData

    Call InitVariables 													

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
							"&txtSlFrCd="			& Trim(arrSlFrCd)		& _
							"&txtSlFrNm="			& Trim(arrSlFrNm)		& _
							"&txtSLToCd="			& Trim(arrSlToCd)		& _
							"&txtSLToNm="			& Trim(arrSlToNm)		& _
							"&txtMaxRows="			& frm1.vspdData.MaxRows	& _
							"&lgStrPrevKeyIndex="	& lgStrPrevKeyIndex
    
    Call RunMyBizASP(MyBizASP, strVal)								
    
	DbQuery = True

End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()													
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
				<TD CLASS="TD5">공장</TD>
				<TD CLASS="TD6"><INPUT TYPE="Text" Name="txtPlantCd" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="공장" TABINDEX=-1>&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14" TABINDEX=-1></TD>
				<TD CLASS="TD5">창고</TD>
				<TD CLASS="TD6"><INPUT TYPE="Text" NAME="txtSLFrCd" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="창고" TABINDEX=-1>&nbsp;<INPUT TYPE="Text" NAME="txtSLFrNm" SIZE=20 tag="14" TABINDEX=-1></TD>
			</TR>
			<TR>
				<TD CLASS="TD5">공급처창고</TD>
				<TD CLASS="TD6"><INPUT TYPE="Text" NAME="txtSLToCd" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="공급처창고" TABINDEX=-1>&nbsp;<INPUT TYPE="Text" NAME="txtSLToNm" SIZE=20 tag="14" TABINDEX=-1></TD>
				<TD CLASS="TD5">공급처</TD>
				<TD CLASS="TD6"><INPUT TYPE="Text" NAME="txtBPCd" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="공급처" TABINDEX=-1>&nbsp;<INPUT TYPE="Text" NAME="txtBpNm" SIZE=20 tag="14" TABINDEX=-1></TD>
			</TR>
		</TABLE>
		</FIELDSET>
		</TD>
	</TR>
	<TR><TD HEIGHT=100%>
		<script language =javascript src='./js/i1321ra1_OBJECT1_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>		
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


