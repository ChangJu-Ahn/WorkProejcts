<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         : 사내재고이동정보 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2001/11/08
'*  8. Modified date(Last)  : 2003/06/03
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2001/11/08
'**********************************************************************************************-->
<HTML>
<HEAD>
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

Dim C_ItemCode
Dim C_ItemName
Dim C_TrackingNo
Dim C_qty
Dim C_BaseUnit

Const BIZ_PGM_QRY_ID    = "i1311rb1.asp"

Dim IsOpenPop
Dim arrReturn
Dim arrParam

Dim arrPlantCd
Dim arrPlantNm
Dim arrFrSlCd
Dim arrFrSlNm
Dim arrToSlCd
Dim arrToSlNm

arrParam        = window.dialogArguments
Set PopupParent = arrParam(0)

arrPlantCd    = arrParam(1)
arrPlantNm    = arrParam(2)
arrFrSlCd     = arrParam(3)
arrFrSlNm     = arrParam(4)
arrToSlCd     = arrParam(5)
arrToSlNm     = arrParam(6)

top.document.title = PopupParent.gActivePRAspName

'==========================================  2.1 InitVariables()  ======================================
Sub InitVariables()
    lgBlnFlgChgValue = False                              
    lgStrPrevKey     = ""                                 
    lgSortKey        = 1
    
      Redim arrReturn(0, 0)
    arrReturn(0,0)		= ""
	Self.Returnvalue	= arrReturn 
	
End Sub

'==========================================  2.2 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtPlantCd.value     = arrPlantCd
	frm1.txtPlantNm.value     = arrPlantNm
	frm1.txtSLCd1.value       = arrFrSlCd
	frm1.txtSLNm1.value       = arrFrSlNm
	frm1.txtSLCd2.value       = arrToSlCd
	frm1.txtSLNm2.value       = arrToSlNm
End Sub

'========================================  2.3 LoadInfTB19029()  =========================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","RA") %>
End Sub

'========================================= 2.6 InitSpreadSheet() =========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20060208", , PopupParent.gAllowDragDropSpread

	With frm1.vspdData
	    .OperationMode = 5
		.ReDraw = false
		.MaxCols = C_BaseUnit+1     
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetEdit  C_ItemCode,   "품목",         15
		ggoSpread.SSSetEdit  C_ItemName,   "품목명",       30
		ggoSpread.SSSetEdit  C_TrackingNo, "Tracking No.", 20
		ggoSpread.SSSetFloat C_Qty,        "수불수량",     18, PopupParent.ggQtyNo, ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gComNum1000, PopupParent.gComNumDec
		ggoSpread.SSSetEdit  C_BaseUnit,   "단위",		   10
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()

		ggoSpread.SSSetSplit2(2)
       .ReDraw = true
    End With
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_ItemCode   = 1
	C_ItemName   = 2
	C_TrackingNo = 3
	C_Qty        = 4
	C_BaseUnit   = 5
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

		C_ItemCode   = iCurColumnPos(1)
		C_ItemName   = iCurColumnPos(2)
		C_TrackingNo = iCurColumnPos(3)
		C_Qty        = iCurColumnPos(4)
		C_BaseUnit   = iCurColumnPos(5)
	End Select
End Sub

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
    
	Call InitVariables												
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call DbQuery()
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

Function vspdData_KeyPress(KeyAscii)
	On error Resume Next
	
	If KeyAscii = 27 Then Call CancelClick()

	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then 
		Call OKClick()
	Elseif KeyAscii = 27 Then
		Call CancelClick()
	End IF
End Function

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then Exit Sub
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) and lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
		DbQuery
	End if
End Sub

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'	설명 : Fnc함수명 으로 시작하는 모든 Function
'********************************************************************************************************* 
Function FncQuery() 

    FncQuery = False                                                        
 
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", PopupParent.VB_YES_NO,"X", "X")
		If IntRetCD = vbNo Then Exit Function
    End If
   
 	ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
	Call InitVariables 													
    
    If DbQuery() = False Then Exit Function

    FncQuery = True		
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
	Call parent.FncExport(C_MULTI)
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(C_MULTI , False)                                
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()	
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery()
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                        
	Call LayerShowHide(1)

    With frm1
		strVal = BIZ_PGM_QRY_ID &	"?txtPlantCd="		& Trim(.txtPlantCd.value)	& _
									"&txtPlantNm="		& Trim(.txtPlantNm.value)	& _
									"&txtSLCd1="		& Trim(.txtSLCd1.value)		& _		
									"&txtSLNm1="		& Trim(.txtSLNm1.value)		& _
									"&txtSLCd2="		& Trim(.txtSLCd2.value)		& _
									"&txtSLNm2="		& Trim(.txtSLNm2.value)		& _	
									"&lgStrPrevKey="	& lgStrPrevKey         
  
        Call RunMyBizASP(MyBizASP, strVal)					
    End With
 
    DbQuery = True
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()												
	frm1.vspdData.Focus
End Function

'========================================================================================
' Function Name : OKClick
'========================================================================================
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



</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
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
						<TD CLASS="TD5">공장</TD>
						<TD CLASS="TD6"><INPUT TYPE="Text" Name="txtPlantCd" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="공장" TABINDEX= "-1">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 tag="14" TABINDEX= "-1"></TD>
						<TD CLASS="TD5"></TD>
						<TD CLASS="TD6"></TD>
					</TR>
					<TR>
						<TD CLASS="TD5">창고</TD>
						<TD CLASS="TD6"><INPUT TYPE="Text" NAME="txtSLCd1" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="창고" TABINDEX= "-1">&nbsp;<INPUT TYPE="Text" NAME="txtSLNm1" SIZE=20 tag="14" TABINDEX= "-1"></TD>
						<TD CLASS="TD5">이동창고</TD>
						<TD CLASS="TD6"><INPUT TYPE="Text" NAME="txtSLCd2" SIZE=10 MAXLENGTH=7 tag="14XXXU" ALT="이동창고" TABINDEX= "-1">&nbsp;<INPUT TYPE="Text" NAME="txtSLNm2" SIZE=20 tag="24" TABINDEX= "-1"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR><TD WIDTH=100% HEIGHT=100% VALIGN=TOP>
		<TABLE <%=LR_SPACE_TYPE_20%>>
			<TR>
				<TD>
			<script language =javascript src='./js/i1311ra1_OBJECT1_vspdData.js'></script>
				</TD>
			</TR>
		</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> >
		</TD>
	</TR>				
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"></IMG></TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"></IMG>
					<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>
					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" TABINDEX= "-1"></IFRAME>		
		</TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
