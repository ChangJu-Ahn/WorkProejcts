<%@ LANGUAGE="VBSCRIPT" %>
<!--'********************************************************************************************************
'*  1. Module Name          : Inventory              
'*  2. Function Name        : DocumentNo Popup       
'*  3. Program ID           :   i1111pa1.asp         
'*  4. Program Name         :                  
'*  5. Program Desc         : 수불번호팝업     
'*  7. Modified date(First) : 2000/04/18       
'*  8. Modified date(Last)  : 2005/08/05       
'*  9. Modifier (First)     : An Chang Hwan    
'* 10. Modifier (Last)      : Lee Seung Wook   
'* 11. Comment              :                  
'********************************************************************************************************-->
<HTML>
<HEAD>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<Script Language="VBS">
Option Explicit

Const BIZ_PGM_ID = "i1111pb1.asp"      

Dim C_DocumentNo
Dim C_Year
Dim C_DocumentDt
Dim C_MovType
Dim C_Plant
Dim C_DocumentText
Dim C_MovTypeNm
'**************** 수불유형명 추가 LSW 2005-08-05 *********

'*********************************************  1.3 변 수 선 언  ****************************************
'* 설명: Constant는 반드시 대문자 표기.                *
'********************************************************************************************************
<!-- #Include file="../../inc/lgvariables.inc" -->
Dim lgQueryFlag     
Dim lgDocumentNo
Dim lgYear
Dim lgFromDt
Dim lgToDt
Dim lgMovType
Dim lgTrnsType

Dim hlgDocumentNo   
Dim hlgYear
Dim hlgFromDt
Dim hlgToDt
Dim hlgMovType
Dim hlgTrnsType
Dim hlgPlantCd      
Dim hlgPlantNm  

Dim arrParam     
Dim arrReturn    
Dim CurrentDate

CurrentDate = "<%=GetSvrDate%>"

arrParam = window.dialogArguments
Set PopupParent = arrParam(0)

top.document.title = PopupParent.gActivePRAspName

'=======================================================================================================
' Function Name : LoadInfTB19029
'=======================================================================================================
Function LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "I","NOCOOKIE","PA") %>
End Function

'==========================================  2.2.1 SetDefaultVal()  =====================================
Sub SetDefaultVal()
 Dim strYear
 Dim strMonth
 Dim strDay
 Dim StartDate

 txtDocumentNo.value = arrParam(1)
 Call ExtractDateFrom(CurrentDate, PopupParent.gServerDateFormat,PopupParent.gServerDateType,strYear,strMonth,strDay)
 
 If arrParam(2) = strYear then
	txtYear.Text   = strYear
	StartDate      = UNIDateAdd("M", -1, CurrentDate, PopupParent.gServerDateFormat)
	txtFromDt.Text = UniConvDateAToB(StartDate, PopupParent.gServerDateFormat,PopupParent.gDateFormat)
	txtToDt.Text   = UniConvDateAToB(CurrentDate, PopupParent.gServerDateFormat,PopupParent.gDateFormat) 
 Else
	txtYear.Text = arrParam(2)
 End if
 
 txtTrnsType.Value	= arrParam(3)
 hlgPlantCd			= arrParam(4)
 
 Self.Returnvalue = Array("")
End Sub

'==========================================  2.2.2 InitSpreadSheet()  ===================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	ggoSpread.Source = vspdData
	ggoSpread.Spreadinit "V20050805", , PopupParent.gAllowDragDropSpread

	vspdData.ReDraw = False
	vspdData.MaxCols = C_DocumentText + 1
	vspdData.MaxRows = 0
	
	Call GetSpreadColumnPos("A")
	
	ggoSpread.SSSetEdit C_DocumentNo,		"수불번호", 16
	ggoSpread.SSSetEdit C_Year,				"년도", 8,2 
	ggoSpread.SSSetEdit C_DocumentDt,		"수불일자", 10,2
	ggoSpread.SSSetEdit C_MovType,			"수불유형", 10,2
	ggoSpread.SSSetEdit C_MovTypeNm,		"수불유형명", 20
	ggoSpread.SSSetEdit C_Plant,			"공장", 8,2
	ggoSpread.SSSetEdit C_DocumentText,		"비고", 27 

	Call ggoSpread.SSSetColHidden(vspdData.MaxCols, vspdData.MaxCols, True)
	ggoSpread.SSSetSplit2(1)
	ggoSpread.SpreadLockWithOddEvenRowColor()
	vspdData.ReDraw = True
	
End Sub

'========================================================================================
' Function Name : InitSpreadPosVariables
'======================================================================================== 
Sub InitSpreadPosVariables()
	C_DocumentNo	= 1
	C_Year			= 2
	C_DocumentDt	= 3
	C_MovType		= 4
	C_MovTypeNm		= 5
	C_Plant			= 6
	C_DocumentText	= 7
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'======================================================================================== 
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
	Case "A"
		ggoSpread.Source = vspdData 
		
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_DocumentNo   = iCurColumnPos(1)
		C_Year         = iCurColumnPos(2)
		C_DocumentDt   = iCurColumnPos(3)
		C_MovType      = iCurColumnPos(4)
		C_MovTypeNm	   = iCurColumnPos(5)
		C_Plant        = iCurColumnPos(6)
		C_DocumentText = iCurColumnPos(7)
	End Select

End Sub

'===========================================  2.3.1 OkClick()  ==========================================
Function OKClick()
 Dim intColCnt
 
 If vspdData.ActiveRow > 0 Then 
	Redim arrReturn(vspdData.MaxCols)
 
	vspdData.Row = vspdData.ActiveRow
    
	vspdData.Col = C_DocumentNo
	arrReturn(0) = vspdData.Text
	vspdData.Col = C_Year
	arrReturn(1) = vspdData.Text
	vspdData.Col = C_DocumentDt
	arrReturn(2) = vspdData.Text
	vspdData.Col = C_MovType
	arrReturn(3) = vspdData.Text
	vspdData.Col = C_Plant
	arrReturn(4) = vspdData.Text
	vspdData.Col = C_DocumentText
	arrReturn(5) = vspdData.Text  

	arrReturn(6) = ""

    arrReturn(7) = hlgPlantNm
    Self.Returnvalue = arrReturn  
 End If
 
 Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()
 Self.Close()
End Function

'=========================================  3.1.1 Form_Load()  ==========================================
Sub Form_Load()
 Call LoadInfTB19029
 Call ggoOper.LockField(Document, "N")                                         
 Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
 Call ggoOper.FormatDate(txtYear, PopupParent.gDateFormat, 3)                  
 Call SetDefaultVal()
 Call InitSpreadSheet()
 Call FncQuery()
 Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
End Sub

'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'=======================================================================================================
Sub txtYear_DblClick(Button) 
    If Button = 1 Then
        txtYear.Action = 7
        Call SetFocusToDocument("M")  
        txtYear.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_DblClick(Button)
'=======================================================================================================
Sub txtFromDt_DblClick(Button)
    If Button = 1 Then
        txtFromDt.Action = 7
        Call SetFocusToDocument("M")  
        txtFromDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtFromDt_Change()
'=======================================================================================================
Sub txtFromDt_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtToDt_DblClick(Button)
'=======================================================================================================
Sub txtToDt_DblClick(Button)
    If Button = 1 Then
        txtToDt.Action = 7
        Call SetFocusToDocument("M")  
        txtToDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtToDt_Change()
'=======================================================================================================
Sub txtToDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")
	gMouseClickStatus = "SPC"   
   
	Set gActiveSpdSheet = vspdData
   
	If vspdData.MaxRows = 0 Then Exit Sub
	
	If Row <= 0 Then
		ggoSpread.Source = vspdData 
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
	If vspdData.MaxRows = 0 Then Exit Sub

	If vspdData.MaxRows > 0 Then
		If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
			Call OKClick()
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
   ggoSpread.Source = vspdData
   Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 

'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   ggoSpread.Source = vspdData
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

'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then Exit Sub
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub

'========================================================================================
' Function Name : txtYear_KeyPress
'========================================================================================
Sub txtYear_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
		Call CancelClick()
	ElseIf KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

'========================================================================================
' Function Name : txtFromDt_KeyPress
'========================================================================================
Function txtFromDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
		Call CancelClick()
	ElseIf KeyAscii = 13 Then
		Call FncQuery()
	End If
End Function

'========================================================================================
' Function Name : txtToDt_KeyPress
'========================================================================================
Function txtToDt_KeyPress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 27 Then
		Call CancelClick()
	ElseIf KeyAscii = 13 Then
		Call FncQuery()
	End If
End Function

'========================================================================================
' Function Name : vspdData_KeyPress
'========================================================================================
Function vspdData_KeyPress(KeyAscii)
    On Error Resume Next
	If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function

'========================================================================================
' Function Name : vspdData_TopLeftChange
'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
	If OldLeft <> NewLeft Then Exit Sub
 	If vspdData.MaxRows < NewTop + VisibleRowCnt(vspdData, NewTop) And lgStrPrevKey <> "" Then
		DbQuery
	End if
End Sub
 

'********************************************  5.1 FncQuery()  *******************************************
Function FncQuery()

    If ValidDateCheck(txtFromDt, txtToDt) = False Then Exit Function

	ggoSpread.Source = vspdData
	ggoSpread.ClearSpreadData

	lgQueryFlag		= "1"
	lgDocumentNo	= Trim(txtDocumentNo.Value)
	lgYear			= txtYear.Text
	lgFromDt		= txtFromDt.Text
	lgToDt			= txtToDt.Text
	lgMovType		= Trim(txtMovType.value)
	lgTrnsType		= Trim(txtTrnsType.value)
	lgStrPrevKey	= ""
 
	If Not chkField(Document, "1") Then Exit Function
	
	If DbQuery() = False Then Exit Function
 
End Function

'********************************************  5.2 DbQuery()  *******************************************
Function DbQuery()

	Dim strVal
	Dim txtMaxRows

    Call LayerShowHide(1)  
	DbQuery = False 

	txtMaxRows = vspdData.MaxRows

	If lgStrPrevKey <> "" Then
		 strVal = BIZ_PGM_ID &	"?txtDocumentNo="	& hlgDocumentNo	& _
								"&txtYear="			& hlgYear		& _
								"&txtFromDt="		& hlgFromDt		& _
								"&txtToDt="			& hlgToDt		& _
								"&txtMovType="		& hlgMovType	& _
								"&txtTrnsType="		& hlgTrnsType	& _
								"&txtPlantCd="		& hlgPlantCd	& _
								"&lgStrPrevKey="	& lgStrPrevKey	& _
								"&txtMaxRows="		& txtMaxRows
	Else
		 strVal = BIZ_PGM_ID &	"?txtDocumentNo="	& lgDocumentNo	& _
								"&txtYear="			& lgYear		& _
								"&txtFromDt="		& lgFromDt		& _
								"&txtToDt="			& lgToDt		& _
								"&txtMovType="		& lgMovType		& _
								"&txtTrnsType="		& lgTrnsType	& _
								"&txtPlantCd="		& hlgPlantCd	& _
								"&lgStrPrevKey="	& lgStrPrevKey	& _
								"&txtMaxRows="		& txtMaxRows
	End if                                                      

	Call RunMyBizASP(MyBizASP, strVal)       
	
	DbQuery = True                                                        

End Function

Function DbQueryOk()     
	vspdData.Focus
End Function

</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" --> 

</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
	<TABLE CELLSPACING=0 CLASS="basicTB">
		<TR>
			<TD HEIGHT=40>
				<FIELDSET CLASS="CLSFLD">
					<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100%>
						<TR>
							<TD CLASS="TD5" NOWRAP>수불번호</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtDocumentNo" SIZE=20 MAXLENGTH=16 tag="11xxxU" ></TD>
							<TD CLASS="TD5" NOWRAP>년도</TD>
							<TD CLASS="TD6" NOWRAP>
							<script language =javascript src='./js/i1111pa1_I232158124_txtYear.js'></script>
							</TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>수불일자</TD>
							<TD CLASS="TD6" NOWRAP> <script language =javascript src='./js/i1111pa1_I255351883_txtFromDt.js'></script>
							 &nbsp;~&nbsp;
							<script language =javascript src='./js/i1111pa1_I964206886_txtToDt.js'></script>
							</TD>
							<TD CLASS="TD5" NOWRAP>수불유형</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtMovType" SIZE=5 MAXLENGTH=3  tag="11xxxU" ></TD>
						</TR>
						<TR>
							<TD CLASS="TD5" NOWRAP>수불구분</TD>
							<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtTrnsType" SIZE=10 MAXLENGTH=18   tag="14xxxU"  ></TD>
							<TD CLASS="TD5" NOWRAP></TD>
							<TD CLASS="TD6" NOWRAP></TD>
						</TR>
					</TABLE>
				</FIELDSET>
			</TD>
		</TR>
		<TR>
			<TD HEIGHT=100%>
				<script language =javascript src='./js/i1111pa1_I721083039_vspdData.js'></script>
			</TD>
		</TR>
		<TR HEIGHT=20>
			<TD WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG></TD>
						<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
												  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>     </TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>  
			</TD>
		</TR>
	</TABLE>
	<DIV ID="MousePT" NAME="MousePT">
		<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
	</DIV>
</BODY>
</HTML>

