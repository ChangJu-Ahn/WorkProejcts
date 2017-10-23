<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : I1521PA1.ASP
'*  4. Program Name         : VMI 수불번호 팝업 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/10
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Ahn Jung Je
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2003/01/10
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

Const BIZ_PGM_ID = "i1521pb1.asp"

Dim C_ItemDocumentNo    
Dim C_DocumentYear
Dim C_DocumentDt
Dim C_PlantCd
Dim C_PlantNm
Dim C_BpCd
Dim C_BpNm
Dim C_DocumentText

Dim arrReturn
Dim arrParent
Dim arrParam
Dim arrDocumentNo
Dim arrDocumentYear
Dim TrnsType
Dim PlantCd

arrParent		= window.dialogArguments

set PopupParent = arrParent(0)
Dim arrTemp
arrTemp = arrParent(1)

arrDocumentNo		= arrTemp(0)
arrDocumentYear		= arrTemp(1)
TrnsType			= arrTemp(2)
PlantCd				= arrTemp(3)

top.document.title = PopupParent.gActivePRAspName

Dim lgOldRow
Dim gblnWinEvent
Dim strReturn

Dim IsOpenPop          

'#########################################################################################################
'												2. Function부 
'######################################################################################################### 
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
	Dim strYear
	Dim strMonth
	Dim strDay
	Dim StartDate

	Call ExtractDateFrom("<%=GetSvrDate%>",PopupParent.gServerDateFormat,PopupParent.gServerDateType,strYear,strMonth,strDay)
	frm1.txtItemDocumentNo.value = arrDocumentNo
	If arrDocumentYear = "" then
		frm1.txtDocumentYear.year = strYear
	Else
		frm1.txtDocumentYear.year = arrDocumentYear
	End if
	frm1.txtTrnsType.value	= TrnsType
	PlantCd					= PlantCd
	
	StartDate = UNIDateAdd("M", -1, "<%=GetSvrDate%>", PopupParent.gServerDateFormat)
	frm1.txtDocumentDt1.Text	= UniConvDateAToB(StartDate, PopupParent.gServerDateFormat,PopupParent.gDateFormat)
	frm1.txtDocumentDt2.Text	= UniConvDateAToB("<%=GetSvrDate%>", PopupParent.gServerDateFormat,PopupParent.gDateFormat)
	
	Self.Returnvalue = Array("")
    
End Sub 

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()


	Call InitSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030425", , PopupParent.gAllowDragDropSpread

	With  frm1.vspdData
		.ReDraw = false
	    .OperationMode = 3
	    .MaxCols = C_DocumentText+1											
	    .MaxRows = 0

		Call GetSpreadColumnPos("A")
		
	    ggoSpread.SSSetEdit  C_ItemDocumentNo,	"수불번호",	15, 0, -1, 18
	    ggoSpread.SSSetEdit  C_DocumentYear,	"년도",		6, 0, -1, 4
	    ggoSpread.SSSetEdit  C_DocumentDt,		"수불일자",	10, 0, -1,10
	    ggoSpread.SSSetEdit	 C_PlantCd,			"공장",		6, 0, -1, 4
	    ggoSpread.SSSetEdit  C_PlantNm,			"공장명",	20, 0, -1, 30
	    ggoSpread.SSSetEdit  C_BpCd,			"공급처",	10, 0, -1, 7
	    ggoSpread.SSSetEdit  C_BpNm,			"공급처명",	20, 0, -1, 30
	    ggoSpread.SSSetEdit  C_DocumentText,	"비고",		30, 0, -1, 50
	    
	    Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
	    		
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SSSetSplit(2)
		
		.ReDraw = true
    End With
End Sub

'========================================================================================================
' Name : InitSpreadPosVariables()	
'========================================================================================================
Sub InitSpreadPosVariables()
	C_ItemDocumentNo			= 1											
	C_DocumentYear				= 2
	C_DocumentDt				= 3
	C_PlantCd					= 4
	C_PlantNm					= 5
	C_BpCd						= 6
	C_BpNm						= 7
	C_DocumentText				= 8
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
		
		C_ItemDocumentNo		= iCurColumnPos(1)
    	C_DocumentYear			= iCurColumnPos(2)
	    C_DocumentDt			= iCurColumnPos(3)
	    C_PlantCd				= iCurColumnPos(4)
	    C_PlantNm				= iCurColumnPos(5)
	    C_BpCd					= iCurColumnPos(6)
	    C_BpNm					= iCurColumnPos(7)
	    C_DocumentText			= iCurColumnPos(8)
	End Select
End Sub

'===========================================  2.3.1 OKClick()  ==========================================
Function OKClick()
	Dim intColCnt
	
	With frm1.vspdData  
		If .ActiveRow > 0 Then 
			Redim arrReturn(.MaxCols - 1)
  
			.Row = .ActiveRow
			.Col = C_ItemDocumentNo
			arrReturn(0) = .Text
			.Col = C_DocumentYear
			arrReturn(1) = .Text
			.Col =	C_DocumentDt
			arrReturn(2) = .Text
			.Col =	C_PlantCd
			arrReturn(3) = .Text
			.Col =	C_PlantNm
			arrReturn(4) = .Text
			.Col =	C_BpCd
			arrReturn(5) = .Text
			.Col =	C_BpNm
			arrReturn(6) = .Text
			.Col =	C_DocumentText
			arrReturn(7) = .Text 
	
			Self.Returnvalue = arrReturn
		End If
	End With
	Self.Close()
End Function

'=========================================  2.3.2 CancelClick()  ========================================
Function CancelClick()		
	Self.Close()
End Function

'#########################################################################################################
'												3. Event부 
'##########################################################################################################
'******************************************  3.1 Window 처리  *********************************************
'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")                                       
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtDocumentYear, PopupParent.gDateFormat, 3)
    
    Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")    

    Call InitSpreadSheet
    Call InitVariables                                                
    Call SetDefaultVal()
    Call InitSpreadSheet()
    
    If DbQuery = False Then Exit Sub
	
End Sub

'=======================================================================================================
'   Event Name : txtDocumentYear_DblClick(Button)
'=======================================================================================================
Sub txtDocumentYear_DblClick(Button) 
    If Button = 1 Then
        frm1.txtDocumentYear.Action = 7
        Call SetFocusToDocument("M")  
        frm1.txtDocumentYear.Focus
    End If
End Sub

Sub txtDocumentYear_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDt1_DblClick(Button)
'=======================================================================================================
Sub txtDocumentDt1_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentDt1.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentDt1.Focus
    End If
End Sub

Sub txtDocumentDt1_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDt2_Change()
'=======================================================================================================
Sub txtDocumentDt1_Change()
    lgBlnFlgChgValue = True
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDt1_DblClick(Button)
'=======================================================================================================
Sub txtDocumentDt2_DblClick(Button)
    If Button = 1 Then
        frm1.txtDocumentDt2.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtDocumentDt2.Focus
    End If
End Sub

Sub txtDocumentDt2_KeyPress(KeyAscii)
	If KeyAscii = 13 Then
		Call FncQuery()
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDocumentDt2_Change()
'=======================================================================================================
Sub txtDocumentDt2_Change()
    lgBlnFlgChgValue = True
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

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 

    FncQuery = False                                                      
    Err.Clear                                                             
    '-----------------------
    'Erase contents area
    '-----------------------
 	ggoSpread.Source = frm1.vspdData
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

'*******************************  5.2 Fnc함수명에서 호출되는 개발 Function  *******************************
'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    
    Call LayerShowHide(1)
    
    DbQuery = False
    
    Err.Clear                                                            
    Dim strVal
    
    strVal = BIZ_PGM_ID	&	"?PlantCd="				& Trim(PlantCd)							& _
							"&txtItemDocumentNo="	& Trim(frm1.txtItemDocumentNo.value)	& _
							"&txtDocumentYear="		& Trim(frm1.txtDocumentYear.year)		& _
							"&txtDocumentDt1="		& Trim(frm1.txtDocumentDt1.text)		& _
							"&txtDocumentDt2="		& Trim(frm1.txtDocumentDt2.text)		& _
							"&txtTrnsType="			& Trim(frm1.txtTrnsType.value)			& _
							"&txtMaxRows="			& frm1.vspdData.MaxRows					& _
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
				<TD CLASS="TD5" NOWRAP>수불번호</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtItemDocumentNo" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="수불번호" TABINDEX="-1"></TD>
				<TD CLASS="TD5" NOWRAP>년도</TD>
				<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/i1521pa1_OBJECT2_txtDocumentYear.js'></script></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>수불일자</TD>
				<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/i1521pa1_OBJECT3_txtDocumentDt1.js'></script>
										&nbsp;~&nbsp;
									   <script language =javascript src='./js/i1521pa1_OBJECT4_txtDocumentDt2.js'></script></TD>
				<TD CLASS="TD5" NOWRAP>수불구분</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtTrnsType" SIZE=6 MAXLENGTH=2 tag="14XXXU" ALT="수불구분" TABINDEX="-1"></TD>
			</TR>
		</TABLE>
		</FIELDSET>
		</TD>
	</TR>
	<TR><TD HEIGHT=100%>
		<script language =javascript src='./js/i1521pa1_OBJECT1_vspdData.js'></script>
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


