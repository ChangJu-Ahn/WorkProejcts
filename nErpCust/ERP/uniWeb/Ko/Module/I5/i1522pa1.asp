<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 
'*  3. Program ID           : I1522PA1.ASP
'*  4. Program Name         : VMI 공장별 품목 팝업 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/01/08
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Lee Seung Wook
'* 10. Modifier (Last)      : Lee Seung Wook
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2003/01/08
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

Const BIZ_PGM_ID = "i1522pb1.asp"

Dim C_ItemCode    
Dim C_ItemName
Dim C_Spec
Dim C_Unit
Dim C_ItemAcct
Dim C_ItemGroupCd
Dim C_ProcurType
Dim C_LotFlg
Dim C_MajorSlCd
Dim C_IssuedSlCd
Dim C_ValidFlg
Dim C_RecvInspecFlg    
Dim C_TrackingFlg
Dim C_LotGenMthd    

Dim arrReturn
Dim arrParent
Dim arrParam1
Dim arrParam2
Dim PlantCd

arrParent		= window.dialogArguments

set PopupParent = arrParent(0)
arrParam1		= arrParent(1)
arrParam2		= arrParent(2)
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

'========================================================================================================
' Name : InitComboBox()	
'========================================================================================================
Sub InitComboBox()
	On Error Resume Next
    Err.Clear
    '------------------------------------------------------------
	' Setting Item Account Combo
	'------------------------------------------------------------
	Call CommonQueryRs(" MINOR_CD, MINOR_NM ", " B_MINOR ", " MAJOR_CD = " & FilterVar("P1001", "''", "S") & "  ", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6)
	Call SetCombo2(frm1.cboItemAccount,lgF0  ,lgF1  ,Chr(11))
End Sub
	
'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	PlantCd					= arrParam1
	frm1.txtItemCd.value	= arrParam2
End Sub 

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()

	ggoSpread.Source = frm1.vspdData
	ggoSpread.Spreadinit "V20030409"
	
	With  frm1.vspdData
		.ReDraw = false
	    .OperationMode = 3
	    .MaxCols = C_TrackingFlg+1											
	    .MaxRows = 0

		Call GetSpreadColumnPos("A")
		
	    ggoSpread.SSSetEdit  C_ItemCode,		"품목",			18, 0, -1, 18
	    ggoSpread.SSSetEdit  C_ItemName,		"품목명",		25, 0, -1, 40
	    ggoSpread.SSSetEdit  C_Spec,			"규격",			25, 0, -1,50
	    ggoSpread.SSSetEdit	 C_Unit,			"단위",			6, 0, -1, 3
	    ggoSpread.SSSetEdit  C_ItemAcct,		"품목계정",		10, 0, -1, 3
	    ggoSpread.SSSetEdit  C_ItemGroupCd,		"품목그룹",		10, 0, -1, 10
	    ggoSpread.SSSetEdit  C_ProcurType,		"조달구분",		10, 0, -1, 2
	    ggoSpread.SSSetEdit  C_LotFlg,			"LOT관리",		12, 0, -1, 1
	    ggoSpread.SSSetEdit  C_MajorSlCd,		"입고창고",		10, 0, -1, 7
	    ggoSpread.SSSetEdit  C_IssuedSlCd,		"출고창고",		10, 0, -1, 7
	    ggoSpread.SSSetEdit  C_ValidFlg,		"유효구분",		10, 0, -1, 1
	    ggoSpread.SSSetEdit	 C_RecvInspecFlg,	"수입검사구분",	14, 0, -1, 1
	    ggoSpread.SSSetEdit	 C_TrackingFlg,		"Tracking 구분",	14, 0, -1, 1
	    ggoSpread.SSSetEdit  C_LotGenMthd,		"LOT채번구분",	14, 0,-1,1
	    
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
	C_ItemCode			= 1										
	C_ItemName			= 2
	C_Spec				= 3
	C_Unit				= 4
	C_ItemAcct			= 5
	C_ItemGroupCd		= 6
	C_ProcurType		= 7
	C_LotFlg			= 8
	C_MajorSlCd			= 9
	C_IssuedSlCd		= 10
	C_ValidFlg			= 11
	C_RecvInspecFlg		= 12
	C_TrackingFlg		= 13
	C_LotGenMthd		= 14
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
	    C_Spec			= iCurColumnPos(3)
	    C_Unit			= iCurColumnPos(4)
	    C_ItemAcct		= iCurColumnPos(5)
	    C_ItemGroupCd	= iCurColumnPos(6)
	    C_ProcurType	= iCurColumnPos(7)
	    C_LotFlg		= iCurColumnPos(8)
	    C_MajorSlCd		= iCurColumnPos(9)
	    C_IssuedSlCd	= iCurColumnPos(10)
	    C_ValidFlg		= iCurColumnPos(11)
	    C_RecvInspecFlg	= iCurColumnPos(12)
		C_TrackingFlg	= iCurColumnPos(13)
		C_LotGenMthd	= iCurColumnPos(14)
	End Select
End Sub

'===========================================  2.3.1 OkClick()  ==========================================
'=	Name : OkClick()																					=
'========================================================================================================
Function OKClick()	
	Dim intColCnt
  
	With frm1.vspdData
		If .ActiveRow > 0 Then 
		
			Redim arrReturn(.MaxCols - 1)
	  
			.Row = .ActiveRow
	     
			For intColCnt = 0 To .MaxCols - 1
				.Col = intColCnt + 1
				arrReturn(intColCnt) = .Text
			Next
	    
			Self.Returnvalue = arrReturn
		End If
	End With
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
    Call InitComboBox()
    Call SetDefaultVal()
    Call InitSpreadSheet()
    
    If DbQuery = False Then Exit Sub
	
End Sub

'**************************  3.2 HTML Form Element & Object Event처리  **********************************
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
    
    strVal = BIZ_PGM_ID	&	"?PlantCd="			& Trim(PlantCd)						& _
							"&txtItemCd="		& Trim(frm1.txtItemCd.value)		& _
							"&txtItemNm="		& Trim(frm1.txtItemNm.value)		& _
							"&cboItemAccount="	& Trim(frm1.cboItemAccount.value)	& _
							"&txtSpec="			& Trim(frm1.txtSpec.value)			& _
							"&txtMaxRows="		& frm1.vspdData.MaxRows				& _
							"&lgStrPrevKeyIndex=" & lgStrPrevKeyIndex
    
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
				<TD CLASS="TD5" NOWRAP>품목</TD>
				<TD CLASS="TD6" COLSPAN = 3 NOWRAP><INPUT TYPE="Text" Name="txtItemCd" SIZE=20 MAXLENGTH=18 tag="11XXXU" ALT="품목" TABINDEX="-1">&nbsp;<INPUT TYPE=TEXT NAME="txtItemNm" SIZE=40 MAXLENGTH=40 tag="11XXXU" ALT="품목명" TABINDEX="-1"></TD>
				
			</TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>품목계정</TD>
				<TD CLASS="TD6" NOWRAP><SELECT NAME="cboItemAccount" ALT="품목계정" STYLE="Width: 160px;" tag="11" TABINDEX="-1"><OPTION VALUE=""></SELECT></TD>
				<TD CLASS="TD5" NOWRAP>규격</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtSpec" SIZE=50 MAXLENGTH=50 tag="11XXXU" ALT="규격" TABINDEX="-1"></TD>
			</TR>
		</TABLE>
		</FIELDSET>
		</TD>
	</TR>
	<TR><TD HEIGHT=100%>
		<script language =javascript src='./js/i1522pa1_OBJECT1_vspdData.js'></script>
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


