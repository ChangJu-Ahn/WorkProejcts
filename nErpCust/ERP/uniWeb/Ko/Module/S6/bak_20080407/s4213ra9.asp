<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        : 																			*
'*  3. Program ID           : S4213RA9
'*  4. Program Name         : 통관참조(CONTAINER 배정)
'*  5. Program Desc         : 										*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2005/01/27																*
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     ::HJO
'* 10. Modifier (Last)      : 
'* 11. Modifier             : 
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/07 : 화면 design												*
'********************************************************************************************************
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit		

Const BIZ_PGM_ID 		= "s4213rb9.asp"                                         
Const C_MaxKey          = 16                                           

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgCookValue 
Dim gblnWinEvent

Dim arrReturn										
Dim arrParam	
Dim arrParent
Dim lgIsOpenPop

Dim iDBSYSDate
Dim EndDate, StartDate

arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
top.document.title = PopupParent.gActivePRAspName

iDBSYSDate = "<%=GetSvrDate%>"

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
'========================================================================================================
Function InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
	lgBlnFlgChgValue = False
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          
    lgSortKey        = 1
			
	gblnWinEvent = False
	ReDim arrReturn(0,0)
	Self.Returnvalue = arrReturn
End Function
'========================================================================================================
Sub SetDefaultVal()
		
	arrParam = arrParent(1)
		
	With frm1
		.txtCCNo.value   = arrParam(0)
	End With

End Sub
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "*", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "*", "NOCOOKIE", "QA") %>
End Sub
'========================================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("s4213ra9","S","A","V20030318",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
		'Call ggoSpread.SSSetColHidden(5,5,True)	
	Call SetSpreadLock       
			      
End Sub
'========================================================================================================
Sub SetSpreadLock()
	frm1.vspdData.OperationMode = 5
	ggoSpread.SpreadLockWithOddEvenRowColor
End Sub
'========================================================================================================
Function OKClick()
	
	Dim intColCnt, intRowCnt, intInsRow

	If frm1.vspdData.SelModeSelCount > 0 Then 

		intInsRow = 0

		Redim arrReturn(frm1.vspdData.SelModeSelCount - 1, frm1.vspdData.MaxCols - 1)

		For intRowCnt = 0 To frm1.vspdData.MaxRows - 1

			frm1.vspdData.Row = intRowCnt + 1

			If frm1.vspdData.SelModeSelected Then
				For intColCnt = 0 To frm1.vspdData.MaxCols - 2
					frm1.vspdData.Col = GetKeyPos("A",intColCnt + 1)
					arrReturn(intInsRow, intColCnt) = frm1.vspdData.Text
				Next

				intInsRow = intInsRow + 1

			End IF
		Next
	End if			
		
	Self.Returnvalue = arrReturn
	Self.Close()
End Function	
'========================================================================================================
Function CancelClick()
	Redim arrReturn(1,1)
	arrReturn(0,0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Sub Form_Load()
	Call LoadInfTB19029				                                           

	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
											
	Call ggoOper.LockField(Document, "N")						

	Call InitVariables														    
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call FncQuery()
End Sub

'========================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)

	If Row = 0 Then Exit Function

	If frm1.vspdData.MaxRows = 0 Then Exit Function

	If Row > 0 Then Call OKClick()

End Function

'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If   

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		If lgPageNo <> "" Then		                                                    
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If
End Sub
'==========================================================================================

'========================================================================================
Function FncQuery() 
    
    FncQuery = False                                                 
    
    Err.Clear                                                        

	
'	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function

	Call ggoOper.ClearField(Document, "2")	         						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

	Call InitVariables												

    If DbQuery = False Then Exit Function							

    FncQuery = True									
        
End Function
'********************************************************************************************************
Function DbQuery()
	Err.Clear															

	DbQuery = False														

	If LayerShowHide(1) = False Then
		Exit Function
	End If

	Dim strVal
		
	With frm1
		
	If lgIntFlgMode = PopupParent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
		strVal = strVal & "&txtCCNo=" & Trim(.txtHCCNO.value)			
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
		strVal = strVal & "&txtCCNo=" & Trim(.txtCCNo.value)				

	End If		
	End With

'--------- Developer Coding Part (End) ------------------------------------------------------------
    strVal = strVal & "&lgPageNo="       & lgPageNo                                 
	strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
	strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
	strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		        
		

	Call RunMyBizASP(MyBizASP, strVal)									

	DbQuery = True														
End Function
'========================================================================================
Function DbQueryOk()														

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	:	frm1.vspdData.SelModeSelected = True		
	Else
		'frm1.txtSoNo.focus
	End If

End Function
'===========================================================================

Function OpenSortPopup()
	
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

</SCRIPT>

<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>
<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
	<TABLE <%=LR_SPACE_TYPE_20%>>
		<TR>
			<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>통관관리번호</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCCNo" SIZE=20 MAXLENGTH=18 TAG="14XXXU"></TD>
						<TD CLASS=TD5 NOWRAP></TD>
						<TD CLASS=TD6 NOWRAP></TD>
					</TR>					
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=* valign=top>
				<TABLE WIDTH="100%" HEIGHT="100%">
					<TR>
						<TD HEIGHT="100%">
							<script language =javascript src='./js/s4213ra9_vaSpread1_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_01%>></TD>
		</TR>
		<TR HEIGHT="20">
			<TD WIDTH="100%">
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=70% NOWRAP>
							<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>
							<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>						
						<TD WIDTH=30% ALIGN=RIGHT>
							<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>
							<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG></TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> SCROLLING=NO noresize FRAMEBORDER=0  framespacing=0 TABINDEX="-1"></IFRAME></TD>
		</TR>
	</TABLE>
<INPUT TYPE=HIDDEN NAME="HCCNO" tag="24">

</FORM>	
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1"></iframe>
  </DIV>
</BODY>
</HTML>