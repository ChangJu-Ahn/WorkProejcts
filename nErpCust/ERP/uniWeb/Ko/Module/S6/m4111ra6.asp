<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : m4111ra6.asp																*
'*  4. Program Name         : 외주출고참조(통관등록에서)												*
'*  5. Program Desc         : 외주출고참조(통관등록에서)												*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2002/04/11																*
'*  8. Modified date(Last)  : 2002/07/10																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son Bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/04/11 : 화면 design												*
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

Const BIZ_PGM_ID 		= "m4111rb6.asp"  
Const C_MaxKey          = 8                                           

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
Dim arrParam
Dim strReturn                                               '☜: Return Parameter Group
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
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    gblnWinEvent = False

	ReDim strReturn(0)
	strReturn(0) = ""
	gblnWinEvent = False
	Self.Returnvalue = strReturn       
End Function
'=======================================================================================================
Sub SetDefaultVal()		
	frm1.txtFromDt.text = StartDate
	frm1.txtToDt.text = EndDate 
End Sub
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S", "NOCOOKIE", "RA") %>
End Sub
'========================================================================================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("m4111RA6","S","A","V20021202",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
	Call SetSpreadLock       
End Sub
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor
End Sub	
'========================================================================================================
Function OKClick()
	Dim intColCnt
	If frm1.vspdData.ActiveRow > 0 Then	
		
		Redim strReturn(frm1.vspdData.MaxCols - 1)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
		For intColCnt = 0 To frm1.vspdData.MaxCols - 1
			frm1.vspdData.Col = GetKeyPos("A",intColCnt + 1)
			strReturn(intColCnt) = frm1.vspdData.Text
		Next	
					
	End If
		
	Self.Returnvalue = strReturn
			
	Self.Close()
	
End Function
'========================================================================================================
Function CancelClick()
	Redim strReturn(0)
	strReturn(0) = ""
	Self.Returnvalue = strReturn
	Self.Close()
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenBizPartner()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "수입자"							<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtApplicant.value)				<%' Code Condition%>
	arrParam(3) = ""									<%' Name Cindition%>
	arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				<%' Where Condition%>
	arrParam(5) = "수입자"							<%' TextBox 명칭 %>

	arrField(0) = "BP_CD"								<%' Field명(0)%>
	arrField(1) = "BP_NM"								<%' Field명(1)%>

	arrHeader(0) = "수입자"							<%' Header명(0)%>
	arrHeader(1) = "수입자명"						<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBizPartner(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenPurGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "구매그룹"							<%' 팝업 명칭 %>
	arrParam(1) = "B_PUR_GRP"								<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtPurGroup.value)					<%' Code Condition%>
	arrParam(3) = ""										<%' Name Cindition%>
	arrParam(4) = ""										<%' Where Condition%>
	arrParam(5) = "구매그룹"								<%' TextBox 명칭 %>

	arrField(0) = "PUR_GRP"									<%' Field명(0)%>
	arrField(1) = "PUR_GRP_NM"								<%' Field명(1)%>

	arrHeader(0) = "구매그룹"							<%' Header명(0)%>
	arrHeader(1) = "구매그룹명"							<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPurGroup(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "공장"								<%' 팝업 명칭 %>
	arrParam(1) = "B_PLANT"									<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtPlant.value)						<%' Code Condition%>
	arrParam(3) = ""										<%' Name Cindition%>
	arrParam(4) = ""										<%' Where Condition%>
	arrParam(5) = "공장"								<%' TextBox 명칭 %>

	arrField(0) = "PLANT_CD"								<%' Field명(0)%>
	arrField(1) = "PLANT_NM"								<%' Field명(1)%>

	arrHeader(0) = "공장"								<%' Header명(0)%>
	arrHeader(1) = "공장명"								<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPlant(arrRet)
	End If
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetBizPartner(arrRet)
	frm1.txtApplicant.Value = arrRet(0)
	frm1.txtApplicantNm.Value = arrRet(1)
	frm1.txtApplicant.focus
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetPurGroup(arrRet)
	frm1.txtPurGroup.Value = arrRet(0)
	frm1.txtPurGroupNm.Value = arrRet(1)
	frm1.txtPurGroup.focus
End Function
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function SetPlant(arrRet)
	frm1.txtPlant.Value = arrRet(0)
	frm1.txtPlantNm.Value = arrRet(1)
	frm1.txtPlant.focus
End Function
'========================================================================================================
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

'********************************************************************************************************
Sub btnApplicantOnClick()
	Call OpenBizPartner()
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
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제 
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
'========================================================================================================
Sub txtFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtFromDt.Action = 7	
		Call SetFocusToDocument("P")
        frm1.txtFromDt.Focus	
	End If
End Sub

Sub txtToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("P")
        frm1.txtToDt.Focus
	End If
End Sub
'=======================================================================================================
Sub txtFromDt_Keypress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

Sub txtToDt_Keypress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub
'*********************************************************************************************************
Function FncQuery() 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               
		
	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function	
						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

    Call InitVariables 														

	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'========================================================================================================
Function DbQuery() 

	Err.Clear														
	DbQuery = False													
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
			strVal = strVal & "&txtApplicant=" & Trim(frm1.txtHApplicant.value)	
			strVal = strVal & "&txtPurGroup=" & Trim(frm1.txtHPurGroup.value)
			strVal = strVal & "&txtFromDt=" & Trim(frm1.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(frm1.txtHToDt.value)
			strVal = strVal & "&txtPlant=" & Trim(frm1.txtHPlant.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001				
			strVal = strVal & "&txtApplicant=" & Trim(frm1.txtApplicant.value)	
			strVal = strVal & "&txtPurGroup=" & Trim(frm1.txtPurGroup.value)
			strVal = strVal & "&txtFromDt=" & Trim(frm1.txtFromDt.text)
			strVal = strVal & "&txtToDt=" & Trim(frm1.txtToDt.text)
			strVal = strVal & "&txtPlant=" & Trim(frm1.txtPlant.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		End If
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

		Call RunMyBizASP(MyBizASP, strVal)		    						
        
    
    
    DbQuery = True    

End Function

'=========================================================================================================
Function DbQueryOk()	    												

	lgIntFlgMode = PopupParent.OPMD_UMODE
	
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	Else
		frm1.txtApplicant.focus
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
				<FIELDSET CLASS="CLSFLD">
					<TABLE <%=LR_SPACE_TYPE_40%>>
						<TR>
							<TD CLASS=TD5>수입자</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="수입자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK ="vbscript:btnApplicantOnClick()">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>구매그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPurGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="구매그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGroup" align=top TYPE="BUTTON" Onclick="vbscript:OpenPurGroup">&nbsp;<INPUT TYPE=TEXT NAME="txtPurGroupNm" SIZE=20 TAG="14"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>발주일</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/m4111ra6_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
								<script language =javascript src='./js/m4111ra6_fpDateTime2_txtToDt.js'></script>
							</TD>	
							<TD CLASS=TD5 NOWRAP>공장</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtPlant" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlant" align=top TYPE="BUTTON" Onclick="vbscript:OpenPlant">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=20 TAG="14"></TD>
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
						<TD HEIGHT="100%" NOWRAP>
							<script language =javascript src='./js/m4111ra6_vaSpread_vspdData.js'></script>
						</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD <%=HEIGHT_TYPE_01%>></TD>
		</TR>
		<TR>
			<TD HEIGHT=30>
				<TABLE <%=LR_SPACE_TYPE_30%>>
					<TR>
						<TD WIDTH=10>&nbsp;</TD>
						<TD WIDTH=70% NOWRAP>
							<IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"></IMG>
							<IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>						
						<TD WIDTH=30% ALIGN=RIGHT>
							<IMG SRC="../../../CShared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"></IMG>
							<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>
						</TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHApplicant" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPurGroup" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPlant" TAG="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
