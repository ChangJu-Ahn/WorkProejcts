<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m1311ra2.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : Open Po Ref Popup ASP														*
'*  6. Comproxy List        : 																			*
'*  7. Modified date(First) : 2002/07/22																*
'*  8. Modified date(Last)  : 2003/06/02					             								*
'*                            
'*  9. Modifier (First)     : Oh Chang Won 																*
'* 10. Modifier (Last)      : Kim Jin Ha		                											*	
'*                            
'* 11. Comment              :																			*
'* 12. Common Coding Guide  :																			*
'* 13. History              :																			*
'********************************************************************************************************
!-->
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit					<% '☜: indicates that All variables must be declared in advance %>

Const BIZ_PGM_ID 		= "m1311rb2.asp"                              '☆: Biz Logic ASP Name
Const C_MaxKey          = 13                                          '☆: key count of SpreadSheet
'================================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'================================================================================================================================
Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 

Dim lgIsOpenPop
Dim C_MaxSelList

Dim arrReturn					<% '--- Return Parameter Group %>
Dim arrParent
Dim arrParam
Dim EndDate, StartDate, iDBSYSDate

'================================================================================================================================	
arrParent = window.dialogArguments
Set PopupParent = arrParent(0)
arrParam= arrParent(1)
top.document.title = PopupParent.gActivePRAspName
'================================================================================================================================
iDBSYSDate = "<%=GetSvrDate%>"
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)
'================================================================================================================================
Function InitVariables()
	lgStrPrevKey     = ""								   'initializes Previous Key
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           'Indicates that no value changed
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          'Indicates that current mode is Create mode
    lgSortKey        = 1   
        
    lgIntGrpCount = 0										<%'⊙: Initializes Group View Size%>

    gblnWinEvent = False
    Redim arrReturn(0,0)        
    Self.Returnvalue = arrReturn     
End Function
'================================================================================================================================
Sub SetDefaultVal()
	frm1.txtPlantCd.value 	= arrParam(1)
	frm1.txtItemCd.value 	= arrParam(2)
		
	if Trim(arrParam(3)) <> "" then
		frm1.txtSpplCd.Value	= arrParam(3)
	end if
End Sub
'================================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "RA") %>
End Sub
'================================================================================================================================
Sub InitSpreadSheet()
	
    Call SetZAdoSpreadSheet("M1311RA2","S","A","V20030526",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock("A")
	frm1.vspdData.OperationMode = 5 
End Sub
'================================================================================================================================
Sub SetSpreadLock(ByVal pOpt)
    IF pOpt = "A" Then
		ggoSpread.Source = frm1.vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
	End If
End Sub	
'================================================================================================================================
Function OKClick()
		
	Dim intColCnt, intRowCnt, intInsRow
	With frm1
			
		If .vspdData.SelModeSelCount > 0 Then 
			intInsRow = 0

			Redim arrReturn(.vspdData.SelModeSelCount+1, .vspdData.MaxCols - 1)
						
			For intRowCnt = 0 To .vspdData.MaxRows - 1
				.vspdData.Row = intRowCnt + 1

				If .vspdData.SelModeSelected Then
					For intColCnt = 0 To .vspdData.MaxCols - 1
						.vspdData.Col = GetKeyPos("A",intColCnt+1)
						frm1.vspdData.Col = GetKeyPos("A",intColCnt+1)
						arrReturn(intInsRow, intColCnt) = .vspdData.Text
					Next

					intInsRow = intInsRow + 1

				End If
			Next
		End if			
	End with
			
	Self.Returnvalue = arrReturn
	Self.Close()
			
End Function
'================================================================================================================================
Function CancelClick()
	Self.Close()
End Function
'================================================================================================================================
Function MousePointer(pstr1)
    Select case UCase(pstr1)
        case "PON"
	  	  window.document.search.style.cursor = "wait"
        case "POFF"
	  	  window.document.search.style.cursor = ""
    End Select
End Function
'================================================================================================================================
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)=UCase(PopupParent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장"	
	arrParam(1) = "B_Plant"				
	
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
    arrField(0) = "Plant_CD"	
    arrField(1) = "Plant_NM"	
    
    arrHeader(0) = "공장"		
    arrHeader(1) = "공장명"		

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPlantCd.Value= arrRet(0)		
		frm1.txtPlantNm.Value= arrRet(1)
		frm1.txtPlantCd.focus
		Set gActiveElement = document.activeElement
	End If	
	
End Function
'================================================================================================================================
Function OpenSppl()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or frm1.txtSpplCd.ClassName= PopupParent.UCN_PROTECTED  Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "외주처"						<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"					<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtSpplCd.value)		<%' Code Condition%>
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "		<%' Where Condition%>
	arrParam(5) = "외주처"						<%' TextBox 명칭 %>

	arrField(0) = "BP_CD"							<%' Field명(0)%>
	arrField(1) = "BP_NM"							<%' Field명(1)%>

	arrHeader(0) = "외주처"						<%' Header명(0)%>
	arrHeader(1) = "외주처명"					<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtSpplCd.Value  = arrRet(0)		
		frm1.txtSpplNm.Value  = arrRet(1)
		frm1.txtSpplCd.focus
		Set gActiveElement = document.activeElement
	End If
End Function
'================================================================================================================================
Sub Form_Load()
    ReDim lgPopUpR(C_MaxSelList - 1,1)
    
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    Call InitVariables											  '⊙: Initializes local global variables
    Call SetDefaultVal	
    Call AppendNumberPlace("6","5","4")
	Call InitSpreadSheet
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	
	if Trim(frm1.txtSpplCd.value)<>""  then
	    Call FncQuery()
	end if    
End Sub
'================================================================================================================================
Function OpenSortPopup()

	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True
	
    arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"), gMethodText),"dialogWidth=" & PopupParent.GROUPW_WIDTH & "px; dialogHeight=" & PopupParent.GROUPW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")
	lgIsOpenPop = False

	If arrRet(0) = "X" Then
		If Err.Number <> 0 Then
			Err.Clear 
		End If
		Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A", arrRet(0), arrRet(1))
       Call InitVariables
       Call InitSpreadSheet
   End If
End Function

'================================================================================================================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Row = 0 Or frm1.vspdData.MaxRows = 0 Then 
		Exit Function
	End If
	    
	With frm1.vspdData 
		If .MaxRows > 0 Then
			If .ActiveRow = Row Or .ActiveRow > 0 Then
				Call OKClick
			End If
		End If
	End With
End Function
'================================================================================================================================
Function vspdData_KeyPress(KeyAscii)
	On Error Resume Next
	If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then
		Call OKClick()
	ElseIf KeyAscii = 27 Then
		Call CancelClick()
	End If
End Function
'================================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If		

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If		 
End Sub
'================================================================================================================================
Function FncQuery() 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
	If Not chkField(Document, "1") Then								'⊙: This function check indispensable field
       Exit Function
    End If
    
    With frm1
        If CompareDateByFormat(.txtFrDt.text,.txtToDt.text,.txtFrDt.Alt,.txtToDt.Alt, _
                   "970025",.txtFrDt.UserDefinedFormat,PopupParent.gComDateType,False) = False And Trim(.txtFrDt.text) <> "" And Trim(.txtToDt.text) <> "" Then
			Call DisplayMsgBox("17a003","X","적용유효일","X")			
			Exit Function
		End if   
	End with
	
	Call ggoOper.ClearField(Document, "2")	  
	ggoSpread.Source = frm1.vspdData	
    ggoSpread.ClearSpreadData       	
    
    Call InitVariables 														'⊙: Initializes local global variables
    
    If DbQuery = False Then Exit Function									
	FncQuery = True		
    
End Function
'================================================================================================================================
Function DbQuery() 

	Err.Clear														'☜: Protect system from crashing
	DbQuery = False													'⊙: Processing is NG
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					'☜: 비지니스 처리 ASP의 상태	
		    strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey 
		    strVal = strVal & "&txtPlantCd=" & .hdnPlantCd.value
		    strVal = strVal & "&txtItemCd=" & .hdnItemCd.value
		    strVal = strVal & "&txtSpplCd=" & .hdnSpplCd.Value
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
		    strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		    strVal = strVal & "&txtPlantCd=" & Trim(.txtPlantCd.value)
		    strVal = strVal & "&txtItemCd=" & Trim(.txtItemCd.value)
		    strVal = strVal & "&txtSpplcd=" & Trim(.txtSpplCd.value)
		End If				
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						'☜: Next key tag 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
		
		Call RunMyBizASP(MyBizASP, strVal)		    						'☜: 비지니스 ASP 를 가동 
        
    End With
    
    DbQuery = True    

End Function
'================================================================================================================================
Function DbQueryOk()	    												'☆: 조회 성공후 실행로직 
	lgIntFlgMode = PopupParent.OPMD_UMODE
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.Row = 1	
		frm1.vspdData.SelModeSelected = True		
	else
		frm1.vspdData.Focus
	End If
End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY SCROLL=NO TABINDEX="-1">
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
						<TD CLASS="TD5" NOWRAP>외주처</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="외주처" NAME="txtSpplCd"  SIZE=10 MAXLENGTH=10 tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORG2Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSppl()">
											   <INPUT TYPE=TEXT ALT="외주처" NAME="txtSpplNm" SIZE=20 tag="14x"></TD>
						<TD CLASS="TD5" NOWRAP></TD>
						<TD CLASS="TD6" NOWRAP></TD>
					</TR>
					<TR>
						<TD CLASS="TD5" NOWRAP>공장</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="공장" NAME="txtPlantCd" SIZE=10 MAXLENGTH=4 tag="14XXXU">&nbsp;&nbsp;&nbsp;&nbsp;
										<INPUT TYPE=TEXT ALT="공장" NAME="txtPlantNm" SIZE=20 tag="14x"></TD>
						<TD CLASS="TD5" NOWRAP>모품목</TD>
						<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="모품목" NAME="txtItemCd"  SIZE=10 MAXLENGTH=18 tag="14XXXU">&nbsp;&nbsp;&nbsp;&nbsp;
										<INPUT TYPE=TEXT ALT="모품목" NAME="txtItemNm" SIZE=20 tag="14x"></TD>
					</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_60%>>
				<TR>
					<TD CLASS="TD5" NOWRAP>적용유효일</TD>
					<TD CLASS="TD6" NOWRAP>
						<table cellspacing=0 cellpadding=0>
							<tr>
								<td>
									<script language =javascript src='./js/m1311ra2_fpDateTime1_txtFrDt.js'></script>
								</td>
								<td>~</td>
								<td>
									<script language =javascript src='./js/m1311ra2_fpDateTime2_txtToDt.js'></script>
								</td>
							<tr>
						</table>
					</TD>
					<TD CLASS="TD6" NOWRAP></TD>
					<TD CLASS="TD6" NOWRAP></TD>
				</TR>
				<TR>
					<TD HEIGHT=100% WIDTH=100% COLSPAN=4>
						<script language =javascript src='./js/m1311ra2_vspdData_vspdData.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </tr>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD >&nbsp;&nbsp;<IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" ></IMG>&nbsp;
						                 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG></TD>
					<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/ok_d.gif"       Style="CURSOR: hand" ALT="OK"     NAME="Ok"     OnClick="OkClick()"         onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"    ></IMG>&nbsp;
                                         <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="hdnPlantCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnPlantNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnItemCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnItemNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSpplCd" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnSpplNm" tag="14">
<INPUT TYPE=HIDDEN NAME="hdnBomNo" tag="14">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
