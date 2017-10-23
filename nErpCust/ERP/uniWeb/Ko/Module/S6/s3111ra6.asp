<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업관리 
'*  2. Function Name        : 
'*  3. Program ID           : s3111ra6
'*  4. Program Name         : 수주참조(통관등록에서)
'*  5. Program Desc         : 수주참조(통관등록에서)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/02
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Choinkuk		
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
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
<!-- #Include file="../../inc/lgvariables.inc" -->	
Const BIZ_PGM_ID 		= "s3111rb6.asp"   
Const C_MaxKey          = 12                                           
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 
Dim lgIsOpenPop                                             '☜: Popup status                          
														
Dim arrReturn												'☜: Return Parameter Group
Dim arrParam
Dim arrParent
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
    Redim arrReturn(0)        
    Self.Returnvalue = arrReturn(0)     
End Function

'********************************************************************************************************
Sub SetDefaultVal()
	frm1.txtFromDt.text = StartDate
	frm1.txtToDt.text = EndDate
End Sub
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub
'========================================================================================================
Sub InitSpreadSheet()
	
	Call SetZAdoSpreadSheet("S3111RA6","S","A","V20030318",PopupParent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
	Call SetSpreadLock       
	    
End Sub
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor
End Sub	
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Function OKClick()

	Dim intColCnt
		
	If frm1.vspdData.ActiveRow > 0 Then	
		
		Redim arrReturn(frm1.vspdData.MaxCols - 1)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
		For intColCnt = 0 To frm1.vspdData.MaxCols - 1
			frm1.vspdData.Col = GetKeyPos("A",intColCnt + 1)		
			arrReturn(intColCnt) = frm1.vspdData.Text
		Next	
					
	End If
		
	Self.Returnvalue = arrReturn(0)
	Self.Close()
	
End Function
'========================================================================================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn(0)
	Self.Close()
End Function
'========================================================================================================
Function OpenConSItemDC(Byval iWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If gblnWinEvent = True Then Exit Function
	gblnWinEvent = True
	
	With frm1
		Select Case iWhere
			' 수입자 
			Case 0
				iArrParam(1) = "B_BIZ_PARTNER"									' TABLE 명칭 
				iArrParam(2) = Trim(.txtApplicant.value)						' Code Condition
				iArrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"         					' Where Condition   " C:매출처 "
				iArrParam(5) = .txtApplicant.alt								' TextBox 명칭 
					
				iArrField(0) = "BP_CD"											' Field명(0)
				iArrField(1) = "BP_NM"											' Field명(1)
		    
				iArrHeader(0) = .txtApplicant.alt								' Header명(0)
				iArrHeader(1) = .txtApplicantNm.alt								' Header명(1)
				
				.txtApplicant.focus
			' 영업그룹 
			Case 1
				iArrParam(1) = "B_SALES_GRP"
				iArrParam(2) = Trim(.txtSalesGroup.value)
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
				iArrParam(5) = .txtSalesGroup.alt
			
				iArrField(0) = "SALES_GRP"
				iArrField(1) = "SALES_GRP_NM"
			    
				iArrHeader(0) = .txtSalesGroup.alt
				iArrHeader(1) = .txtSalesGroupNm.Alt
				
				.txtSalesGroup.focus
			' 수주형태	
			Case 2
				iArrParam(1) = "S_SO_TYPE_CONFIG"
				iArrParam(2) = Trim(frm1.txtSOType.Value)
		
				iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "
				iArrParam(5) = .txtSOType.alt
				
			    iArrField(0) = "SO_TYPE"
			    iArrField(1) = "SO_TYPE_NM"
			    
			    iArrHeader(0) = .txtSOType.Alt
			    iArrHeader(1) = .txtSoTypeNm.Alt
			    
			    .txtSOType.focus
		End Select
	End With
	
	iArrParam(0) = iArrParam(5)											' 팝업 명칭	

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If iArrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(iArrRet, iWhere)
	End If	
	
End Function
'-------------------------------------------------------------------------------------------------------
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtApplicant.Value = arrRet(0)
			.txtApplicantNm.Value = arrRet(1) 	
			.txtApplicant.focus		   
		Case 1
			.txtSalesGroup.Value = arrRet(0)
			.txtSalesGroupNm.Value = arrRet(1)
			.txtSalesGroup.focus
		Case 2
			.txtSOType.Value = arrRet(0)
			.txtSOTypeNm.Value = arrRet(1)			 
			.txtSOType.focus
		End Select
	End With
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
    
    With frm1
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
			strVal = strVal & "&txtApplicant=" & Trim(.txtHApplicant.value)				
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtHSalesGroup.value)
			strVal = strVal & "&txtSOType=" & Trim(.txtHSOType.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtHFromDt.value)
			strVal = strVal & "&txtToDt=" & Trim(.txtHToDt.value)						
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey     
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtApplicant=" & Trim(.txtApplicant.value)
			strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
			strVal = strVal & "&txtSOType=" & Trim(.txtSOType.value)
			strVal = strVal & "&txtFromDt=" & Trim(.txtFromDt.text)
			strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)			
			strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey
		End If				
			
        strVal = strVal & "&lgPageNo="		 & lgPageNo						        
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        

		Call RunMyBizASP(MyBizASP, strVal)		    						
        
    End With
    
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
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="수입자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConSItemDC 0">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" ALT="수입자명" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" ALT="영업그룹명" SIZE=20 TAG="14"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5>수주형태</TD>
							<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtSOType" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="수주형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoType" align=top TYPE="BUTTON" ONCLICK ="vbscript:OpenConSItemDC 2">&nbsp;<INPUT TYPE=TEXT NAME="txtSoTypeNm" ALT="수주형태명" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>수주일</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/s3111ra6_fpDateTime1_txtFromDt.js'></script>&nbsp;~&nbsp;
								<script language =javascript src='./js/s3111ra6_fpDateTime2_txtToDt.js'></script>
							</TD>	
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
							<script language =javascript src='./js/s3111ra6_vaSpread_vspdData.js'></script>
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
							<IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)"></IMG>
						</TD>
						<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm " WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>
	
<INPUT TYPE=HIDDEN NAME="txtHApplicant" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHPayTerms" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtHSOType" TAG="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
