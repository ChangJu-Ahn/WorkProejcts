<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'************************************************************************************
'*  1. Module Name          : 영업
'*  2. Function Name        : 
'*  3. Program ID           : s3211ra6.asp
'*  4. Program Name         : L/C참조(통관등록에서)
'*  5. Program Desc         : L/C참조(통관등록에서)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/04
'*  8. Modified date(Last)  : 2002/04/27
'*  9. Modifier (First)     : Kim Hyungsuk
'* 10. Modifier (Last)      : Kwak Eunkyoung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 
'*							: -2000/04/04 : 화면 design
'*                          : -2002/04/27 : ADO변환
'**************************************************************************************
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit                              

Const BIZ_PGM_ID 		= "s3211rb6.asp"                             
Const C_MaxKey          = 4                                           
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim IsOpenPop  
Dim gblnWinEvent											'☜: ShowModal Dialog(PopUp) 														   
Dim arrReturn												'☜: Return Parameter Group
Dim arrParam
Dim arrParent
Dim lgIsOpenPop

arrParent = window.dialogArguments
Set PopupParent = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, PopupParent.gServerDateFormat, PopupParent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, PopupParent.gDateFormat)

'========================================================================================================
Function InitVariables()
    Redim arrReturn(0)        

	lgStrPrevKey     = ""								   
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          
    lgSortKey        = 1   
    
    gblnWinEvent = False    
    Self.Returnvalue = arrReturn     
End Function
'=======================================================================================================
Sub SetDefaultVal()
	frm1.txtFromDt.text = StartDate
	frm1.txtToDt.text = EndDate
	If lgSGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtSalesGroup, "Q") 
        	frm1.txtSalesGroup.value = lgSGCd
	End If
End Sub
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S", "NOCOOKIE", "RA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "RA") %>
End Sub
'========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("S3211RA6","S","A","V20030318",PopupParent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
	Call SetSpreadLock       
    
End Sub
'========================================================================================================
Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor
End Sub	
'========================================================================================================
Function OKClick()

	
	With frm1.vspdData
		If .ActiveRow > 0 Then	
			Redim arrReturn(1)
			.Row = .ActiveRow
			
			.Col = GetKeyPos("A",1)		' L/C 번호
			arrReturn(0) = .Text
			
			.Col = GetKeyPos("A",4)		' L/C 종류
			arrReturn(1) = .Text
		End If
	End With

	Self.Returnvalue = arrReturn
	Self.Close()

End Function
'========================================================================================================
Function CancelClick()
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
				iArrParam(1) = "B_BIZ_PARTNER"								
				iArrParam(2) = Trim(.txtApplicant.value)						
				iArrParam(3) = ""											
				iArrParam(4) = "BP_TYPE In (" & FilterVar("CS", "''", "S") & "," & FilterVar("C", "''", "S") & " )"						
				iArrParam(5) = .txtApplicant.alt							
				
			    iArrField(0) = "BP_CD"										
			    iArrField(1) = "BP_NM"										
			    
			    iArrHeader(0) = .txtApplicant.alt							
			    iArrHeader(1) = .txtApplicantNm.Alt							
			    
			    .txtApplicant.focus
			    
			' 영업그룹    
			Case 1
			If frm1.txtSalesGroup.className = "protected" Then
				gblnWinEvent = False
				Exit Function
			End if	
				iArrParam(1) = "B_SALES_GRP"
				iArrParam(2) = Trim(.txtSalesGroup.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
				iArrParam(5) = .txtSalesGroup.Alt
	
				iArrField(0) = "SALES_GRP"
				iArrField(1) = "SALES_GRP_NM"
	
				iArrHeader(0) = .txtSalesGroup.Alt
				iArrHeader(1) = .txtSalesGroupNm.Alt
				
				.txtSalesGroup.focus
				
			' 화폐단위
			Case 2
				iArrParam(1) = "B_CURRENCY"
				iArrParam(2) = Trim(.txtCurrency.value)
				iArrParam(3) = ""
				iArrParam(4) = ""
				iArrParam(5) = .txtCurrency.Alt
	
				iArrField(0) = "Currency"
				iArrField(1) = "Currency_desc"
	
				iArrHeader(0) = .txtCurrency.Alt
				iArrHeader(1) = "화폐명"
				
				.txtCurrency.focus
				
		End Select
	End With
	
	iArrParam(0) = iArrParam(5)												' 팝업 명칭	

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
			.txtApplicant.value 	= arrRet(0)
			.txtApplicantNm.value	= arrRet(1)
			.txtApplicant.focus
		Case 1		
			.txtSalesGroup.value 	= arrRet(0)
			.txtSalesGroupNm.value 	= arrRet(1)
			.txtSalesGroup.focus
		Case 2		
			.txtCurrency.Value		= arrRet(0)
			.txtCurrency.focus
		End Select
	End With
End Function
'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029													

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)    	
	Call ggoOper.LockField(Document, "N")                             
	Call InitVariables	
	Call GetValue_ko441()										  
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
	if Button = 1 then
		frm1.txtFromDt.Action = 7
		Call SetFocusToDocument("P")
        frm1.txtFromDt.Focus
	End if
End Sub

Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
		Call SetFocusToDocument("P")
        frm1.txtToDt.Focus
	End if
End Sub
'=====================================================================================================
Sub txtCurrency_Change()
	ggoOper.FormatFieldByObjectOfCur txtDocAmt, txtCurrency.value, ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, gDateFormat, gComNum1000,gComNumDec
End Sub
'=======================================================================================================
Sub txtFromDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

Sub txtToDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

Sub txtDocAmt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub
'********************************************************************************************************* %>
Function FncQuery() 
    
    FncQuery = False                                                        
    
    Err.Clear                                                              
	
	
	If ValidDateCheck(frm1.txtFromDt, frm1.txtToDt) = False Then Exit Function
        						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData

    Call InitVariables 														

    frm1.vspdData.Maxrows = 0

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
			strVal = strVal & "&txtApplicant="		& Trim(.txtHApplicant.value)			
			strVal = strVal & "&txtSalesGroup="		& Trim(.txtHSalesGroup.value)			
			strVal = strVal & "&txtFromDt="			& Trim(.txtHFromDt.value)				
			strVal = strVal & "&txtToDt="			& Trim(.txtHToDt.value)					
			strVal = strVal & "&txtCurrency="		& Trim(.txtHCurrency.value)				
			strVal = strVal & "&txtDocAmt="			& Trim(.txtHDocAmt.value)				
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey								
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001																																
			strVal = strVal & "&txtApplicant="		& Trim(.txtApplicant.value)				
			strVal = strVal & "&txtSalesGroup="		& Trim(.txtSalesGroup.value)			
			strVal = strVal & "&txtFromDt="			& Trim(.txtFromDt.text)			
			strVal = strVal & "&txtToDt="			& Trim(.txtToDt.text)			
			strVal = strVal & "&txtCurrency="		& Trim(.txtCurrency.value)		
			strVal = strVal & "&txtDocAmt="			& Trim(.txtDocAmt.value)		
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey						
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
		frm1.txtCurrency.focus
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
							<TD CLASS=TD5 NOWRAP>수입자</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtApplicant" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="수입자"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnApplicant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;<INPUT TYPE=TEXT NAME="txtApplicantNm" ALT="수입자명" SIZE=20 TAG="14"></TD>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 1">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" ALT="영업그룹명" SIZE=20 TAG="14"></TD>				
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>화폐</TD>
							<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtCurrency" SIZE=10 MAXLENGTH=3 TAG="11XXXU" ALT="화폐"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCurrency" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 2"></TD>
							<TD CLASS=TD5 NOWRAP>개설금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<TABLE CELLSPACING=0 HEIGHT=100%>
									<TR>										
										<TD><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=fpDoubleSingle3 NAME="txtDocAmt" style="HEIGHT: 20px; WIDTH: 150px" tag="11X2Z" ALT="개설금액" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
									</TR>
								</TABLE>
							</TD>								
						</TR>
						<TR>	
							<TD CLASS=TD5 NOWRAP>개설일</TD>
							<TD CLASS=TD6 NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtFromDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="개설시작일"></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtToDt" CLASS=FPDTYYYYMMDD tag="11X1" Title="FPDATETIME" ALT="개설종료일"></OBJECT>');</SCRIPT>
							</TD>
							<TD CLASS=TD5 NOWRAP></TD>
							<TD CLASS=TD6 NOWRAP></TD>
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
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" id=vaSpread TITLE="SPREAD"> <PARAM NAME="MaxRows" Value=0> <PARAM NAME="MaxCols" Value=0> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
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
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
											  <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC= "../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME></TD>
		</TR>
	</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHApplicant" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHSalesGroup" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHFromDt" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHToDt" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHDocAmt" TAG="14">
<INPUT TYPE=HIDDEN NAME="txtHCurrency" TAG="14">
</FORM>

<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
