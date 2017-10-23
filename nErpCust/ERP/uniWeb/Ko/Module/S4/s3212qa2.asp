<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업관리 
'*  2. Function Name        : 
'*  3. Program ID           : s3212qa2
'*  4. Program Name         : Local L/C 상세조회 
'*  5. Program Desc         : Local L/C 상세조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2002/08/23
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Park in sik
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              
<!-- #Include file="../../inc/lgvariables.inc" --> 

Const BIZ_PGM_ID        = "s3212qb2.asp"
Const BIZ_PGM_JUMP_ID	= "s3212ma2"
Const C_MaxKey          = 18                                   

Dim IsOpenPop 
Dim lgIsOpenPop
Dim IscookieSplit   
Dim iDBSYSDate
Dim EndDate, StartDate

Dim lgKeyPos
Dim lgKeyPosVal

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

'========================================================================================================= 
Sub InitVariables()
    lgBlnFlgChgValue = False                               
    lgStrPrevKey     = ""                                  
    lgSortKey        = 1	
	lgPageNo         = ""
    lgIntFlgMode     = parent.OPMD_CMODE                          	
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtOpenFrDt.text = StartDate
	frm1.txtOpenToDt.text = EndDate
End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================================================================================= 
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("S3212QA2","S","A","V20030318", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
End Sub

'========================================================================================================= 
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    End With
End Sub

'========================================================================================================= 
Function OpenConSItemDC(Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case iWhere
	Case 0
		arrParam(1) = "B_BIZ_PARTNER"						
		arrParam(2) = Trim(frm1.txtconBp_cd.Value)			
		arrParam(3) = ""									
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"				
		arrParam(5) = "개설신청인"							
	
		arrField(0) = "BP_CD"								
		arrField(1) = "BP_NM"								
    
		arrHeader(0) = "개설신청인"							
		arrHeader(1) = "개설신청인명"						

	Case 1
		arrParam(1) = "B_ITEM"								
		arrParam(2) = Trim(frm1.txtItem_cd.Value)			
		arrParam(3) = Trim(frm1.txtItem_Nm.Value)			
		arrParam(4) = ""									
		arrParam(5) = "품목"							
	
		arrField(0) = "ITEM_CD"								
		arrField(1) = "ITEM_NM"								
                arrField(2) = "SPEC"
    
		arrHeader(0) = "품목"							
		arrHeader(1) = "품목명"							
                arrHeader(2) = "규격"


	Case 2
		arrParam(1) = "B_SALES_ORG"							
		arrParam(2) = Trim(frm1.txtSalesOrg.Value)			
		arrParam(3) = Trim(frm1.txtSalesOrgNm.Value)		
		arrParam(4) = ""									
		arrParam(5) = "영업조직"						
	
		arrField(0) = "SALES_ORG"							
		arrField(1) = "SALES_ORG_NM"						
    
		arrHeader(0) = "영업조직"						
		arrHeader(1) = "영업조직명"						

	Case 3
		arrParam(1) = "B_SALES_GRP"							
		arrParam(2) = Trim(frm1.txtSalesGroup.Value)		
		arrParam(3) = Trim(frm1.txtSalesGroupNm.Value)		
		arrParam(4) = ""									
		arrParam(5) = "영업그룹"						
	
		arrField(0) = "SALES_GRP"							
		arrField(1) = "SALES_GRP_NM"						
    
		arrHeader(0) = "영업그룹"						
		arrHeader(1) = "영업그룹명"						

	End Select

	arrParam(0) = arrParam(5)								
	arrParam(3) = ""
        If iWhere = 1 Then
			arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Else
	                arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		       "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
        End If
	
	lgIsOpenPop = False
	
	With frm1
		Select Case iWhere
		    Case 0
		    	.txtconBp_cd.focus
		    Case 1
		    	.txtItem_cd.focus 
		    Case 2
		    	.txtSalesOrg.focus
		    Case 3
		    	.txtSalesGroup.focus
		End Select
	End With	

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConSItemDC(arrRet, iWhere)
	End If	
	
End Function

'========================================================================================================= 
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		Case 0
			.txtconBp_cd.value = arrRet(0) 
			.txtconBp_Nm.value = arrRet(1)   
		Case 1
			.txtItem_cd.value = arrRet(0) 
			.txtItem_Nm.value = arrRet(1)   
		Case 2
			.txtSalesOrg.value = arrRet(0)
			.txtSalesOrgNm.value = arrRet(1)  
		Case 3
			.txtSalesGroup.value = arrRet(0) 
			.txtSalesGroupNm.value = arrRet(1)   
		End Select
	End With
End Function

'========================================================================================================= 
Function CookiePage(ByVal Kubun)

	Dim strTemp, arrVal
	Dim strCookie, i

	Const CookieSplit = 4877						

	If Kubun = 1 Then								

		Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
		
		WriteCookie CookieSplit , IsCookieSplit					
		Call PgmJump(BIZ_PGM_JUMP_ID)

	ElseIf Kubun = 0 Then							

		strTemp = ReadCookie(CookieSplit)

		If strTemp = "" then Exit Function

		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then 
			WriteCookie CookieSplit , ""
			Exit Function
		End If
		Dim iniSep

		frm1.txtconBp_cd.value =  arrVal(0)
		frm1.txtconBp_Nm.value =  arrVal(1)
		frm1.txtBillType.value =  arrVal(2)
		frm1.txtBillTypeNm.value = arrVal(3) 
		frm1.txtSalesOrg.value =  arrVal(4)
		frm1.txtSalesOrgNm.value = arrVal(5) 
		frm1.txtSalesGroup.value =  arrVal(6)
		frm1.txtSalesGroupNm.value = arrVal(7) 
		frm1.txtItem_cd.value =  arrVal(8)
		frm1.txtItem_Nm.value = arrVal(9)

		If Err.number <> 0 Then
			Err.Clear
			WriteCookie CookieSplit , ""
			Exit Function 
		End If

		Call MainQuery()

		WriteCookie CookieSplit , ""

	End IF
End Function

'========================================================================================================= 
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================= 
Function PopZAdoConfigGrid()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'========================================================================================================= 
Sub Form_Load()
    Call LoadInfTB19029														
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
    ReDim lgKeyPos(C_MaxKey)
    ReDim lgKeyPosVal(C_MaxKey)
	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")							
    
    frm1.txtconBp_cd.focus	
    
End Sub

'========================================================================================================= 
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    Call SetPopupMenuItemInf("00000000001")
    
    gMouseClickStatus = "SPC"
    
    Set gActiveSpdSheet = frm1.vspdData
       
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
	If Row <> 0 Then
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
		frm1.vspdData.Col = GetKeyPos("A",10)
		IscookieSplit = frm1.vspdData.text
	Else
		IscookieSplit = ""
	End if
    
End Sub

'========================================================================================================= 
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If    
   	
	If CheckRunningBizProcess = True Then
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

'========================================================================================================= 
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================= 
Sub rdoTexIssueFlg1_OnClick()
	frm1.txtRadio.value = frm1.rdoTexIssueFlg1.value
End Sub

Sub rdoTexIssueFlg2_OnClick()
	frm1.txtRadio.value = frm1.rdoTexIssueFlg2.value
End Sub

Sub rdoTexIssueFlg3_OnClick()
	frm1.txtRadio.value = frm1.rdoTexIssueFlg3.value
End Sub
	
'========================================================================================================= 
Sub txtOpenFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtOpenFrDt.Action = 7
	End If
End Sub

Sub txtOpenToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtOpenToDt.Action = 7
	End If
End Sub

'========================================================================================================= 
Sub txtOpenFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
Sub txtOpenToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub


'========================================================================================================= 
Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 														
	
	If ValidDateCheck(frm1.txtOpenFrDt, frm1.txtOpenToDt) = False Then Exit Function

    Call DbQuery															'☜: Query db data

    FncQuery = True		
End Function

'=========================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'=========================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'=========================================================================================================
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                    
End Function

'=========================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    FncExit = True
End Function

'=========================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               
				
		If   LayerShowHide(1) = False Then
		         Exit Function 
		End If


    
    With frm1

		If lgIntFlgMode = parent.OPMD_UMODE Then	
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
		strVal = strVal & "&txtconBp_cd=" & Trim(.HconBp_cd.value)
		strVal = strVal & "&txtSalesGroup=" & Trim(.HSalesGroup.value)
		strVal = strVal & "&txtItem_cd=" & Trim(.HItem_cd.value)
		strVal = strVal & "&txtOpenFrDt=" & Trim(.HOpenFrDt.value)
		strVal = strVal & "&txtOpenToDt=" & Trim(.HOpenToDt.value)
		Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
		strVal = strVal & "&txtconBp_cd=" & Trim(.txtconBp_cd.value)
		strVal = strVal & "&txtSalesGroup=" & Trim(.txtSalesGroup.value)
		strVal = strVal & "&txtItem_cd=" & Trim(.txtItem_cd.value)
		strVal = strVal & "&txtOpenFrDt=" & Trim(.txtOpenFrDt.text)
		strVal = strVal & "&txtOpenToDt=" & Trim(.txtOpenToDt.text)
		End if

        strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      
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
	lgIntFlgMode = parent.OPMD_UMODE
								
    Call SetToolbar("11000000000111")							
    
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
       frm1.txtconBp_cd.focus	
    End if  
        
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>>&nbsp;</TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Local L/C상세</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>					
					<TD WIDTH="*">&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>개설신청인</TD>
									<TD CLASS="TD6"><INPUT NAME="txtconBp_cd" ALT="개설신청인" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 0">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" SIZE=20 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>품목</TD>
									<TD CLASS=TD6><INPUT NAME="txtItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnStoRo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 1">&nbsp;<INPUT NAME="txtItem_Nm" TYPE="Text" SIZE=20 tag="14"></TD>
								</TR>	
								<TR>	
									<TD CLASS=TD5 NOWRAP>영업그룹</TD>
									<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtSalesGroup" SIZE=10 MAXLENGTH=4 TAG="11XXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesGroup" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC 3">&nbsp;<INPUT TYPE=TEXT NAME="txtSalesGroupNm" SIZE=20 TAG="14"></TD>
									<TD CLASS=TD5 NOWRAP>개설일</TD>
									<TD CLASS=TD6 NOWRAP>
										<script language =javascript src='./js/s3212qa2_fpDateTime1_txtOpenFrDt.js'></script>&nbsp;~&nbsp;
										<script language =javascript src='./js/s3212qa2_fpDateTime2_txtOpenToDt.js'></script>
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
								<TD HEIGHT="100%">
								<script language =javascript src='./js/s3212qa2_vaSpread1_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<td <%=HEIGHT_TYPE_01%>></td>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%"><TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH="*" ALIGN=RIGHT><a href = "vbscript:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">LOCAL L/C내역등록</a></TD>
			</TR>
		</TABLE></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		                    FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">

<INPUT TYPE=HIDDEN NAME="HconBp_cd" tag="24"> 
<INPUT TYPE=HIDDEN NAME="HSalesGroup" tag="24">
<INPUT TYPE=HIDDEN NAME="HItem_cd" tag="24">
<INPUT TYPE=HIDDEN NAME="HOpenFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HOpenToDt" tag="24">

<INPUT TYPE=HIDDEN NAME="txtRadio" tag="14">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41  TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
