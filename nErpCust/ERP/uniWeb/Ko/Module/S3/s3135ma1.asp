<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : S3135MA1
'*  4. Program Name         : 수주진행별조회 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2002/02/15
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Ahn tae hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncServer.asp" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgMark                                                  
Dim ArrParam(7)
Dim lgIsOpenPop
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)


'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID        = "s3135mb1.asp"
Const C_MaxKey          = 5                                

Const C_TrackingNo = 1			'Tracking No
Const C_SoNo = 2				'수주번호 
Const C_SoSeq = 3				'수주순번 
Const C_ItemCd = 4				'품목코드 
Const C_ItemNm = 5				'품목명 
Const C_SoQty = 6				'수주량 
Const C_PlanQty = 7				'주생산계획수량 
Const C_ProdQty = 8 			'생산실적수량 
Const C_ProdInQty = 9			'생산입고수량 
Const C_DnQty = 10				'출하수량 
Const C_GiQty = 11				'출고수량 
Const C_BillQty = 12			'매출수량 
                                         
'--------------- 개발자 coding part(변수선언,End)-------------------------------------------------------------

 Dim IsOpenPop						' Popup
 Dim arrValue(3)

'==================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode    
    lgBlnFlgChgValue = False                           'Indicates that no value changed
    lgPageNo     = ""                                  'initializes Previous Key
    lgSortKey        = 1

End Sub

'==================================================================================================================
Sub SetDefaultVal()	
End Sub

'==================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "MA") %>
End Sub

'==================================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("S3136QA1","S","A","V20021106", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )    
    Call SetSpreadLock    
End Sub

'==================================================================================================================
Sub SetSpreadLock()
    With frm1
    .vspdData.ReDraw = False
    ggoSpread.Source = .vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
    .vspdData.ReDraw = True
    End With
End Sub

'==================================================================================================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub

'==================================================================================================================
Sub OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pSpdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub

'==================================================================================================================
Function OpenConSoNo()
	Dim iCalledAspName
	Dim strRet

	If IsOpenPop = True Then Exit Function
			
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("s3111pa1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3111pa1", "x")
		IsOpenPop = False
		exit Function
	end if
	IsOpenPop = True
		
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, ""), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If strRet = "" Then
		Exit Function
	Else
		frm1.txtSoNo.value = strRet
		frm1.txtSoNo.focus
	End If	

End Function	

'==================================================================================================================
Function OpenPurReq()
	Dim iCalledAspName
	Dim strRet
		
	On Error Resume Next
	
	If frm1.vspdData.MaxRows <= 0 Then lgIntFlgMode = OPMD_CMODE
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")				
		Exit Function
	End If
    
    If IsOpenPop = True Then Exit Function
    
    Call vspdData_Click(frm1.vspdData.ActiveCol , frm1.vspdData.ActiveRow)
    
	'20021227 kangjungu dynamic popup
	iCalledAspName = AskPRAspName("m3145ra1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "m3145ra1", "x")
		IsOpenPop = False
		exit Function
	end if
    IsOpenPop = True
    
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrValue), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

    IsOpenPop = False
End Function	

'==================================================================================================================
Function OpenProdReq()
	Dim iCalledAspName
	Dim strRet
				
	On Error Resume Next
	
	If frm1.vspdData.MaxRows <= 0 Then lgIntFlgMode = OPMD_CMODE
	
	If lgIntFlgMode <> parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")				
		Exit Function
	End If

    If IsOpenPop = True Then Exit Function
    
    Call vspdData_Click(frm1.vspdData.ActiveCol , frm1.vspdData.ActiveRow)
    
	iCalledAspName = AskPRAspName("m3146ra1")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "m3146ra1", "x")
		IsOpenPop = False
		exit Function
	end if
    IsOpenPop = True
    
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrValue), _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function	

'==================================================================================================================
Function OpenConDnPopup(ByVal iWhere)
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	Case 1
		'2002-10-07 s3135pa1.asp 추가 
		Dim strRet
		
		Dim arrTNParam(5), i

		For i = 0 to UBound(arrTNParam)
			arrTNParam(i) = ""
		Next	

		'20021227 kangjungu dynamic popup
		iCalledAspName = AskPRAspName("s3135pa1")	
		if Trim(iCalledAspName) = "" then
			IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s3135pa1", "x")
			IsOpenPop = False
			exit Function
		end if

		strRet = window.showModalDialog(iCalledAspName, Array(window.parent, arrTNParam), _
			"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False

		If strRet = "" Then
			Exit Function
		Else
			frm1.txtTrackingNo.value = strRet 
		End If		
		
		frm1.txtTrackingNo.focus
						
	Case 2							
		
		arrParam(1) = "b_item"									
		arrParam(2) = Trim(frm1.txtItemCode.Value)				
		arrParam(3) = ""                             			
		arrParam(4) = "PHANTOM_FLG = " & FilterVar("N", "''", "S") & " "										
		arrParam(5) = "품목"								
	
		arrField(0) = "Item_cd"									
		arrField(1) = "Item_nm"									
		arrField(2) = "Spec"
		
		arrHeader(0) = "품목"								
		arrHeader(1) = "품목명"								
		arrHeader(2) = "규격"			
		
		arrParam(0) = arrParam(5)		
		
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
			
		IsOpenPop = False
        If arrRet(0) = "" Then
			Exit Function
		Else
			Call SetConDnPopup(arrRet,iWhere)
		End If	    		
			
		frm1.txtItemCode.focus
	End Select
	
End Function

'==================================================================================================================
Function SetConDnPopup(Byval arrRet,ByVal iWhere)

	With frm1
		Select Case iWhere
		Case 1
			.txtTrackingNo.value = arrRet(0) 
		Case 2
			.txtItemCode.value = arrRet(0) 
			.txtItemCodeNm.value = arrRet(1)   
		End Select
	End With
	
End Function

'==================================================================================================================
Sub SetQuerySpreadColor(ByVal lRow)
	
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.Source = .vspdData
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspdData.ReDraw = True
    End With
    
End Sub

'==================================================================================================================
Sub Form_Load()
    Call LoadInfTB19029														
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
   
   	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")						

    frm1.txtTrackingNo.focus   

End Sub

'==================================================================================================================
Sub btnBillToPartyOnClick()
	Call OpenBizPartner()
End Sub
'==================================================================================================================
Sub btnTaxBizAreaOnClick()
	Call OpenTaxBizArea()
End Sub
'==================================================================================================================
Sub btnSalesGroupOnClick()
	Call OpenSalesGroup()
End Sub
'==================================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )

	Call SetPopupMenuItemInf("00000000001") 
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
	
    If Row <= 0 Then
       
       ggoSpread.Source = frm1.vspdData
       If lgSortKey = 1 Then
			ggoSpread.SSSort Col		'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If
       
       Exit Sub
    End If 
       
	If Row <> 0 Then
		With frm1.vspdData
			.Row = Row
			.Col = GetKeyPos("A",1)			'C_TrackingNo
			arrValue(0) = .text
		
			.Col = GetKeyPos("A",4)			'C_ItemCd
			arrValue(1) = .text

			.Col = GetKeyPos("A",5)			'C_ItemNm
			arrValue(2) = .text  

			.Col = GetKeyPos("A",2)			'C_SoNo
			arrValue(3) = .text  
		End With
	Else
		arrValue(0) = ""
		arrValue(1) = ""
		arrValue(2) = ""
		arrValue(3) = ""
	End If
	
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row) 
			
End Sub

'==================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'==================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True

End Sub

'==================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub
'==================================================================================================================
Sub rdoTaxTypeFlg1_OnClick()
	frm1.txtTaxRadio.value = frm1.rdoTaxTypeFlg1.value
End Sub

Sub rdoTaxTypeFlg2_OnClick()
	frm1.txtTaxRadio.value = frm1.rdoTaxTypeFlg2.value
End Sub

Sub rdoTaxTypeFlg3_OnClick()
	frm1.txtTaxRadio.value = frm1.rdoTaxTypeFlg3.value
End Sub
	
Sub rdoTexIssueFlg1_OnClick()
	frm1.txtIssueRadio.value = frm1.rdoTexIssueFlg1.value
End Sub

Sub rdoTexIssueFlg2_OnClick()
	frm1.txtIssueRadio.value = frm1.rdoTexIssueFlg2.value
End Sub

Sub rdoTexIssueFlg3_OnClick()
	frm1.txtIssueRadio.value = frm1.rdoTexIssueFlg3.value
End Sub
'==================================================================================================================
Sub txtlssueFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtlssueFrDt.Action = 7
	End If
End Sub

Sub txtlssueToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtlssueToDt.Action = 7
	End If
End Sub

'==================================================================================================================
Sub txtlssueFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtlssueToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'==================================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO, "X", "X")		
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										
    Call InitVariables														

    If Not chkField(Document, "1") Then									
       Exit Function
    End If

    Call DbQuery																

    FncQuery = True																
        
End Function

'==================================================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'==================================================================================================================
Function FncPrev() 
    On Error Resume Next                                                    
End Function

'==================================================================================================================
Function FncNext() 
    On Error Resume Next                                                    
End Function

'==================================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_SINGLEMULTI)
End Function

'==================================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLEMULTI, False)
End Function

'==================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'==================================================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO, "X", "X")   '☜ 바뀐부분 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vb
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'==================================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               
	Call LayerShowHide(1)
    
    With frm1

<%'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------%>
	If lgIntFlgMode = parent.OPMD_UMODE Then  
	    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
	    strVal = strVal & "&txtTrackingNo=" & Trim(.txtHTrackingNo.value)				    
		strVal = strVal & "&txtItemCode=" & Trim(.txtHItemCode.value)
		strVal = strVal & "&txtSoNo=" & Trim(.txtHSoNo.value)        
   	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001									
		strVal = strVal & "&txtTrackingNo=" & Trim(.txtTrackingNo.value)				
		strVal = strVal & "&txtItemCode=" & Trim(.txtItemCode.value)
		strVal = strVal & "&txtSoNo=" & Trim(.txtSoNo.value)		
	End If			
<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>
        strVal = strVal & "&lgPageNo="   & lgPageNo                      '☜: Next key tag
        strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
    
    Call RunMyBizASP(MyBizASP, strVal)										
    End With
    
    DbQuery = True


End Function
   
'==================================================================================================================
Function DbQueryOk()														
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE												
	lgBlnFlgChgValue = False
	
    Call ggoOper.LockField(Document, "Q")									
	Call SetQuerySpreadColor(1)
	Call SetToolbar("11000000000111")

    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
       'frm1.txtTrackingNo.focus	
    End if 
	
End Function


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
		<TD <%=HEIGHT_TYPE_00%>></TD>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>수주진행조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPurReq">구매정보참조</A>|<A href="vbscript:OpenProdReq">생산정보참조</A></TD>
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
									<TD CLASS="TD5" NOWRAP>Tracking No</TD>
									<TD CLASS="TD6"><INPUT NAME="txtTrackingNo" ALT="Tracking No" TYPE="Text" MAXLENGTH=25 SiZE=30 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConDnPopup 1"></TD>
									<TD CLASS="TD5" NOWRAP>품목</TD>
									<TD CLASS="TD6"><INPUT NAME="txtItemCode" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSDN" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConDnPopup 2">&nbsp;<INPUT NAME="txtItemCodeNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>수주번호</TD>
									<TD CLASS="TD6"><INPUT NAME="txtSoNo" ALT="수주번호" TYPE="Text" MAXLENGTH=18 SiZE=30 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSoHDR" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConSoNo()"></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
							<TABLE <%=LR_SPACE_TYPE_20%>>
								<TR>
									<TD HEIGHT="100%">
										<script language =javascript src='./js/s3135ma1_I206714206_vspdData.js'></script>
									</TD>
								</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR>  
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		                        FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">

<!--INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24"-->
<INPUT TYPE=HIDDEN NAME="txtHTrackingNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHItemCode" tag="24"-->
</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
