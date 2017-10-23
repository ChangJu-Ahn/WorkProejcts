<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업관리																	*
'*  2. Function Name        :																			*
'*  3. Program ID           : s3211ma7.asp																*
'*  4. Program Name         : Local L/C현황조회															*
'*  5. Program Desc         : Local L/C현황조회															*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2000/04/10																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Son bum Yeol																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*							  2. 2000/03/22 : Coding Start												*
'********************************************************************************************************
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                              
<!-- #Include file="../../inc/lgvariables.inc" --> 

Const BIZ_PGM_ID = "s3211mb7.asp"												
Const BIZ_PGM_JUMP_ID	= "s3211ma2"
Const C_MaxKey          = 16                                           

Dim IsOpenPop 
Dim lgIsOpenPop   
   
Dim GridLCNo
Dim GridLCDocNo
Dim GridLcAmendSeq
Dim GridCur
Dim GridLocAmt
Dim GridSoNo

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
    lgStrPrevKey	 = ""                       
    lgPageNo         = ""
    lgBlnFlgChgValue = False                    
    lgIntFlgMode	 = Parent.OPMD_CMODE        
    lgSortKey        = 1    
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	frm1.txtFromDate.Text	= StartDate
	frm1.txtToDate.Text		= EndDate
	frm1.txtApplicantCd.focus	
End Sub


'=========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub


'========================================================================================================
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("S3211QA7","S","A","V20030318", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
    
End Sub

'===========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		ggoSpread.SpreadLockWithOddEvenRowColor()
		.vspdData.ReDraw = True
    End With
End Sub

'=========================================================================================================
Function OpenLCHdrRef()
	Dim arrRet
	Dim arrParam(1)
	Dim iCalledAspName
	Dim IntRetCD
	
	On Error Resume Next

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End IF

	If IsOpenPop = True Then Exit Function

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If GridLCNo = "" Then
		Call DisplayMsgBox("209001", Parent.VB_YES_NO, "x", "x" )
		Exit Function
	End IF

	IsOpenPop = True

	arrParam(0) = GridLCNo				' 검색조건이 있을경우 파라미터 
	arrParam(1) = GridSoNo
	
	iCalledAspName = AskPRAspName("s3211ra3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3211ra3", "X")
		IsOpenPop = False
		Exit Function
	End If
	
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function


'==================================================================================================================
Function OpenLCDtlRef()
	Dim arrRet
	Dim arrParam(4)
	Dim iCalledAspName
	Dim IntRetCD
	
	On Error Resume Next

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "x", "x", "x")
		Exit Function
	End IF

	If IsOpenPop = True Then Exit Function

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If GridLCNo = "" Then
		Call DisplayMsgBox("209001", Parent.VB_YES_NO, "x", "x")
		Exit Function
	End IF
	
	IsOpenPop = True

	arrParam(0) = GridLCNo				
	arrParam(1) = GridLCDocNo				
	arrParam(2) = GridLcAmendSeq				
	arrParam(3) = GridCur				
	arrParam(4) = GridLocAmt				

	iCalledAspName = AskPRAspName("s3212ra3")

	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "s3212ra3", "X")
		IsOpenPop = False
		Exit Function
	End If
   	
	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False

End Function

'============================================================================================================
Function OpenConPopup(ByVal iType)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iType
	Case 1		'--수입자--
		arrParam(0) = "개설신청인"						
		arrParam(1) = "B_BIZ_PARTNER"							
		arrParam(2) = Trim(frm1.txtApplicantCd.value)		
		arrParam(3) = Trim(frm1.txtApplicantNm.value)	
		arrParam(4) = "BP_TYPE <= " & FilterVar("CS", "''", "S") & ""							
		arrParam(5) = "개설신청인"							
			
		arrField(0) = "BP_CD"									
		arrField(1) = "BP_NM"									
			
		arrHeader(0) = "개설신청인"							
		arrHeader(1) = "개설신청인명"						

	Case 2		'--영업그룹--
		arrParam(0) = "영업그룹"
		arrParam(1) = "B_SALES_GRP"						
		arrParam(2) = Trim(frm1.txtSalesGrpCd.value)		
		arrParam(3) = Trim(frm1.txtSalesGrpNm.value)	
		arrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "					
		arrParam(5) = "영업그룹"					
		
	    arrField(0) = "SALES_GRP"						
	    arrField(1) = "SALES_GRP_NM"					
	    
	    arrHeader(0) = "영업그룹"					
	    arrHeader(1) = "영업그룹명"					

	Case 3		'--화폐--
		arrParam(0) = "화폐"
		arrParam(1) = "B_CURRENCY"						
		arrParam(2) = Trim(frm1.txtCur.value)		
		arrParam(3) = ""								
		arrParam(4) = ""								
		arrParam(5) = "화폐"					
		
	    arrField(0) = "CURRENCY"						
	    arrField(1) = "CURRENCY_DESC"					
	    
	    arrHeader(0) = "화폐"					
	    arrHeader(1) = "화폐명"					

	Case 4		'--개설은행--
		arrParam(0) = "개설은행"
		arrParam(1) = "B_BANK"								
		arrParam(2) = Trim(frm1.txtOpenBankCd.value)		
		arrParam(3) = Trim(frm1.txtOpenBankNm.value)	
		arrParam(4) = ""									
		arrParam(5) = "개설은행"						
		
	    arrField(0) = "BANK_CD"								
	    arrField(1) = "BANK_NM"							
	    
	    arrHeader(0) = "개설은행"						
	    arrHeader(1) = "개설은행명"						

	End Select
    
    arrParam(3) = ""
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	Select Case iType
	    Case 1
	    	frm1.txtApplicantCd.focus 
	    Case 2
	    	frm1.txtSalesGrpCd.focus
	    Case 3
	    	frm1.txtCur.focus
	    Case 4
	    	frm1.txtOpenBankCd.focus
	End Select

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetConPopup(arrRet,iType)
	End If	
	
End Function

'============================================================================================================
Function SetConPopup(Byval arrRet,ByVal iType)

	Select Case iType
	Case 1
		frm1.txtApplicantCd.value = arrRet(0) 
		frm1.txtApplicantNm.value = arrRet(1)
	Case 2
		frm1.txtSalesGrpCd.value = arrRet(0) 
		frm1.txtSalesGrpNm.value = arrRet(1)   
	Case 3
		frm1.txtCur.value = arrRet(0) 
	Case 4
		frm1.txtOpenBankCd.value = arrRet(0) 
		frm1.txtOpenBankNm.value = arrRet(1)   
	End Select

End Function

'============================================================================================================
Sub CookiePage(Byval Kubun)

	Const CookieSplit = 4877
	
	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)
	
	If Kubun = 1 Then
		WriteCookie CookieSplit , GridLCNo
	End IF
	
End Sub

'============================================================================================================
Function JumpChgCheck()
	
	Dim IntRetCD

	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	Call CookiePage(1)
	Call PgmJump(BIZ_PGM_JUMP_ID)

End Function

'============================================================================================================
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

'============================================================================================================
Sub Form_Load()
    
    Call LoadInfTB19029														
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   
	ReDim lgKeyPos(C_MaxKey)
    ReDim lgKeyPosVal(C_MaxKey)
	Call InitVariables                                                      
	Call SetDefaultVal
	Call InitSpreadSheet
    Call SetToolbar("11000000000011")							
    frm1.txtApplicantCd.focus
    
End Sub

'============================================================================================================
Sub txtFromDate_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromDate.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtFromDate.Focus
    End If
End Sub

Sub txtToDate_DblClick(Button)
    If Button = 1 Then
        frm1.txtToDate.Action = 7
        Call SetFocusToDocument("M")	
        frm1.txtToDate.Focus
    End If
End Sub

'============================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
	
	Call SetPopupMenuItemInf("00000000001")
	
	gMouseClickStatus = "SPC"
	
	Set gActiveSpdSheet = frm1.vspdData

	If Row = 0 Then
		frm1.vspdData.OperationMode = 0
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If
    Else
		frm1.vspdData.OperationMode = 3	
	End If
    
	If Row <> 0 Then
			frm1.vspdData.Row = frm1.vspdData.ActiveRow
			frm1.vspdData.Col = GetKeyPos("A",1)
			GridLCNo = frm1.vspdData.text
			 		
			frm1.vspdData.Col = GetKeyPos("A",2)
			GridLCDocNo = frm1.vspdData.text 		
			
			frm1.vspdData.Col = GetKeyPos("A",3)
			GridLcAmendSeq = frm1.vspdData.text 		
			
			frm1.vspdData.Col = GetKeyPos("A",5)
			GridCur = frm1.vspdData.text 		
			
			frm1.vspdData.Col = GetKeyPos("A",6)
			GridLocAmt = frm1.vspdData.text 		
			
			frm1.vspdData.Col = GetKeyPos("A",16)
			GridSoNo = frm1.vspdData.text
			 		
	Else
		GridLCNo		= ""
		GridLCDocNo		= ""
		GridLcAmendSeq	= ""
		GridCur			= ""
		GridLocAmt		= ""
		GridSoNo		= ""
	End If
	
End Sub

'============================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'============================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'============================================================================================================
Function vspdData_DblClick(Col, Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		End If
	End If
End Function

'============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If CheckRunningBizProcess = True Then
		Exit Sub
	End If	
				
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
    	If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End If

End Sub

'============================================================================================================
Sub txtFromDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
Sub txtToDate_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
Sub txtFromLocAmt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub
Sub txtToLocAmt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'============================================================================================================ 
 Function FncQuery() 

    FncQuery = False                                                        
    
    Err.Clear                                                               
	
    Call ggoOper.ClearField(Document, "2")									
    Call InitVariables 			    
   
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "x", "x")	
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

	If ValidDateCheck(frm1.txtFromDate, frm1.txtToDate) = False Then Exit Function

	If UNICDbl(frm1.txtFromLocAmt.text) > UNICDbl(frm1.txtToLocAmt.text) Then
		Call DisplayMsgBox("970023", "x", frm1.txtToLocAmt.Alt, frm1.txtFromLocAmt.Alt)   		
		frm1.txtFromLocAmt.Focus
		Set gActiveElement = document.activeElement 
		Exit Function
	End If
   
    If DbQuery = False Then Exit Function

    FncQuery = True															

End Function

'============================================================================================================ 
Function FncPrint() 
	FncPrint = False                                                             
    Err.Clear                                                                    
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                       
    FncPrint = True                                                              
End Function

'============================================================================================================
Function FncExcel() 
	FncExcel = False                                                             
    Err.Clear                                                                    

	Call Parent.FncExport(Parent.C_MULTI)

    FncExcel = True                                                              
End Function

'============================================================================================================
Function FncFind() 
    FncFind = False                                                              
    Err.Clear                                                                    

	Call Parent.FncFind(Parent.C_MULTI, True)

    FncFind = True                                                               
End Function

'============================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'============================================================================================================
Function FncExit()
	Dim IntRetCD

    FncExit = False                                                              
    Err.Clear                                                                    
    
    FncExit = True                                                               
End Function

'============================================================================================================
Function DbQuery() 
    
    Dim strVal
    
    DbQuery = False
    
    Err.Clear 	
	Call LayerShowHide(1) 
	
    With frm1

		If lgIntFlgMode = Parent.OPMD_UMODE Then
		    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001			
			strVal = strVal & "&txtApplicantCd=" & Trim(.HApplicantCd.value)
			strVal = strVal & "&txtSalesGrpCd=" & Trim(.HSalesGrpCd.value)
			strVal = strVal & "&txtFromLocAmt=" & Trim(.HFromLocAmt.value)
			strVal = strVal & "&txtToLocAmt=" & Trim(.HToLocAmt.value)
			strVal = strVal & "&txtCur=" & Trim(.HCur.value)
			strVal = strVal & "&txtFromDate=" & Trim(.HFromDate.value)
			strVal = strVal & "&txtToDate=" & Trim(.HToDate.value)
			strVal = strVal & "&txtOpenBankCd=" & Trim(.HOpenBankCd.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001			
			strVal = strVal & "&txtApplicantCd=" & Trim(.txtApplicantCd.value)
			strVal = strVal & "&txtSalesGrpCd=" & Trim(.txtSalesGrpCd.value)
			strVal = strVal & "&txtFromLocAmt=" & Trim(.txtFromLocAmt.Text)
			strVal = strVal & "&txtToLocAmt=" & Trim(.txtToLocAmt.Text)
			strVal = strVal & "&txtCur=" & Trim(.txtCur.value)
			strVal = strVal & "&txtFromDate=" & Trim(.txtFromDate.Text)
			strVal = strVal & "&txtToDate=" & Trim(.txtToDate.Text)
			strVal = strVal & "&txtOpenBankCd=" & Trim(.txtOpenBankCd.value)
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			
		End If

        strVal = strVal & "&lgPageNo="       & lgPageNo                                  
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
	
		
		Call RunMyBizASP(MyBizASP, strVal)										
        
    End With
    
    DbQuery = True

End Function

'============================================================================================================
Function DbQueryOk()														
	
	lgBlnFlgChgValue = False	
    lgIntFlgMode	 = Parent.OPMD_UMODE
	
	Call ggoOper.LockField(Document, "Q")									
	Call SetToolbar("11000000000111")										

    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
    Else
       frm1.txtApplicantCd.focus	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>Local L/C현황조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenLCHdrRef">LOCAL L/C 상세정보</A> | <A href="vbscript:OpenLCDtlRef">LOCAL L/C 내역정보</A>					
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
							<TD CLASS=TD5 NOWRAP>개설신청인</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtApplicantCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMLCBp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup 1">&nbsp;<INPUT NAME="txtApplicantNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
							<TD CLASS=TD5 NOWRAP>영업그룹</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesGrpCd"  TYPE="Text" MAXLENGTH="4" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMLCSaleGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup 2">&nbsp;<INPUT NAME="txtSalesGrpNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>개설금액</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/s3211ma7_fpDoubleSingle1_txtFromLocAmt.js'></script>							
								&nbsp;~&nbsp;
								<script language =javascript src='./js/s3211ma7_fpDoubleSingle2_txtToLocAmt.js'></script>							
							</TD>
							<TD CLASS=TD5 NOWRAP>화폐</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtCur" TYPE="Text" MAXLENGTH="3" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMLCCur" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup 3"></TD>
						</TR>
						<TR>
							<TD CLASS=TD5 NOWRAP>개설일</TD>
							<TD CLASS=TD6 NOWRAP>
								<script language =javascript src='./js/s3211ma7_fpDateTime1_txtFromDate.js'></script>
								&nbsp;~&nbsp;
								<script language =javascript src='./js/s3211ma7_fpDateTime2_txtToDate.js'></script>
							</TD>
							<TD CLASS=TD5 NOWRAP>개설은행</TD>
							<TD CLASS=TD6 NOWRAP><INPUT NAME="txtOpenBankCd" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnMLCBank" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenConPopup 4">&nbsp;<INPUT NAME="txtOpenBankNm" TYPE="Text" MAXLENGTH="30" SIZE=25 tag="14"></TD>
						</TR>
				</TABLE></TD>
			</TR>
			<TR>
				<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
			</TR>
			<TR>
				<TD WIDTH=100% HEIGHT=* valign=top><TABLE <%=LR_SPACE_TYPE_20%>>
					<TR>
						<TD HEIGHT="100%">
							<script language =javascript src='./js/s3211ma7_I453851765_vspdData.js'></script>
						</TD>
					</TR></TABLE>
				</TD>
			</TR>
		</TABLE></TD>
	</TR>
	<TR>
		<td <%=HEIGHT_TYPE_01%>></td>
	</TR>
	<TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
			<TR>
				<TD WIDTH="*" ALIGN=RIGHT><A HREF = "VBSCRIPT:JumpChgCheck()">LOCAL L/C등록</a></TD>				
			</TR></TABLE>
      </TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		                    FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPostGiFlag" tag="14">

<INPUT TYPE=HIDDEN NAME="HApplicantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HSalesGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HFromLocAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="HToLocAmt" tag="24">
<INPUT TYPE=HIDDEN NAME="HCur" tag="24">
<INPUT TYPE=HIDDEN NAME="HFromDate" tag="24">
<INPUT TYPE=HIDDEN NAME="HToDate" tag="24">
<INPUT TYPE=HIDDEN NAME="HOpenBankCd" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41  TABINDEX="-1" src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
