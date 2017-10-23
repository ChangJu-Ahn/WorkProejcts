<%@ LANGUAGE="VBSCRIPT" %>
<!--======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : 
'*  3. Program ID           : A5140RA1
'*  4. Program Name         : 
'*  5. Program Desc         : Ado query Sample with DBAgent(Multi + Multi)
'*  6. Component List       :
'*  7. Modified date(First) : 2001/04/18
'*  8. Modified date(Last)  : 2003/06/05
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :  2002/11/25 : ASP Standard for Include improvement
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAMain.vbs">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAEvent.vbs">				</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliPAOperation.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliVariables.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentA.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incCliRdsQuery.vbs">			</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/Cookie.vbs">					</SCRIPT>
<SCRIPT LANGUAGE = "JavaScript"SRC = "../../inc/incImage.js">					</SCRIPT>
<SCRIPT LANGUAGE = "VBScript" SRC = "../../inc/incEB.vbs">						</SCRIPT>
<SCRIPT LANGUAGE = "VBScript">

Option Explicit 
'==========================================================================================

Dim arrParent
Dim arrParam
'==========================================================================================
arrParent		= window.dialogArguments
Set PopupParent = arrParent(0)
arrParam		= arrParent(1)

top.document.title = PopupParent.gActivePRAspName

Const C_MASTER = 1
Const C_DETAIL = 2


Const BIZ_PGM_ID        = "a5140rb1.asp"                         '☆: Biz logic spread sheet for #1
Const BIZ_PGM_ID1       = "a5140rb2.asp"                         '☆: Biz logic spread sheet for #2

'==========================================================================================
Const C_MaxKey            = 6                                    '☆☆☆☆: Max key value

'==========================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->

'==========================================================================================
Dim lgDocCur
Dim lgPageNo_A
Dim lgPageNo_B
Dim lgSortKey_A
Dim lgSortKey_B
Dim lgIsOpenPop

'==========================================================================================
Sub InitVariables()

    lgBlnFlgChgValue	= False
    lgIntFlgMode		= PopupParent.OPMD_CMODE

    lgPageNo_A			= ""
    lgSortKey_A			= 1

    lgPageNo_B			= ""
    lgSortKey_B			= 1

End Sub

'==========================================================================================
Sub SetDefaultVal()
	frm1.txtBatchNo.value  = arrParam(0)
End Sub

'==========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<% Call LoadInfTB19029A("Q", "A","NOCOOKIE","RA") %>
	<% Call LoadBNumericFormatA("Q", "A","NOCOOKIE","RA") %>
End Sub

'==========================================================================================
Function CancelClick()
	Self.Close()
End Function

'==========================================================================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("A5140RA1", "S", "A", "V20030510", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X")
	Call SetZAdoSpreadSheet("A5140RA1_DTL", "S", "B", "V20030510", PopupParent.C_SORT_DBAGENT, frm1.vspdData2, C_MaxKey, "X", "X")
	Call SetSpreadLock ("A")
	Call SetSpreadLock ("B")
End Sub

'==========================================================================================
Sub InitComboBox()
End Sub

'==========================================================================================
Sub SetSpreadLock( iOpt )
    If iOpt = "A" Then
       With frm1
          .vspdData.ReDraw = False
          ggoSpread.Source = .vspdData
          ggoSpread.SpreadLock 1 , -1
          .vspdData.ReDraw = True
       End With
    Else
       With frm1
            .vspdData2.ReDraw = False
            ggoSpread.Source = .vspdData2
            ggoSpread.SpreadLock 1, -1
            .vspdData2.ReDraw = True
       End With
    End If
End Sub


'==========================================================================================
Sub Form_Load()
	Call LoadInfTB19029
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")
	Call InitVariables
	Call InitComboBox
	Call SetDefaultVal
	Call InitSpreadSheet()
	Call FncQuery()
End Sub


'==========================================================================================
'   Event Name : GetDucCur()
'   Event Desc :
'==========================================================================================
Function GetDucCur()
	Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    Dim strBizAreaCd, strBizAreaNm
    Dim strSelect
    Dim strFrom
    Dim strWhere
    Dim arrTemp

    GetDucCur = False
    strSelect	= "isnull(doc_cur,'')"
    strFrom		= "a_gl_item"
    strWhere	= "gl_no= " & FilterVar(frm1.txtGlNo.value, "''", "S") & ""

    If CommonQueryRs(strSelect, strFrom, strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		arrTemp		= split(lgF0, Chr(11))
		lgDocCur	= arrTemp(0)
		if Trim(lgDocCur) = "" Then
			GetDucCur = False
		Else
			GetDucCur = True
		End If
	End If

End Function


'==========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'==========================================================================================
Function FncQuery()

    Dim IntRetCD

    FncQuery = False
    Err.Clear

    Call ggoOper.ClearField(Document, "2")
    Call InitVariables 

     If Trim(frm1.txtBatchNo.value) = "" And Trim(frm1.txtRefNo.value) = "" Then
		Call DisplayMsgBox("900014", "X", "X", "X")
		Call CancelClick()
		Exit Function
    End If

	frm1.vspdData.MaxRows = 0
    Call DbQuery(C_MASTER)

    FncQuery = True	
End Function



'==========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub



'==========================================================================================
Sub SetPrintCond(StrEbrFile, VarDateFr, VarDateTo, VarDeptCd, VarBizAreaCd, varGlNoFr, varGlNoTo)
	Dim intRetCd

	StrEbrFile = "a5121ma1"

	VarDateFr = UniConvDateToYYYYMMDD(frm1.txtGlDt.Value, PopupParent.gDateFormat,"")
	VarDateTo = UniConvDateToYYYYMMDD(frm1.txtGlDt.Value, PopupParent.gDateFormat,"")
' 회계전표의 key는 temp_GL_NO이기 때문에 temp_GL_NO만 넘긴다.
	VarDeptCd = "%"
	VarBizAreaCd = "%"
	varGlNoFr = Trim(frm1.txtBatchNo.value)
	varGlNoTo = Trim(frm1.txtBatchNo.value)
End Sub


'==========================================================================================
Function DbQuery(pDirect)
	Dim strVal

    DbQuery = False
    
    Err.Clear 
    
    Select Case pDirect
		Case  C_MASTER
			
			Call LayerShowHide(1)
			    
			With frm1
			'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------
				strVal = BIZ_PGM_ID & "?txtBatchNo=" & Trim(.txtBatchNo.value)
'				strVal = strVal & "&txtRefNo=" & Trim(.txtRefNo.value)
				strVal = strVal & "&txtBatchNo_Alt=" & Trim(.txtBatchNo.Alt)
				strVal = strVal & "&txtRefNo_Alt=" & Trim(.txtRefNo.Alt)
			'--------------- 개발자 coding part(실행로직,End)------------------------------------------------
			    strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
			    strVal = strVal & "&lgPageNo="       & lgPageNo_A                          '☜: Next key tag
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
				strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
				strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
				strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

			    Call RunMyBizASP(MyBizASP, strVal)

			End With

		Case C_Detail
			frm1.vspdData2.MaxRows = 0 
			Call LayerShowHide(1)

			With frm1
				strVal = BIZ_PGM_ID1 & "?txtBatchNo=" & Trim(GetKeyPosVal("A", 3))
				strVal = strVal & "&txtSeq=" & Trim(GetKeyPosVal("A", 4))
			    strVal = strVal & "&lgPageNo="       & lgPageNo_B                          '☜: Next key tag
				strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("B")
				'strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("B")
				strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("B"))

			    Call RunMyBizASP(MyBizASP, strVal)
			End With

	End Select 
    DbQuery = True

End Function


'==========================================================================================
Function DbQueryOk( iOpt)
	Dim lngRows
	Dim strTableid
	Dim strColid
	Dim strColNm
	Dim strMajorCd
	Dim strNmwhere
	Dim arrVal

    lgIntFlgMode     = PopupParent.OPMD_UMODE

	If iOpt = 1 Then
       Call vspdData_Click(1,1)
       frm1.vspdData.focus
	End If

	Call ggoOper.LockField(Document, "Q")
End Function


'==========================================================================================
'	Name : OpenConItemCd()
'	Description : Item PopUp
'==========================================================================================
Function OpenConItemCd()


End Function

'==========================================================================================
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

'==========================================================================================
Sub vspdData_Click( Col,  Row)
    Dim ii

	gMouseClickStatus = "SPC"

    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey_A = 1 Then
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 2
        Else
            ggoSpread.SSSort, lgSortKey_A
            lgSortKey_A = 1
        End If
        Exit Sub
    End If

	If Col < 1 Then Exit Sub
	Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
    Call DbQuery(C_DETAIL)
End Sub

'==========================================================================================
Sub vspdData2_Click( Col,  Row)
    Dim ii
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey_B = 1 Then
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 2
        Else
            ggoSpread.SSSort, lgSortKey_B
            lgSortKey_B = 1
        End If
        Exit Sub
    End If

	gMouseClickStatus = "SP2C"

End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub

'==========================================================================================
Sub vspdData2_MouseDown(Button,Shift,x,y)
	If Button <> "1" And gMouseClickStatus = "SP2C" Then
		gMouseClickStatus = "SP2CR"
	End If

End Sub

'==========================================================================================
Sub vspdData_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgPageNo_A <> "" Then
'           Call DisableToolBar(PopupParent.TBC_QUERY)
           If DbQuery(C_MASTER) = False Then
'              Call RestoreToolBar()
              Exit Sub
           End if
		End If
   End if
End Sub

'==========================================================================================
Sub vspdData2_TopLeftChange( OldLeft ,  OldTop ,  NewLeft ,  NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgPageNo_B <> "" Then
'           Call DisableToolBar(PopupParent.TBC_QUERY)
           If DbQuery(C_DETAIL) = False Then
'              Call RestoreToolBar()
              Exit Sub
          End if
		End If
   End if
End Sub

'==========================================================================================
Sub fpdtFromEnterDt_DblClick(Button)
	If Button = 1 then
		frm1.fpdtFromEnterDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpdtFromEnterDt.Focus
	End if
End Sub
'==========================================================================================
Sub fpdtToEnterDt_DblClick(Button)
	If Button = 1 then
		frm1.fpdtToEnterDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.fpdtToEnterDt.Focus
	End if
End Sub

'==========================================================================================
Sub fpdtFromEnterDt_Keypress(KeyAscii)
	If KeyAscii = 13 Then
	   Call MainQuery()
	End If
End Sub

'==========================================================================================
Sub fpdtToEnterDt_Keypress(KeyAscii)
	If KeyAscii = 13 Then
	   Call MainQuery()
	End If
End Sub


'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		'차변금액 

		ggoOper.FormatFieldByObjectOfCur .txtDrAmt, lgDocCur, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
		'대변금액 
		ggoOper.FormatFieldByObjectOfCur .txtCrAmt, lgDocCur, PopupParent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, PopupParent.gDateFormat, PopupParent.gComNum1000, PopupParent.gComNumDec
	End With

End Sub
'===================================== CurFormatNumSprSheet()  ======================================
'	Name : CurFormatNumSprSheet()
'	Description : 화면에서 일괄적으로 Rounding되는 Numeric Spread Sheet
'====================================================================================================
Sub CurFormatNumSprSheet()

End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>


<BODY SCROLL=NO TABINDEX="-1">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="post">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
				<TR>
					<TD CLASS=TD5 NOWRAP>Batch 번호</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBatchNo" MAXLENGTH="18" SIZE=20  ALT ="Batch 번호" tag="14XXXU"></TD>
					<TD CLASS=TD5 NOWRAP>참조번호</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtRefNo" MAXLENGTH="30" SIZE=32 ALT ="참조번호" tag="14XXXU"></TD>
				</TR>
				<TR>
					<TD CLASS="TD5" NOWRAP>전표일자</TD>
					<TD CLASS="TD6" NOWRAP><INPUT NAME="txtGlDt" ALT="전표일자" SIZE = "10" MAXLENGTH="10" STYLE="TEXT-ALIGN: Center" tag="24X1"></TD>
					<TD CLASS=TD5 NOWRAP>만기일자</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtDueDt" MAXLENGTH="10" SIZE=10 ALT ="만기일자" tag="14XXXU"></TD>
				</TR>
<!--				<TR>
					<TD CLASS=TD5 NOWRAP>차대합계</TD>
					<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5140ra1_OBJECT1_txtDrAmt.js'></script>&nbsp;
					<script language =javascript src='./js/a5140ra1_OBJECT2_txtCrAmt.js'></script></TD>
					<TD CLASS=TD5 NOWRAP>차대합계(자국)</TD>
					<TD CLASS=TD6 NOWRAP><script language =javascript src='./js/a5140ra1_OBJECT3_txtDrLocAmt.js'></script>&nbsp;
					<script language =javascript src='./js/a5140ra1_OBJECT4_txtCrLocAmt.js'></script></TD>
				</TR>
-->				<TR>
					<TD CLASS=TD5 NOWRAP>거래처</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpcd1" MAXLENGTH="30" SIZE=10 ALT ="거래처" tag="14XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtBpcd1Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
					<TD CLASS=TD5 NOWRAP>전표입력경로</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE="Text" NAME="txtInputType" SIZE=10 MAXLENGTH=10 tag="14XXXU" ALT="전표입력경로코드"> <INPUT TYPE="Text" NAME="txtInputTypeNm" SIZE=18 tag="14X" ALT="전표입력경로명"></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>입/출금처</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpcd2" MAXLENGTH="30" SIZE=10 ALT ="입/출금처" tag="14XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtBpcd2Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
					<TD CLASS=TD5 NOWRAP>계산서발행처</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBpcd3" MAXLENGTH="30" SIZE=10 ALT ="계산서발행처" tag="14XXXU">&nbsp;<INPUT TYPE=TEXT NAME="txtBpcd3Nm"  SIZE="20" MAXLENGTH="50" TAG="24"></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>입고번호</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInvDocNo" MAXLENGTH="50" SIZE=32 ALT ="입고번호" tag="14XXXU"></TD>
					<TD CLASS=TD5 NOWRAP>입고일</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtInvDt" MAXLENGTH="10" SIZE=10 ALT ="입고일" tag="14XXXU"></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>선하증권</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBlDocNo" MAXLENGTH="35" SIZE=32 ALT ="선하증권" tag="14XXXU"></TD>
					<TD CLASS=TD5 NOWRAP>선하증권일자</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtBlDt" MAXLENGTH="10" SIZE=10 ALT ="선하증권일자" tag="14XXXU"></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>L/C 번호</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLcDocNo" MAXLENGTH="35" SIZE=32 ALT ="L/C 번호" tag="14XXXU"></TD>
					<TD CLASS=TD5 NOWRAP>L/C 일자</TD>
					<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtLcDt" MAXLENGTH="10" SIZE=10 ALT ="L/C 일자" tag="14XXXU"></TD>
				</TR>
				<TR>
					<TD CLASS=TD5 NOWRAP>비고</TD>
					<TD CLASS=TD6 NOWRAP colspan=3><INPUT TYPE=TEXT NAME="txtGlDesc" MAXLENGTH="128" SIZE=70 ALT ="비고" tag="14XXXU"></TD>
				</TR>
				</TABLE>
			</FIELDSET>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
	</TR>
	<TR HEIGHT=100%>
		<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD>
						<script language =javascript src='./js/a5140ra1_vspdData_vspdData.js'></script>
					</TD>
				</TR>
				<TR HEIGHT="60%">
					<TD>
						<script language =javascript src='./js/a5140ra1_vspdData2_vspdData2.js'></script>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
				<TD ></TD>
				<TD ALIGN=RIGHT> <IMG SRC="../../../CShared/image/query_d.gif"    Style="CURSOR: hand" ALT="Search" NAME="Search" OnClick="FncQuery()"        onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)" >	 </IMG>&nbsp;
								 <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" OnClick="OpenSortPopup()"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" ></IMG>&nbsp;
                                 <IMG SRC="../../../CShared/image/cancel_d.gif"   Style="CURSOR: hand" ALT="CANCEL" NAME="Cancel" OnClick="CancelClick()"     onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)">	 </IMG>&nbsp;&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  tabindex=-1>></IFRAME></TD>
	</TR>
</TABLE>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>
</BODY>
</HTML>
