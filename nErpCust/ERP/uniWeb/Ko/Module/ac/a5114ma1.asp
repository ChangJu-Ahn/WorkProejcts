
<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Cash
'*  3. Program ID           : A5114MA1
'*  4. Program Name         : 현금출납장 
'*  5. Program Desc         : 현금출납장 조회 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/05/24
'*  8. Modified date(Last)  : 2004/01/12
'*  9. Modifier (First)     : Cho Ig Sung
'* 10. Modifier (Last)      : Kim Chang Jin
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->
<!-- #Include file="../../inc/IncSvrHTML.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">


<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAEvent.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliMAOperation.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit																	'☜: indicates that All variables must be declared in advance

'******************************************  1.2 Global 변수/상수 선언  ***********************************
'	1. Constant는 반드시 대문자 표기.
'**********************************************************************************************************

Const BIZ_PGM_ID = "a5114mb1.asp"												'☆: 비지니스 로직 ASP명 
'========================================================================================================= 
Const C_MaxKey          = 1                                    '☆☆☆☆: Max key value
Dim lgIsOpenPop

 '==========================================  1.2.2 Global 변수 선언  =====================================
'	1. 변수 표준에 따름. prefix로 g를 사용함.
'	2.Array인 경우는 ()를 반드시 사용하여 일반 변수와 구별해 됨 
'========================================================================================================= 
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim lgStrGlNo
Dim lgStrItemSeq

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False
    lgIntGrpCount = 0
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""
    lgStrGlNo = ""
    lgStrItemSeq = ""
    lgLngCurRows = 0
    lgPageNo         = ""
    
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
	Dim StartDate, ServerDate
	Dim strYear, strMonth, strDay

	
	Call	ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, "01")
	ServerDate = UniConvYYYYMMDDToDate(Parent.gDateFormat, strYear, strMonth, strDay)
	

	frm1.txtFromGlDt.text	= StartDate
	frm1.txtToGlDt.text		= ServerDate
	

End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
		<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
    	<% Call loadInfTB19029A("Q", "A", "NOCOOKIE", "QA") %>
End Sub


'========================================================================================================= 
Sub InitSpreadSheet()

	Call SetZAdoSpreadSheet("a5114ma1","S","A","V20030221",parent.C_SORT_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
    
End Sub


'========================================================================================================= 
Sub SetSpreadLock(ByVal pOpt)
   If pOpt = "A" Then
        With frm1

        .vspdData.ReDraw = False
        ggoSpread.SpreadLockWithOddEvenRowColor()
        ggoSpread.SSSetProtected	.vspdData.MaxCols,-1,-1
        .vspdData.ReDraw = True

        End With
    End if
End Sub



'------------------------------------------  OpenBizArea()  -------------------------------------------------
'	Name : OpenBizArea()
'	Description : Cost Center PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizArea()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

    arrParam(0) = "사업장 팝업"			' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"					' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizAreaCd.Value)	' Code Condition
	arrParam(3) = ""							' Name Cindition

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If

	arrParam(5) = "사업장 코드"			
	
   	arrField(0) = "BIZ_AREA_CD"	     				' Field명(0)
	arrField(1) = "BIZ_AREA_NM"			    		' Field명(1)
    
	arrHeader(0) = "사업장코드"				' Header명(0)
	arrHeader(1) = "사업장명"				    ' Header명(1)
    
    		
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBizAreaCd.focus
		Exit Function
	Else
		
		Call SetReturnVal(arrRet,1)
	End If	

End Function

'------------------------------------------  OpenBizArea1()  -------------------------------------------------
'	Name : OpenBizArea1()
'	Description : Cost Center PopUp
'--------------------------------------------------------------------------------------------------------- 
Function OpenBizArea1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

    arrParam(0) = "사업장 팝업"					' 팝업 명칭 
	arrParam(1) = "B_BIZ_AREA"						' TABLE 명칭 
	arrParam(2) = Trim(frm1.txtBizAreaCd1.Value)	' Code Condition
	arrParam(3) = ""								' Name Cindition

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
	Else
		arrParam(4) = ""
	End If

	arrParam(5) = "사업장 코드"			
	
   	arrField(0) = "BIZ_AREA_CD"	     				' Field명(0)
	arrField(1) = "BIZ_AREA_NM"			    		' Field명(1)
    
	arrHeader(0) = "사업장코드"					' Header명(0)
	arrHeader(1) = "사업장명"				    ' Header명(1)
    		
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtBizAreaCd1.focus
		Exit Function
	Else
		Call SetReturnVal(arrRet,2)
	End If	

End Function

'========================================================================================
Function OpenPopupGL() 
	Dim arrRet
	Dim arrParam(1)	
	Dim arrField
	Dim intFieldCount
	Dim i
	Dim j

	If lgIsOpenPop = True Then Exit Function
	
	With frm1.vspdData									'z_ado_field_inf의 내용이 바뀌면..이곳을 반드시 확인해야한다.
		If .maxrows > 0 Then	
			arrField = Split(GetSQLSelectListDataType("A"),",")
			intFieldCount = UBound(arrField,1)
			For i = 0 To  intFieldCount -1
				If Trim(arrField(i)) = "C.GL_NO" Then				
					Exit For
				End if
			Next
		
			.Row = .ActiveRow
			.Col = i + 6			
		
			arrParam(0) = Trim(.Text)	'결의전표번호 
			arrParam(1) = ""			'Reference번호 
		End if	
	End With

	lgIsOpenPop = True
   
	arrRet = window.showModalDialog("../../ComAsp/a5120ra1.asp", Array(window.parent,arrParam), _
		     "dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
End Function

'========================================================================================
Function SetReturnVal(byval arrRet,byval field_fg)
	With frm1	
		Select case field_fg
			case 1	'OpenBizArea				
				.txtBizAreaCd.focus
				.txtBizAreaCd.Value		= arrRet(0)
				.txtBizAreaNm.Value		= arrRet(1)
			case 2	'OpenBizArea1				
				.txtBizAreaCd1.focus
				.txtBizAreaCd1.Value	= arrRet(0)
				.txtBizAreaNm1.Value	= arrRet(1)				
		End select	
	End With
End Function

'========================================================================================
Function PopZAdoConfigGrid()
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	ElseIf arrRet(0) = "R" Then
	   Call ggoOper.ClearField(Document, "2")	   
	   Exit Function        
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function


'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                                                     '⊙: Load table , B_numeric_format
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field    
    Call ggoOper.FormatField(Document, "1", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2", ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet
    Call InitVariables                                                      '⊙: Initializes local global variables

     '----------  Coding part  -------------------------------------------------------------
    Call SetDefaultVal
    Call SetToolBar("11000000000001")										'⊙: 버튼 툴바 제어 
    

	' 권한관리 추가 
	Dim xmlDoc
	
	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc) 
	
	' 사업장 
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' 내부부서 
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text
	
	' 내부부서(하위포함)
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text
	
	' 개인 
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text
	
	Set xmlDoc = Nothing

    
    frm1.txtFromGlDt.focus 
    frm1.txtTDrAmt.allownull = False
    frm1.txtTCrAmt.allownull = False
    frm1.txtTSumAmt.allownull = False

    frm1.txtNDrAmt.allownull = False
    frm1.txtNCrAmt.allownull = False
    frm1.txtNSumAmt.allownull = False

    frm1.txtSDrAmt.allownull = False
    frm1.txtSCrAmt.allownull = False
    frm1.txtSSumAmt.allownull = False

End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'========================================================================================

Sub txtFromGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtFromGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromGlDt.Focus
    End If
End Sub

'========================================================================================
Sub txtToGlDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtToGlDt.Action = 7
        Call SetFocusToDocument("M")
        frm1.txtFromGlDt.Focus
    End If
End Sub

'========================================================================================
Sub txtFromGlDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================
Sub txtToGlDt_Change()
    lgBlnFlgChgValue = True
End Sub

'========================================================================================

'========================================================================================================
'   Event Name : txtBizAreaCd_Onchange()
'   Event Desc : 사업장코드를 직접입력할경우에 사업장코드명을 설정해준다.
'========================================================================================================
sub txtBizAreaCd_Onchange()
	Dim strCd
	Dim strWhere
	Dim IntRetCD

	strCd = Trim(frm1.txtBizAreaCd.value)
	If strCd = "" Then
		frm1.txtBizAreaNm.value = ""
	Else
		If lgAuthBizAreaCd <> "" AND UCASE(lgAuthBizAreaCd) <> UCASE(strCd) Then
			frm1.txtBizAreaNm.value = ""
			'IntRetCD = DisplayMsgBox("124200","x","x","x")
		Else
			strWhere = "BIZ_AREA_CD = " & FilterVar(strCd, "''", "S")
			
			Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			if Trim(Replace(lgF0,Chr(11),"")) = "X" then
				frm1.txtBizAreaNm.value = ""
			else
				frm1.txtBizAreaNm.value = Trim(Replace(lgF0,Chr(11),""))
			end if
		End If
	End If
	
End sub

'========================================================================================================
'   Event Name : txtBizAreaCd1_Onchange()
'   Event Desc : 사업장코드를 직접입력할경우에 사업장코드명을 설정해준다.
'========================================================================================================
sub txtBizAreaCd1_Onchange()
	Dim strCd
	Dim strWhere
	Dim IntRetCD

	strCd = Trim(frm1.txtBizAreaCd1.value)
	If strCd = "" Then
		frm1.txtBizAreaNm1.value = ""
	Else
		If lgAuthBizAreaCd <> "" AND UCASE(lgAuthBizAreaCd) <> UCASE(strCd) Then
			frm1.txtBizAreaNm1.value = ""
			'IntRetCD = DisplayMsgBox("124200","x","x","x")
		Else
			strWhere = "BIZ_AREA_CD = " & FilterVar(strCd, "''", "S")
			
			Call CommonQueryRs("BIZ_AREA_NM","B_BIZ_AREA",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
			if Trim(Replace(lgF0,Chr(11),"")) = "X" then
				frm1.txtBizAreaNm1.value = ""
			else
				frm1.txtBizAreaNm1.value = Trim(Replace(lgF0,Chr(11),""))
			end if
		End If
	End If
 
End sub

'========================================================================================
Sub txtFromGlDt_Keypress(Key)
    If Key = 13 Then
		frm1.txtToGlDt.focus
        FncQuery()
    End If
End Sub

'========================================================================================
Sub txtToGlDt_Keypress(Key)
    If Key = 13 Then
		frm1.txtFromGlDt.focus
        FncQuery()
    End If
End Sub

'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"	'Split 상태코드 
    Set gActiveSpdSheet = frm1.vspdData
    
    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If
    End If
	Call SetSpreadColumnValue("A", frm1.vspdData, Col, Row)
End Sub

'========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    With frm1.vspdData 
    If Row >= NewRow Then
        Exit Sub
    End If
    End With
End Sub


'========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
    	If lgPageNo <> "" Then
           Call DisableToolBar(Parent.TBC_QUERY)
           If DbQuery = False Then
              Call RestoreToolBar()
              Exit Sub
           End if
    	End If
    End If
End Sub

'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    Call InitVariables
    															'⊙: Initializes local global variables
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Function
    End If
    
    If UniConvDateToYYYYMMDD(frm1.txtFromGlDt.text,Parent.gDateFormat,"") > UniConvDateToYYYYMMDD(frm1.txtToGlDt.text, Parent.gDateFormat,"")Then
		IntRetCD = DisplayMsgBox("113118", "X", "X", "X")			'⊙: "Will you destory previous data"
		Exit Function
    End If
    
     '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data
       
    FncQuery = True																'⊙: Processing is OK
    
End Function




'========================================================================================
Function FncPrint() 
    On Error Resume Next 
    Parent.Fncprint()
End Function


'========================================================================================
Function FncExcel() 
    Call Parent.FncExport(Parent.C_MULTI)
End Function


'========================================================================================
Function FncFind() 
    Call Parent.FncFind(Parent.C_MULTI, False)
End Function

'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If
    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)
End Sub

'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	
    FncExit = True
End Function

'========================================================================================
Function DbQuery() 
	Dim strVal
    DbQuery = False
    
	On Error Resume Next                                                    '☜: Protect system from crashing		       
    Err.Clear                                                               '☜: Protect system from crashing
        
	Call LayerShowHide(1)
	
    With frm1
	    'If lgIntFlgMode = Parent.OPMD_UMODE Then
		'	strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜: 
		'	strVal = strVal & "&txtBizAreaCd=" & Trim(.hBIZ_AREA_CD.value)				'☆: 조회 조건 데이타 
		'	strVal = strVal & "&txtBizAreaCd_Alt=" & Trim(.txtBizAreaCd.Alt)
		'	strVal = strVal & "&txtFromGlDt=" & Trim(.hFromGlDt.value)
		'	strVal = strVal & "&txtToGlDt=" & Trim(.hToGlDt.value)
		'	strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		'	strVal = strVal & "&txtfiscenddt=" & Trim(.hfiscEndDt.value)
		'Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001							'☜: 
		strVal = strVal & "&txtBizAreaCd=" & .txtBizAreaCd.value				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtBizAreaCd_Alt=" & Trim(.txtBizAreaCd.Alt)
		strVal = strVal & "&txtBizAreaCd1=" & .txtBizAreaCd1.value				'☆: 조회 조건 데이타 
		strVal = strVal & "&txtBizAreaCd1_Alt=" & Trim(.txtBizAreaCd1.Alt)			
		strVal = strVal & "&txtFromGlDt=" & Trim(.txtFromGlDt.text)
		strVal = strVal & "&txtToGlDt=" & Trim(.txtToGlDt.text)
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
		strVal = strVal & "&txtfiscenddt=" & Trim(.hfiscEndDt.value)
		'End If 
		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '☜: Next key tag
	    strVal = strVal & "&lgPageNo="       & lgPageNo         
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))	
			   
		' 권한관리 추가 
		strVal = strVal & "&lgAuthBizAreaCd="	& lgAuthBizAreaCd			' 사업장 
		strVal = strVal & "&lgInternalCd="		& lgInternalCd				' 내부부서 
		strVal = strVal & "&lgSubInternalCd="	& lgSubInternalCd			' 내부부서(하위포함)
		strVal = strVal & "&lgAuthUsrID="		& lgAuthUsrID				' 개인 

		Call RunMyBizASP(MyBizASP, strVal)
    End With
    
    DbQuery = True
    
End Function

'========================================================================================
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE
        
    Call ggoOper.LockField(Document, "Q")
	Call SetToolBar("11000000000001")
	frm1.vspdData.focus
	Set gActiveElement = document.ActiveElement
End Function


'========================================================================================
Function DbSave()
End Function


'========================================================================================
Function DbSaveOk()
End Function


'========================================================================================
Function DbDelete() 
End Function




</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>


<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=RIGHT><A HREF="VBSCRIPT:OpenPopupGL()">회계전표</A></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%>
				</TR>
				<TR>
				    <TD HEIGHT=20 WIDTH=100%>	
						<FIELDSET CLASS="CLSFLD">
							<TABLE  <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>회계일</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name="txtFromGlDt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="회계일" tag="12" id=fpDateTime1></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
	                                                       <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name="txtToGlDt" CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="회계일" tag="12" id=fpDateTime2></OBJECT>');</SCRIPT></TD>
									<TD CLASS="TD5" NOWRAP>사업장코드</TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd"  SIZE=13 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="11XXXU" ALT="사업장코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizArea()">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm" SIZE=28 tag="14">&nbsp;~</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP><INPUT CLASS="clstxt" TYPE=TEXT NAME="txtBizAreaCd1" SIZE=13 MAXLENGTH=20 STYLE="TEXT-ALIGN: left" tag="11XXXU" ALT="사업장코드"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCostCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBizArea1()">&nbsp;<INPUT TYPE=TEXT NAME="txtBizAreaNm1" SIZE=28 tag="14"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH="100%"></TD>
				</TR>
				<TR>	
					<TD WIDTH="100%" HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD HEIGHT="100%" colspan=7>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" > <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
								</TD>
							</TR>
							<TR>
								<TD CLASS=TD5 Align=middle NOWRAP>이월금액</TD>
								<TD CLASS=TD5 NOWRAP>입금</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="이월금액(차변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>출금</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="이월금액(대변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>잔액</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtTSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="이월금액(잔액)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
							</TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>발생금액</TD>
								<TD CLASS=TD5 NOWRAP>입금</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="당기금액(차변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>출금</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="당기금액(대변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>수지</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtNSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="당기금액(잔액)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>
    					    </TR>
							<TR>
								<TD CLASS=TD5 NOWRAP>누계금액</TD>
								<TD CLASS=TD5 NOWRAP>입금</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSDrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="당기금액(대변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>출금</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSCrAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="당기금액(차변)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>								
								<TD CLASS=TD5 NOWRAP>잔액</TD>
								<TD CLASS=TD5 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> name="txtSSumAmt" style="HEIGHT: 20px; LEFT: 0px; TOP: 0px; WIDTH: 150px" title=FPDOUBLESINGLE ALT="당기금액(잔액)" tag="24X2" id=OBJECT1></OBJECT>');</SCRIPT></TD>														
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hBIZ_AREA_CD" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hFromGlDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hToGlDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hfiscDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="hfiscEndDt" tag="24" TABINDEX="-1">
	
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

