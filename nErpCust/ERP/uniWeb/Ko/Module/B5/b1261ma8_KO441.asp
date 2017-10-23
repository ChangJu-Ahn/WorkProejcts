<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : sales																		*
'*  2. Function Name        : 기준정보 																	*
'*  3. Program ID           : B1261MA8    																*
'*  4. Program Name         : 거래처조회																*
'*  5. Program Desc         : 거래처조회																*
'*  6. Comproxy List        :																			*
'*  7. Modified date(First) : 2001/01/05																*
'*  8. Modified date(Last)  : 2001/01/05																*
'*  9. Modifier (First)     : Kim Hyungsuk																*
'* 10. Modifier (Last)      : Sonbumyeol
'* 11. Comment              :	
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									
'*                            this mark(⊙) Means that "may  change"									
'*                            this mark(☆) Means that "must change"									
'* 13. History              : 
'********************************************************************************************************
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit				'☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" --> 

Dim lgIsOpenPop                                             <%'☜: Popup status                          %> 

Dim lgTypeCD                                                <%'☜: 'G' is for group , 'S' is for Sort    %>
Dim lgFieldCD                                               <%'☜: 필드 코드값                           %>
Dim lgFieldNM                                               <%'☜: 필드 설명값                           %>
Dim lgFieldLen                                              <%'☜: 필드 폭(Spreadsheet관련)              %>
Dim lgFieldType                                             <%'☜: 필드 설명값                           %>
Dim lgDefaultT                                              <%'☜: 필드 기본값                           %>
Dim lgNextSeq                                               <%'☜: 필드 Pair값                           %>
Dim lgKeyTag                                                <%'☜: Key 정보                                %>
Dim lgNextSeq_T                                             <%'☜: 필드 Pair값                           %>
Dim lgKeyTag_T                                              <%'☜: Key 정보                              %>

Dim lgSortTitleNm                                           <%'☜: Orderby popup용 데이타(필드설명)      %>
Dim lgSortFieldCD1                                          <%'☜: Orderby popup용 데이타(필드코드)      %>

Dim lgMark                                                  <%'☜: 마크                                  %>
Dim lgKeyPos                                                <%'☜: Key위치                               %>
Dim lgKeyPosVal                                             <%'☜: Key위치 Value                         %>

Dim IsOpenPop
Dim arrParam	

'--------------- 개발자 coding part(변수선언,Start)-----------------------------------------------------------
Const BIZ_PGM_ID = "b1261mb8_ko441.asp"	
Const BIZ_PGM_JUMP_ID = "b1261ma1_ko441"											'☆: 비지니스 로직 ASP명 

Const C_MaxKey          = 2                                    '☆☆☆☆: Max key value

<% '----------------  공통 Global 변수값 정의  ----------------------------------------------------------- %>
Dim lsConcd
Dim lsConNm          
<% '++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ %>

'========================================================================================================= 
Sub InitVariables()

   lgBlnFlgChgValue = False                               'Indicates that no value changed
   lgStrPrevKey     = ""                                  'initializes Previous Key
   lgPageNo         = ""
   lgSortKey        = 1
   lgIntFlgMode = parent.OPMD_CMODE	
    
End Sub

'========================================================================================================
Sub SetDefaultVal()
	
'--------------- 개발자 coding part(실행로직,Start)--------------------------------------------------%>
	frm1.txtBp_cdFrom.focus
'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------%>
End Sub

'========================================================================================================
<% '== 조회,출력 == %>
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>
End Sub


'========================================================================================================
Sub InitSpreadSheet()    

	Call SetZAdoSpreadSheet("B1261MA8_KO441","S","A","V20030425", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock     
      
End Sub
		
'========================================================================================================
Sub SetSpreadLock()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'========================================================================================================
Sub SetSpreadColor(ByVal lRow)
    
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired	1, lRow, lRow
    ggoSpread.SSSetProtected 2, lRow, lRow
    ggoSpread.SSSetProtected 3, lRow, lRow
    ggoSpread.SSSetProtected 5, lRow, lRow
    ggoSpread.SSSetProtected 6, lRow, lRow
	ggoSpread.SSSetProtected 7, lRow, lRow
    ggoSpread.SSSetProtected 9, lRow, lRow
    ggoSpread.SSSetProtected 11, lRow, lRow
    ggoSpread.SSSetProtected 12, lRow, lRow
    .vspdData.ReDraw = True
    
    End With

End Sub

'========================================================================================================
Function PopZAdoConfigGrid()
	
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & Parent.SORTW_WIDTH & "px; dialogHeight=" & Parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'========================================================================================================
Function OpenBp_cd(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "거래처"							<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"						<%' TABLE 명칭 %>

	If iWhere = 0 Then
		arrParam(2) = Trim(frm1.txtBp_cdFrom.value)		<%' Code Condition%>
		arrParam(3) = ""								<%' Name Cindition%>

		frm1.txtBp_cdFrom.focus 
	Else
		arrParam(2) = Trim(frm1.txtBp_cdTo.value)		<%' Code Condition%>
		arrParam(3) = ""	
														<%' Name Cindition%>
		frm1.txtBp_cdTo.focus 
	End If

	arrParam(4) = ""									<%' Where Condition%>
	arrParam(5) = "거래처"							<%' TextBox 명칭 %>
	
    arrField(0) = "BP_CD"								<%' Field명(0)%>
    arrField(1) = "BP_NM"								<%' Field명(1)%>
    
    arrHeader(0) = "거래처"							<%' Header명(0)%>
    arrHeader(1) = "거래처약칭"						<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetBpCode(arrRet, iWhere)
	End If	
	
End Function

'========================================================================================================= 
Function SetBpCode(Byval arrRet, Byval iWhere)

	With frm1

		If iWhere = 0 Then
			.txtBp_cdFrom.value = arrRet(0) 
			.txtBp_nmFrom.value = arrRet(1)   
		Else
			.txtBp_cdTo.value = arrRet(0) 
			.txtBp_nmTo.value = arrRet(1)   
		End If

	End With
	
End Function

'========================================================================================================= 
Function CookiePage(ByVal Kubun)
	
	On Error Resume Next

	Const CookieSplit = 4877						<%'Cookie Split String : CookiePage Function Use%>
	
	Dim strTemp, arrVal

	Call vspdData_Click(frm1.vspdData.ActiveCol,frm1.vspdData.ActiveRow)

	If Kubun = 1 Then

		WriteCookie CookieSplit , lsConcd  & parent.gRowSep & lsConnm

	ElseIf Kubun = 0 Then

		strTemp = ReadCookie(CookieSplit)
			
		If strTemp = "" then Exit Function
			
		arrVal = Split(strTemp, parent.gRowSep)

		If arrVal(0) = "" Then Exit Function
		
		frm1.txtBp_cdFrom.value =  arrVal(0)
		frm1.txtBp_nmFrom.value =  arrVal(1)

		If Err.number <> 0 then
			Err.Clear
			WriteCookie CookieSplit , ""
			exit function
		end if
		
		Call MainQuery()		
			
		WriteCookie CookieSplit , ""
				
	End If
End Function


'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029										'⊙: Load table , B_numeric_format
    
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)	
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
		
    '----------  Coding part  -------------------------------------------------------------

	Call InitVariables														    '⊙: Initializes local global variables
	Call SetDefaultVal		
	Call InitSpreadSheet()
    Call SetToolbar("11000000000011")				'⊙: 버튼 툴바 제어 
	Call CookiePage(0)
	frm1.txtBp_cdFrom.focus
  
End Sub

'========================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub

'==========================================================================================
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
			ggoSpread.SSSort Col				'Sort in ascending
			lgSortKey = 2
	   Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in descending
			lgSortKey = 1
       End If
       
       Exit Sub
    End If   

	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",1)
	lsConcd=frm1.vspdData.Text
    
	frm1.vspdData.Row = Row
	frm1.vspdData.Col = GetKeyPos("A",2)
	lsConnm=frm1.vspdData.Text    
	
    
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)

	
End Sub

'==========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

'==========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

    lgBlnFlgChgValue = True
    
End Sub

Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
End Sub

'========================================================================================
Sub vspdData_TopLeftChange(OldLeft , OldTop , NewLeft , NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    
	If CheckRunningBizProcess = True Then  Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
    	If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
    	End If
    End If   
	
    
End Sub

'========================================================================================
Function FncQuery() 

    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
	
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
	
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")						'⊙: Clear Contents  Field
    Call InitVariables 											'⊙: Initializes local global variables
    Call SetDefaultVal
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then							'⊙: This function check indispensable field
       Exit Function
    End If
	
    '-----------------------
    'Radio Button Check area
    '-----------------------
	If frm1.rdoUsage_flag1.checked = True Then
		frm1.txtRadioFlag.value  = "" 
	ElseIf frm1.rdoUsage_flag2.checked = True Then
		frm1.txtRadioFlag.value = "Y" 
	ElseIf frm1.rdoUsage_flag3.checked = True Then
		frm1.txtRadioFlag.value = "N" 
	End If

	If frm1.rdoBp_type1.checked = True Then 
		frm1.txtRadioType.value = ""
	ElseIf frm1.rdoBp_type2.checked = True Then
		frm1.txtRadioType.value = "S"
	ElseIf frm1.rdoBp_type3.checked = True Then
		frm1.txtRadioType.value = "C"
	ElseIf frm1.rdoBp_type4.checked = True Then
		frm1.txtRadioType.value = "*"
	End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery																'☜: Query db data

    FncQuery = True																'⊙: Processing is OK
	
End Function

'========================================================================================
Function FncPrint() 
    ggoSpread.Source = frm1.vspdData
	Call parent.FncPrint()
End Function

'========================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function

'========================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_MULTI, False)
End Function

'========================================================================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = frm1.vspdData.MaxCols
    
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
       iRet = DisplayMsgBox("900030", "X", iColumnLimit , "X")
       Exit Function  
    End If   
    
    Frm1.vspdData.ScrollBars = SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = SS_SCROLLBAR_BOTH
    
End Function


'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", VB_YES_NO,"x","x")   '☜ 바뀐부분 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vb
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      

    DbQuery = False
    
    
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If

    
    Err.Clear                                                               '☜: Protect system from crashing

	Dim strVal
	    
    With frm1
		If lgIntFlgMode = parent.OPMD_UMODE Then 
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtBp_cdFrom=" & Trim(.HBp_cdFrom.value)
			strVal = strVal & "&txtBp_cdTo=" & Trim(.HBp_cdTo.value)
			strVal = strVal & "&txtRadioFlag=" & Trim(.HRadioFlag.value)
			strVal = strVal & "&txtRadioType=" & Trim(.HRadioType.value)	
            strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey              '☜: Next key tag
            			
		else
			
			strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001			
			strVal = strVal & "&txtBp_cdFrom=" & Trim(.txtBp_cdFrom.value)
			strVal = strVal & "&txtBp_cdTo=" & Trim(.txtBp_cdTo.value)
			strVal = strVal & "&txtRadioFlag=" & Trim(.txtRadioFlag.value)
			strVal = strVal & "&txtRadioType=" & Trim(.txtRadioType.value)	
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
            strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
			
		end if
		<%'--------------- 개발자 coding part(실행로직,End)------------------------------------------------%>	
			strVal = strVal & "&lgPageNo="       & lgPageNo                '☜: Next key tag
			strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")			 
			strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
			strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))
			
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
        
    End With
   
    DbQuery = True

End Function

'========================================================================================
Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
	lgIntFlgMode = parent.OPMD_UMODE                   'Indicates that current mode is Update mode
	lgBlnFlgChgValue = False
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")									'⊙: This function lock the suitable field
	'폴더/조회/입력 
	'/삭제/저장/한줄In
	'/한줄Out/취소/이전 
	'/다음/복사/엑셀 
	'/인쇄/찾기 
    Call SetToolbar("11000000000111")										'⊙: 버튼 툴바 제어 
    
    If frm1.vspdData.MaxRows > 0 Then
       frm1.vspdData.Focus
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>거래처조회(KO441)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					<TD WIDTH=* align=right></TD>
					<TD WIDTH=10>&nbsp;</TD>	
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
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
									<TD CLASS="TD5" NOWRAP>거래처</TD>
									<TD CLASS="TD6" NOWRAP>
										<input NAME="txtBp_cdFrom" TYPE="Text" MAXLENGTH="10" tag="11XXXU" size="10"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgBp_cdFrom" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenBp_cd frm1.txtBp_cdFrom.value,0"> 
										<input NAME="txtBp_nmFrom" TYPE="Text" tag="14" size="20"></TD>
									<TD CLASS="TD5" NOWRAP>거래처구분</TD>
									<TD CLASS="TD6" NOWRAP style="TEXT-ALIGN:center;">
										<input type=radio CLASS="RADIO" name="rdoBp_type" id="rdoBp_type1" value="CS" tag = "11" checked>
											<label for="rdoBp_type1">전체</label>
										<input type=radio CLASS = "RADIO" name="rdoBp_type" id="rdoBp_type2" value="S" tag = "11">
											<label for="rdoBp_type2">매입</label>
										<input type=radio CLASS="RADIO" name="rdoBp_type" id="rdoBp_type3" value="C" tag = "11">
											<label for="rdoBp_type3">매출</label>
										<input type=radio CLASS = "RADIO" name="rdoBp_type" id="rdoBp_type4" value="*" tag = "11">
											<label for="rdoBp_type4">기타</label></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP></td>
									<TD CLASS="TD6" NOWRAP>~
										<input NAME="txtBp_cdTo" TYPE="Text" MAXLENGTH="10" tag="11XXXU" size="10"><img SRC="../../../CShared/image/btnPopup.gif" NAME="ImgBp_cdTo" align="top" TYPE="BUTTON" WIDTH="16" HEIGHT="20" ONCLICK="vbscript:OpenBp_cd frm1.txtBp_cdTo.value,1"> 
										<input NAME="txtBp_nmTo" TYPE="Text" tag="14" size="20"></TD>
									<TD CLASS="TD5" NOWRAP>사용여부</TD>
									<TD CLASS="TD6" NOWRAP style="TEXT-ALIGN:center;">
									<input type=radio CLASS="RADIO" name="rdoUsage_flag" id="rdoUsage_flag1" value="" tag = "11" checked>
										<label for="rdoUsage_flag1">전체</label>
									<input type=radio CLASS = "RADIO" name="rdoUsage_flag" id="rdoUsage_flag2" value="Y" tag = "11">
										<label for="rdoUsage_flag2">사용</label>
									<input type=radio CLASS="RADIO" name="rdoUsage_flag" id="rdoUsage_flag3" value="N" tag = "11">
										<label for="rdoUsage_flag3">미사용</label></TD>
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
										<script language =javascript src='./js/b1261ma8_I726795876_vspdData.js'></script>
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
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
			    <TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=* ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID)" ONCLICK="VBSCRIPT:CookiePage 1">거래처등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../Blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> 
		                FRAMEBORDER=0 SCROLLING=No noresize framespacing=0 TABINDEX="-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>

<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">

<INPUT TYPE=HIDDEN NAME="HBp_cdFrom" tag="24">
<INPUT TYPE=HIDDEN NAME="HBp_cdTo" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadioFlag" tag="24">
<INPUT TYPE=HIDDEN NAME="HRadioType" tag="24">
			
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm" TABINDEX="-1" ></iframe>
</DIV>
</BODY>
</HTML>
