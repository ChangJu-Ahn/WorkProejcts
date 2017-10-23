
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : Basis Architect
'*  2. Function Name        : Master Data(Calendar 생성)
'*  3. Program ID           : B1503ma1.asp
'*  4. Program Name         : B1503ma1.asp
'*  5. Program Desc         : 칼렌다생성 
'*  6. Modified date(First) : 2000/10/02
'*  7. Modified date(Last)  : 2002/12/16
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Sim Hae Young
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_Query  = "B1502mb1.asp"
Const BIZ_PGM_ID = "B1503mb1.asp"
Const BIZ_PGM_COMMON_HOL = "B1502ma1"
Const BIZ_PGM_CHANGE_CAL = "B1501ma1"

Dim C_Month
Dim C_Day
Dim C_Remark

<!-- #Include file="../../inc/lgvariables.inc" -->

Sub InitSpreadPosVariables()
    C_Month     = 1
    C_Day       = 2
    C_Remark    = 3
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	ggoSpread.Source = frm1.vspdData	
   'patch version
    ggoSpread.Spreadinit "V20021202",,parent.gAllowDragDropSpread    
    
	.ReDraw = false

    .MaxCols = C_Remark + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
	.Col = .MaxCols														'☆: 사용자 별 Hidden Column
	.ColHidden = True    
	       
    .MaxRows = 0
    ggoSpread.ClearSpreadData
	
	Call AppendNumberPlace("6","2","0")
    Call GetSpreadColumnPos("A")  

	ggoSpread.SSSetFloat C_Month,"월" ,8,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","12"
    ggoSpread.SSSetFloat C_Day,"일" ,8,"6",ggStrIntegeralPart, ggStrDeciPointPart,parent.gComNum1000,parent.gComNumDec,,,,"1","31"
    ggoSpread.SSSetEdit C_Remark, "사유", 39,,,30
	
	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False
    ggoSpread.SpreadLock C_Month, -1, C_Month
    ggoSpread.SpreadLock C_Day, -1, C_Day
    ggoSpread.SpreadLock C_Remark, -1, C_Remark
	ggoSpread.SSSetProtected .vspdData.MaxCols, -1, -1
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired		C_Month, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_Day, pvStartRow, pvEndRow
    ggoSpread.SSSetRequired		C_Remark, pvStartRow, pvEndRow
    .vspdData.ReDraw = True
    
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_Month     = iCurColumnPos(1)
            C_Day       = iCurColumnPos(2)
            C_Remark    = iCurColumnPos(3)
    End Select    
End Sub

Function LoadCommonHol()
    
    PgmJump(BIZ_PGM_COMMON_HOL)

End Function

Function LoadChangeCal()
    
    PgmJump(BIZ_PGM_CHANGE_CAL)

End Function

Sub Form_Load()

    Dim strYear
    Dim strMonth
    Dim strDay
    

    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
                                                                                <%'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call InitSpreadSheet                                                    '⊙: Setup the Spread sheet

    Call SetToolbar("1000000000001111")										'⊙: 버튼 툴바 제어 
    Call DbQuery
    
	Call ggoOper.FormatDate(frm1.txtYear, parent.gDateFormat, 3)


    Call ExtractDateFrom("<%= GetSvrDate %>",parent.gServerDateFormat, parent.gServerDateType, strYear,strMonth,strDay)
    
    frm1.txtYear.Year  = strYear
	
	frm1.txtYear.focus
End Sub

Function FncQuery()

End Function

Function FncExit()
    FncExit = True
End Function

Function FncPrint()
    Call parent.FncPrint()
End Function

Function FncFind()
    Call parent.FncFind(parent.C_MULTI, False)
End Function

Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000011111") 

    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
    End If

	frm1.vspdData.Row = Row

End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
	
End Sub

Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

Function DbQuery() 
    Dim strVal
    
    Err.Clear      
    DbQuery = False    
    
    With frm1    
        strVal = BIZ_Query & "?txtMode=" & parent.UID_M0001							'☜: 
        strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows        
	    
	    Call RunMyBizASP(MyBizAsp, strVal)										'☜: 비지니스 ASP 를 가동    
    End With
    
    DbQuery = True
End Function

Function DbQueryOk()														'☆: 조회 성공후 실행로직 
	
End Function

Function DbSaveOk()			
    '☆: 카렌다생성후 메시지 
    Call DisplayMsgBox("183114", "X", "X", "X")
End Function

Function btnBatch_OnClick()
	Dim intRetCD
	Dim strCount
	
    If frm1.txtyear.text="" then
        Call DisplayMsgBox("121214", "X", "X", "X")
        Exit Function
    End If
    
    ''이미 생성된 자료유무체크 
    Call CommonQueryRs(" Count(*) " , " B_CALENDAR ", " year(calendar_dt) =  " & Trim(frm1.txtyear.text), _	
    	               lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)    
	strCount = lgF0
	
	If strCount > 0 then    
	    IntRetCD = DisplayMsgBox("800397", parent.VB_YES_NO, "X", "X")'''800397
	Else
	    IntRetCD = DisplayMsgBox("900018", parent.VB_YES_NO, "X", "X")'''900018
	End If
	
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	Call chkCheckBox ()
	
    With frm1
		.txtInsrtUserId.value = parent.gUsrID
		
		Call LayerShowHide(1)
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동 
    
    End With    
End Function


Function chkCheckBox()

	If frm1.chkSun.checked = true Then
		frm1.chkSun.value = "Y"
	Else
		frm1.chkSun.value = "N"
	End If
	
	If frm1.chkMon.checked = true Then
		frm1.chkMon.value = "Y"
	Else
		frm1.chkMon.value = "N"
	End If	
	
	If frm1.chkTue.checked = true Then
		frm1.chkTue.value = "Y"
	Else
		frm1.chkTue.value = "N"
	End If

	If frm1.chkWed.checked = true Then
		frm1.chkWed.value = "Y"
	Else
		frm1.chkWed.value = "N"
	End If

	If frm1.chkThu.checked = true Then
		frm1.chkThu.value = "Y"
	Else
		frm1.chkThu.value = "N"
	End If

	If frm1.chkFri.checked = true Then
		frm1.chkFri.value = "Y"
	Else
		frm1.chkFri.value = "N"
	End If

	If frm1.chkSat.checked = true Then
		frm1.chkSat.value = "Y"
	Else
		frm1.chkSat.value = "N"
	End If
	
End Function

Sub txtYear_Keypress(Key)
    If Key = 13 Then
        call btnBatch_OnClick()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	


</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB3" CELLSPACING=0 CELLPADDING=0>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>칼렌다생성</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5">생성년도</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/b1503ma1_fpDateTime1_txtYear.js'></script></TD>
									<DIV  style="display:none;"><input type="text" ID="txtDummy" NAME="txtDummy" TITLE="txtDummy"></div></TD>
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
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD COLSPAN=2 ALIGN=left>공휴일 선택사항</TD>
							</TR>
							<TR>
								<TD COLSPAN=2 ALIGN=left><HR></TD>
							</TR>
							<TR>
								<TD COLSPAN=2 ALIGN=left>1.요일 선택</TD>
							</TR>							
							<TR>
								<TD COLSPAN=2>
									<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="Y" NAME="chkSun" checked>일
								    <INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkMon">월
									<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkTue">화
									<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkWed">수
									<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkThu">목
									<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkFri">금
									<INPUT TYPE=CHECKBOX CLASS="RADIO" TAG="21X" VALUE="N" NAME="chkSat">토
								</TD>
							</TR>
							<TR>
								<TD COLSPAN=2 ALIGN=left><HR></TD>
							</TR>
							<TR>
								<TD COLSPAN=2 ALIGN=left>2.지정 공휴일</TD>
							</TR>					
							<TR>
								<TD HEIGHT="100%" WIDTH="100%" COLSPAN=2>
								<script language =javascript src='./js/b1503ma1_vspdData_vspdData.js'></script>
								</TD>
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
	<TR HEIGHT=20>
		<TD>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>					
					<TD><BUTTON NAME="btnBatch" CLASS="CLSMBTN" Flag=1>칼렌다 생성</BUTTON></TD>
					<TD WIDTH=* ALIGN=RIGHT><A href="vbscript:LoadCommonHol">공통휴일등록</A>&nbsp;|&nbsp;<A href="vbscript:LoadChangeCal">칼렌다수정</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>	
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="b1502mb1.asp" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

