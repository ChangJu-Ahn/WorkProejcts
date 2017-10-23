<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          	: Human Resources
'*  2. Function Name        	: HR(근무Calendar등록)
'*  3. Program ID           	: H4001ma1.asp
'*  4. Program Name         	: H4001ma1.asp
'*  5. Program Desc         	: 근무칼렌다등록 
'*  6. Modified date(First) 	: 2001/05/28
'*  7. Modified date(Last)  	: 2003/06/11
'*  8. Modifier (First)     	: Hwang Jeong-won
'*  9. Modifier (Last)      	: Lee SiNa
'* 10. Comment              	:
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<STYLE TYPE="text/css">
	.Header {height:24; font-weight:bold; text-align:center; color:darkblue}
	.Day {height:22;cursor:Hand;
		font-size:17; font-weight:bold; Border:0; text-align:right}
	.DummyDay {height:22;cursor:;
		font-size:12; font-weight:; Border:0; text-align:right}
</STYLE>
<MAP NAME="CalButton">
	<AREA SHAPE=RECT COORDS="1, 1, 20, 20" ALT="Year -" onClick="ChangeMonth(-12)">
	<AREA SHAPE=RECT COORDS="20, 1, 40, 20" ALT="Month -" onClick="ChangeMonth(-1)">
	<AREA SHAPE=RECT COORDS="40, 1, 60, 20" ALT="Month +" onClick="ChangeMonth(1)">
	<AREA SHAPE=RECT COORDS="60, 1, 80, 20" ALT="Year +" onClick="ChangeMonth(12)">
</MAP>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "H4001mb1.asp"						           '☆: Biz Logic ASP Name
Const CChnageColor    = "#f0fff0"

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop          

Dim lgLastDay
Dim lgStartIndex
Dim lgArrDate(31, 3)

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                                               '⊙: Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                                                '⊙: Indicates that no value changed
    lgIntGrpCount = 0                                                       '⊙: Initializes Group View Size

	Dim iRow, iCol
	For iRow = 1 To 6
		For iCol = 1 To 7
			If frm1.All.tblCal.Rows(iRow).Cells(iCol-1).Style.backgroundColor = CChnageColor Then
				frm1.All.tblCal.Rows(iRow).Cells(iCol-1).Style.backgroundColor = "white"
				frm1.txtDate((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
				frm1.txtDesc((iRow - 1) * 7 + iCol - 1).Style.backgroundColor = "white"
			End If
		Next
	Next

End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay
	
	Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
	frm1.txtValidDt.focus 		
	frm1.txtValidDt.Year = strYear 		 '년월일 default value setting
	frm1.txtValidDt.Month = strMonth 
	
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "H","NOCOOKIE","MA") %>

End Sub
'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   lgKeyStream =               Frm1.txtBA_CD.Value   & parent.gColSep         'You Must append one character(parent.gColSep)
   lgKeyStream = lgKeyStream & Frm1.txtValidDt.Year  & parent.gColSep         'You Must append one character(parent.gColSep)
   lgKeyStream = lgKeyStream & Right("0" & Frm1.txtValidDt.Month,2)  & parent.gColSep         'You Must append one character(parent.gColSep)
   lgKeyStream = lgKeyStream & Frm1.cboWork.Value    & parent.gColSep         'You Must append one character(parent.gColSep)
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr
    Dim iNameArr
        
    Err.Clear                                                               '☜: Clear error no
	On Error Resume Next

    Call CommonQueryRs(" MINOR_CD, MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0047", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iCodeArr = "" & Chr(11) & lgF0
    iNameArr = "" & Chr(11) & lgF1
    Call SetCombo2(frm1.cboWork, iCodeArr, iNameArr, Chr(11))
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtValidDt, parent.gDateFormat, 2)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
	
    Call SetDefaultVal()
	Call SetToolbar("1100100000011111")												'⊙: Set ToolBar
	Call InitVariables

	frm1.txtBA_CD.focus
    Call InitComboBox
	Call CookiePage (0)                                                             '☜: Check Cookie
			
End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub

'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()

    Dim IntRetCD 
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If

    If   txtBA_CD_Onchange() then
        Exit Function
    End If

	Call InitVariables
    
    Call MakeKeyStream("Q")
                                                             '☜: Query db data
    Call DisableToolBar(parent.TBC_QUERY)					'Query 버튼을 disable시킴 

	If DBQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If

    FncQuery = True                                                              '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
                                                        '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = False Then 
        IntRetCD = DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If

    If   txtBA_CD_Onchange()  then
        Exit Function
    End If
    	
    Call DisableToolBar(parent.TBC_SAVE)
	IF DBsave =  False Then
		Call RestoreToolBar()
		Exit Function
	End If
    
    FncSave = True                                                              '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncCopy
' Desc : developer describe this line Called by MainSave in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

     If LayerShowHide(1) = False then
    	Exit Function 
    End if

    strVal = BIZ_PGM_ID & "?txtMode="          & parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
	
    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	 If LayerShowHide(1) = False then
    		Exit Function 
    	End if

	With Frm1
		.txtMode.value        = parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()

End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

	lgIntFlgMode      = parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    Frm1.txtBA_CD.focus 

	Call SetToolbar("1100100000011111")
    Call ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   

End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call InitVariables
    lgIntFlgMode      = parent.OPMD_UMODE

End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'========================================================================================================
' Name : OpenbizareaInfo()
' Desc : developer describe this line
'========================================================================================================
Function OpenbizareaInfo(Byval strCode)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True    

	arrParam(0) = "사업장POPUP"					<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_AREA"						<%' TABLE 명칭 %>
	arrParam(2) = frm1.txtBA_CD.value 						<%' Code Condition%>
	arrParam(3) = ""'frm1.txtBA_NM.value								<%' Name COndition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "사업장코드"			
	
    arrField(0) = "BIZ_AREA_CD"						<%' Field명(0)%>
    arrField(1) = "BIZ_AREA_NM"						<%' Field명(1)%>    
    arrHeader(0) = "사업장코드"					<%' Header명(0)%>
    arrHeader(1) = "사업장명"					<%' Header명(1)%>
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
	    frm1.txtBA_CD.focus
	    Exit Function
	Else
		Call SetbizareaInfo(arrRet)
	End If	

End Function

'========================================================================================================
' Name : SetBizAreaInfo()
' Desc : developer describe this line
'========================================================================================================
Function SetBizAreaInfo(ByVal arrRet)

	With frm1
		.txtBA_CD.value = arrRet(0)
		.txtBA_NM.value = arrRet(1)		
		.txtBA_CD.focus
	End With
	
End Function

Sub DescChange(iDate)
	Dim strDesc
	Dim index
	index = iDate - 1

	If frm1.txtDate(index).className = "DummyDay" Then
		Exit Sub
	End If
	
	Call SetChange(iDate)

	strDesc = frm1.txtDesc(index).value
	frm1.txtDesc(index).value = ""
	
	frm1.txtDesc(index).value = strDesc
	frm1.txtDesc(index).title = strDesc
End Sub

Sub HoliChange(iDate)
	Dim index
	index = iDate - 1

	If frm1.txtDate(index).className = "DummyDay" Then
		Exit Sub
	End If

	Call SetChange(iDate)
	
	If frm1.txtHoli(index).value = "H" Then
		If (index+1) Mod 7 = 0 Then
			frm1.txtDate(index).style.color = "blue"
			frm1.txtHoli(index).value = "S"						
		Else
			frm1.txtDate(index).style.color = "black"
			frm1.txtHoli(index).value = "D"			
		End If

	Else
		frm1.txtDate(index).style.color = "red"
		frm1.txtHoli(index).value = "H"
	End if
End Sub

Sub SetChange(iDate)
	Dim index
	index = iDate - 1

	lgBlnFlgChgValue = True
	
	frm1.All.tblCal.Rows(Int((index+7)/7)).Cells(index Mod 7).Style.backgroundColor = CChnageColor
	frm1.txtDate(index).Style.backgroundColor = CChnageColor
	frm1.txtDesc(index).Style.backgroundColor = CChnageColor
End Sub

Sub ChangeMonth(i)
    Dim strVal
    Dim dtDate
    Dim IntRetCD

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X") '☜ 바뀐부분 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Sub
		End If
    End If

    Call InitVariables	'⊙: Initializes local global variables
	
	On Error Resume Next
	Err.Clear
	
    
    dtDate = UniConvYYYYMMDDToDate(parent.gAPDateFormat, frm1.hYear.value, frm1.hMonth.value, "01")

    If Err.Number <> 0 Then                         'Check if there is retrived data        
        Err.Clear
		Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Sub
    End If

	dtDate = UNIDateAdd("m", i, dtDate, parent.gAPDateFormat)
		
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
    strVal = strVal & "&txtYear=" & Year(dtDate)							'☆: 조회 조건 데이타 
    strVal = strVal & "&txtMonth=" & Month(dtDate)							'☆: 조회 조건 데이타 

	 If LayerShowHide(1) = False then
    		Exit Sub
    	End if
	Call RunMyBizASP(MyBizASP, strVal)
End Sub

'========================================================================================================
'   Event Name : txtBA_CD_OnChange
'   Event Desc :
'========================================================================================================
Function txtBA_CD_OnChange()    
    Dim IntRetCd
 
    If frm1.txtBA_CD.value = "" Then
        frm1.txtBA_NM.value = ""
    ELSE    
        IntRetCd = CommonQueryRs(" biz_area_nm "," b_biz_area "," biz_area_cd =  " & FilterVar(frm1.txtBA_CD.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCd = false Then
            Call DisplayMsgBox("800142","X","X","X")                              
            frm1.txtBA_NM.value = ""
            frm1.txtBA_CD.focus
            Set gActiveElement = document.activeElement   
            txtBA_CD_OnChange=true                                     			
            Exit Function
        Else
            frm1.txtBA_NM.value=Trim(Replace(lgF0,Chr(11),""))
        End If
    End If
End Function

'=======================================
'   Event Name : txtValidDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
'=======================================
Sub txtValidDt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")     
        frm1.txtValidDt.Action = 7
        frm1.txtValidDt.focus
    End If
End Sub
'==========================================================================================
'   Event Name : txtValidDt_KeyDown()
'   Event Desc : 조회조건부의 txtValidDt_KeyDown시 EnterKey일 경우는 Query
'==========================================================================================
Sub txtValidDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery 'Call FncQuery()
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="NO">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>근무카렌더등록</font></td>
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
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>사업장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtBA_CD" MAXLENGTH="10" SIZE=10  ALT ="사업장코드" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="vbscript:Call OpenbizareaInfo(frm1.txtba_cd.value)">
								                        <INPUT NAME="txtBA_NM" MAXLENGTH="50" SIZE=20 ALT ="사업장명" tag="14X"></TD>
								<TD CLASS="TD5">&nbsp;</TD>
								<TD CLASS="TD6">&nbsp;</TD>
							</TR>
							<TR>
								<TD CLASS="TD5">해당년월</TD>
								<TD CLASS="TD6">
								<script language =javascript src='./js/h4001ma1_I482632297_txtValidDt.js'></script></TD>									
								<TD CLASS="TD5">근무조구분</TD>
								<TD CLASS="TD6"><SELECT NAME="cboWork" tag="12X" STYLE="WIDTH: 150px;" ALT="근무조구분"></SELECT></TD>
							</TR>					
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD>
						<TABLE ID="tblCal" WIDTH=100% HEIGHT=100% BORDER=1 CELLSPACING=0 CELLPADDING=0 ALIGN="center">
							<THEAD CLASS="Header">
								<TR>
									<TD>일요일</TD><TD>월요일</TD><TD>화요일</TD><TD>수요일</TD><TD>목요일</TD><TD>금요일</TD><TD>토요일</TD>
								</TR>
				        	</THEAD>
							<TBODY>
<%
Dim i, j, k
k = 1
For i=1 To 6
%>
					            <TR>
<%
	For j=1 To 7
%>
									<TD ALIGN="Center">
										<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 ALIGN="Center">
											<TR>
												<TD ALIGN="Left">
													<INPUT type="hidden" name="txtHoli" size=1 maxlength=1 disabled>
													<INPUT type="text" name="txtDate" class="DummyDay" size=2 maxlength=2 tag=2 
														tabindex=-1 readonly disabled onclick="HoliChange(<%=k%>)">												
												</TD>
											</TR>
											<TR>
												<TD ALIGN="Left">
													<INPUT type="text" name="txtDesc" MaxLength=30 Style="Width:100%;Border:0;text-align:center" disabled tag=2 onchange="DescChange(<%=k%>)" >
												</TD>
											</TR>
										</TABLE>
									</TD>
<%
		k = k + 1
	Next
%>
								</TR>
<%
Next
%>							</TBODY>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD HEIGHT=5 WIDTH=100%></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hYear" tag="24">
<INPUT TYPE=HIDDEN NAME="hMonth" tag="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>
