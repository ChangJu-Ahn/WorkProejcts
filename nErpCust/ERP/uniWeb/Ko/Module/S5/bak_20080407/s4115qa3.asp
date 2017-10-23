<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4115QA3
'*  4. Program Name         : 출고현황조회(출고일자)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/05/23
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kwakeunkyoung
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

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMaOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit                             '☜: indicates that All variables must be declared in advance
	
Const BIZ_PGM_ID 		= "S4115QB3.asp"                              
Const C_MaxKey          = 9					                          '☆: SpreadSheet의 키의 갯수 

<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================
'=                       2.4 User-defind Variables
'========================================
Dim lgIsOpenPop                                          
Dim IsOpenPop  

Dim lgCookValue 

Dim lgSaveRow 

Dim lgStrColorFlag

<% 
   BaseDate     = GetSvrDate                                                        
%>  

Dim FromDateOfDB
Dim ToDateOfDB

FromDateOfDB	= UNIConvDateAToB(UniDateAdd("m",-1,"<%=BaseDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)
ToDateOfDB		= UNIConvDateAToB(UniDateAdd("m", 0,"<%=BaseDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)

Dim lgStartRow
Dim lgEndRow

Const C_PopDnType	=	0

'========================================	
Sub InitVariables()

    lgStrPrevKey     = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                      
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    
End Sub

'========================================
Sub SetDefaultVal()

	With frm1
		.txtConFromDt.Text	= cstr(FromDateOfDB)
		.txtConToDt.Text	= cstr(ToDateOfDB)
		.rdoConf.checked		= True
		.txtConRdoFlag.value	= .rdoConf.value   
	End With
	
End Sub							
	
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>

End Sub


'========================================
Sub InitSpreadSheet()

    Call SetZAdoSpreadSheet("S4115QA3","S","A", "V20030527", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
    Call SetSpreadLock()   
     
End Sub


'========================================
Sub SetSpreadLock()

    With frm1
		.vspdData.ReDraw = False
	
		
		ggoSpread.SpreadLock 1 , -1
		
	
		.vspdData.ReDraw = True
    End With
    
End Sub

'========================================
Sub Form_Load()

    Call LoadInfTB19029				    

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolBar("1100000000001111")										
    
    frm1.txtConFromDt.focus   

End Sub


'========================================
Function FncQuery() 

    On Error Resume Next                                                          
    Err.Clear                                                                     

	If ValidDateCheck(frm1.txtConFromDt, frm1.txtConToDt) = False Then Exit Function

    FncQuery = False                                                              
    
    Call ggoOper.ClearField(Document, "2")									      
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.ClearSpreadData()
    
    Call InitVariables 														      
    
    If Not chkField(Document, "1") Then								              
       Exit Function
    End If

    If DbQuery = False Then 
       Exit Function
    End If   

    If Err.number = 0 Then
       FncQuery = True                                                             
    End If   

    Set gActiveElement = document.ActiveElement  
    
End Function


'========================================
Function FncPrint()

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncPrint = False                                                              
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        

    If Err.number = 0 Then
       FncPrint = True                                                            
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================
Function FncExcel() 

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncExcel = False                                                              

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncExport(parent.C_MULTI)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncExcel = True                                                             
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================
Function FncFind() 

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncFind = False                                                               

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call Parent.FncFind(parent.C_MULTI, True)
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then
       FncFind = True                                                             
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'=====================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================
Function FncExit()

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncExit = False                                                               
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then
       FncExit = True                                                             
    End If   

    Set gActiveElement = document.ActiveElement   

End Function


'========================================
Function DbQuery() 

	Dim strVal

    On Error Resume Next                                                          
    Err.Clear                                                                     

    DbQuery = False
    
	Call LayerShowHide(1)


        strVal = BIZ_PGM_ID
        
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
        If lgIntFlgMode  <> parent.OPMD_UMODE Then									'☜: This means that it is first search
			
			With frm1

				.txtHConFromDt.value	= .txtConFromDt.text
				.txtHConToDt.value		= .txtConToDt.text	
			
				.txtHConDnType.value		= .txtConDnType.value

				.txtHConRdoFlag.value		= .txtConRdoFlag.value
							
				.txtHlgSelectListDT.value	= GetSQLSelectListDataType("A") 
				.txtHlgTailList.value		= MakeSQLGroupOrderByList("A")
				.txtHlgSelectList.value		= EnCoding(GetSQLSelectList("A"))
			End With	
        End If    
        

        lgStartRow = frm1.vspdData.MaxRows + 1										'포멧팅 적용하는 시작Row
        
    '--------- Developer Coding Part (End) ------------------------------------------------------------

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    If Err.number = 0 Then
       DbQuery = True																
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQueryOk()												

    On Error Resume Next                                                          
    Err.Clear                                                                     

	lgBlnFlgChgValue = False
    lgIntFlgMode     = parent.OPMD_UMODE										  '⊙: Indicates that current mode is Update mode
    lgSaveRow        = 1
    
	'--------- Developer Coding Part (Start) ----------------------------------------------------------

    Call SetQuerySpreadColor
	
	
	'--------- Developer Coding Part (End) ----------------------------------------------------------
	
    Set gActiveElement = document.ActiveElement   

End Function

'=========출고상세현황================================================================================
Function OpenGIDtlRef()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(7)	
	
	On Error Resume Next

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		frm1.txtConDnType.focus
		Exit Function
	End IF

	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
	
	With frm1
			
		.vspdData.Row	= .vspdData.ActiveRow
		
		.vspdData.Col	= GetKeyPos("A",1)
		If .vspdData.Text = "0" Then
			
			.vspdData.Col	= GetKeyPos("A",2) '출하형태 
			arrParam(3)		= .vspdData.Text

			.vspdData.Col	= GetKeyPos("A",3) '출하형태명 
			arrParam(4)		= .vspdData.Text

			.vspdData.Col	= GetKeyPos("A",4) '출고일 
			arrParam(5)		= .vspdData.Text
			arrParam(6)		= .vspdData.Text
					
		ElseIf .vspdData.text = "1" Then

			.vspdData.Col	= GetKeyPos("A",2) '출하형태 
			arrParam(3)		= .vspdData.Text

			.vspdData.Col	= GetKeyPos("A",3) '출하형태명 
			arrParam(4)		= .vspdData.Text

			arrParam(5)		= .txtConFromDt.text
			arrParam(6)		= .txtConToDt.text									
						
		Else 

			arrParam(3)		= ""			   '출하형태 
			arrParam(4)		= ""			   '출하형태명 

			arrParam(5)		= .txtConFromDt.text
			arrParam(6)		= .txtConToDt.text									
		
		End If

		arrParam(0)		= ""					'출하번호 
		arrParam(1)		= ""					'납품처 
		arrParam(2)		= ""					'납품처명 
																	
		If .rdoConf.checked = True Then				 									
			arrParam(7) = .rdoConf.value
		ElseIf .rdoNonConf.checked = True Then
			arrParam(7) = .rdoNonConf.value			
		End If	
		
	End With
	   
	iCalledAspName = AskPRAspName("s4116pa4")	

	If Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s4116pa4", "x")
		lgIsOpenPop = False
		Exit Function
	End if

	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent, arrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False


End Function


'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere

	'출하유형 
	Case C_PopDnType
		iArrParam(1) = "(select distinct mov_type from s_so_type_config) A, b_minor B"	
		iArrParam(2) = Trim(frm1.txtConDnType.value)					
		iArrParam(3) = ""											
		iArrParam(4) = "A.mov_type = B.minor_cd and B.major_cd = " & FilterVar("I0001", "''", "S") & ""	
		iArrParam(5) = frm1.txtConDnType.alt					
		
		iArrField(0) = "ED15" & Parent.gColSep & "A.mov_type"			
		iArrField(1) = "ED30" & Parent.gColSep & "B.minor_nm"			
    
	    iArrHeader(0) = frm1.txtConDnType.alt					
	    iArrHeader(1) = frm1.txtConDnTypeNm.alt					

		frm1.txtConDnType.focus

	End Select
	
	iArrParam(0) = iArrParam(5)										 
	
	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function


'========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	With frm1
		Select Case pvIntWhere

		Case C_PopDnType
			.txtConDnType.value = pvArrRet(0) 
			.txtConDnTypeNm.value = pvArrRet(1)  
			
		End Select
	End With

	SetConPopup = True		
	
End Function

'========================================
'	Name : SetQuerySpreadColor()
'	Description : 스프레트시트의 특정 컬럼의 배경색상을 변경 
'========================================
Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	Dim Spread
	
	Set Spread = frm1.vspdData
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)

	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)

		Spread.Col = -1
		Spread.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				Spread.BackColor = RGB(204,255,153) '연두 
			Case "2"
				Spread.BackColor = RGB(176,234,244) '하늘색 
			Case "3"
				Spread.BackColor = RGB(224,206,244) '연보라 
			Case "4"  
				Spread.BackColor = RGB(251,226,153) '연주황 
			Case "5" 
				Spread.BackColor = RGB(255,255,153) '연노랑 
		End Select
	Next

End Sub

'===============================================
Sub PopZAdoConfigGrid()

  If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
     Exit Sub
  End If

  Call OpenOrderBy("A")
  
End Sub

'========================================
Sub OpenOrderBy(ByVal pvPsdNo)
	Dim arrRet
	
	If lgIsOpenPop = True Then Exit Sub
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(pvPsdNo),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then												' Means that nothing is happened!!!
	   Exit Sub
	Else
	   Call ggoSpread.SaveXMLData(pvPsdNo,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Sub


'========================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
    
    If Frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If

    If Row <= 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort Col
            lgSortKey = 2
        Else
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 1
        End If
        Exit Sub
    End If
    

    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		

End Sub


'========================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
    If Row <= 0 Then

	

    End If
	
End Sub


'========================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData



End Sub


'========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
    


End Sub    


'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery()
    	End If
    End If
    
End Sub

'========================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConFromDt.Focus
	End If
End Sub

'========================================
Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConToDt.Focus
	End If
End Sub

'========================================
Sub txtConFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

'========================================
Sub txtConToDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출고현황조회(출하형태)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenGIDtlRef">출고상세현황</A></TD>
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
						<FIELDSET>
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>출하형태</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConDnType" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="출하형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopDnType)  ">
															<INPUT TYPE=TEXT NAME="txtConDnTypeNm" SIZE=20 tag="14" ALT="출하형태명"></TD>
									<TD CLASS="TD5" NOWRAP>출고일</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<script language =javascript src='./js/s4115qa3_OBJECT1_txtConFromDt.js'></script>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<script language =javascript src='./js/s4115qa3_OBJECT2_txtConToDt.js'></script>
											</TD>
										</TR>
									</TABLE>							        
							        </TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>출고여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoConfFlag" id="rdoConf" value="Y" tag = "11X">
											<label for="rdoConf">출고</label>&nbsp;&nbsp;
									</TD>
									<TD CLASS="TD5" NOWRAP></TD>
									<TD CLASS="TD6" NOWRAP></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/s4115qa3_vspdData_vspdData.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtConRdoFlag"		tag="14" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConFromDt"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConToDt"		tag="24" TABINDEX="-1">
				
<INPUT TYPE=HIDDEN NAME="txtHConDnType"		tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConRdoFlag"		tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHlgSelectListDT"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgTailList"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectList"		tag="24" TABINDEX="-1">				

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
