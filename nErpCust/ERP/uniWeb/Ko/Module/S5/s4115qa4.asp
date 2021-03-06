<%@ LANGUAGE="VBSCRIPT" %>
<%
'===================================================================
'*  1. Module Name          : Template1
'*  2. Function Name        : 
'*  3. Program ID           : 
'*  4. Program Name         : 
'*  5. Program Desc         : Ado query Sample with DBAgent(Sort)
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/04/01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      :
'* 11. Comment              :
'========================================
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
Option Explicit

Const BIZ_PGM_ID 		= "S4115QB4.asp" 
Const C_MaxKey          = 20				                         

<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================
'=										User-defind Variables
'========================================
Dim lgIsOpenPop                                          
Dim IsOpenPop  

Dim lgSaveRow 

<% 
   BaseDate     = GetSvrDate                                                        
%>  

Dim FirstDateOfDB 
Dim LastDateOfDB

Dim FromDateOfDB
Dim ToDateOfDB

FirstDateOfDB	= UNIConvDateAToB(UNIGetFirstDay("<%=BaseDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)
LastDateOfDB	= UNIConvDateAToB(UNIGetLastDay ("<%=BaseDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)

FromDateOfDB	= UNIConvDateAToB(UniDateAdd("m",-1,"<%=BaseDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)
ToDateOfDB		= UNIConvDateAToB(UniDateAdd("m", 0,"<%=BaseDate%>",parent.gServerDateFormat), parent.gServerDateFormat,parent.gDateFormat)

Dim lgStartRow
Dim lgEndRow
Dim lgStrColorFlag

Const C_PopSoldToParty	=	0										
Const C_PopDnType		=	1

'========================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                     
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0    
End Sub


'========================================
Sub SetDefaultVal()
	Frm1.txtConFromDt.Text		= cstr(FromDateOfDB)
	Frm1.txtConToDt.Text		= cstr(ToDateOfDB)
End Sub							
	
									
'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>

End Sub


'========================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("S4115QA4","S","A", "V20030529", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
    Call SetSpreadLock()       
End Sub


'========================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		'------ Developer Coding part (Start ) --------------------------------------------------------------		
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
   
    Frm1.txtConFromDt.Focus

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
    On Error Resume Next                                                        
    Err.Clear                                                                    

    DbQuery = False
    
	Call LayerShowHide(1)

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With frm1
        If lgIntFlgMode  <> parent.OPMD_UMODE Then								

			'원래는 Get방식이나 조건부가 많으면 POST방식으로 넘김 
			.txtHConFromDt.value		= .txtConFromDt.text
			.txtHConToDt.value			= .txtConToDt.text									
			.txtHConSoldToParty.value	= .txtConSoldToParty.value	
			.txtHConDnType.value		= .txtConDnType.value							
																
			If .rdoConf.checked = True Then				 									
				.txtHConRdoConfFlag.value = .rdoConf.value
			ElseIf .rdoNonConf.checked = True Then
				.txtHConRdoConfFlag.value = .rdoNonConf.value			
			End If
			
			.txtHlgSelectListDT.value	= GetSQLSelectListDataType("A") 
			.txtHlgTailList.value		= MakeSQLGroupOrderByList("A")
			.txtHlgSelectList.value		= EnCoding(GetSQLSelectList("A"))
        End If    
        
        .txtHlgPageNo.value	= lgPageNo
        
        lgStartRow = .vspdData.MaxRows + 1										
    End With
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
    lgIntFlgMode     = parent.OPMD_UMODE										 
    lgSaveRow        = 1
    
	'--------- Developer Coding Part (Start) ----------------------------------------------------------
	Call SetQuerySpreadColor
	'--------- Developer Coding Part (End) ----------------------------------------------------------
	
    Set gActiveElement = document.ActiveElement   
End Function


'========================================
Sub SetQuerySpreadColor()

	Dim iArrColor1, iArrColor2
	Dim iLoopCnt
	
	iArrColor1 = Split(lgStrColorFlag,Parent.gRowSep)
	
	For iLoopCnt=0 to ubound(iArrColor1,1) - 1
		iArrColor2 = Split(iArrColor1(iLoopCnt),Parent.gColSep)

		frm1.vspdData.Col = -1
		frm1.vspdData.Row =  iArrColor2(0)
		
		Select Case iArrColor2(1)
			Case "1"
				frm1.vspdData.BackColor = RGB(204,255,153) '연두 
			Case "2"
				frm1.vspdData.BackColor = RGB(176,234,244) '하늘색 
			Case "3"
				frm1.vspdData.BackColor = RGB(224,206,244) '연보라 
			Case "4"  
				frm1.vspdData.BackColor = RGB(251,226,153) '연주황 
			Case "5" 
				frm1.vspdData.BackColor = RGB(255,255,153) '연노랑 
		End Select
	Next

End Sub

'========================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere
	'납품처 
	Case C_PopSoldToParty									
		iArrParam(1) = "B_BIZ_PARTNER"									
		iArrParam(2) = Trim(frm1.txtConSoldToParty.value)				
		iArrParam(3) = ""											
		iArrParam(4) = "BP_TYPE IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"											
		iArrParam(5) = frm1.txtConSoldToParty.alt						
		
		iArrField(0) = "ED15" & Parent.gColSep & "BP_CD"		
		iArrField(1) = "ED30" & Parent.gColSep & "BP_NM"		
    
	    iArrHeader(0) = frm1.txtConSoldToParty.alt					
	    iArrHeader(1) = frm1.txtConSoldToPartyNm.alt					

		frm1.txtConSoldToParty.focus 	
		
	'출하형태 
	Case C_PopDnType	
		iArrParam(1) = "B_MINOR A, I_MOVETYPE_CONFIGURATION B"								
		iArrParam(2) = Trim(frm1.txtConDnType.value)			
		iArrParam(3) = ""											
		iArrParam(4) = "A.MINOR_CD=B.MOV_TYPE AND (B.TRNS_TYPE = " & FilterVar("DI", "''", "S") & " OR (B.TRNS_TYPE = " & FilterVar("ST", "''", "S") & " AND B.STCK_TYPE_FLAG_DEST = " & FilterVar("T", "''", "S") & " )) AND A.MAJOR_CD=" & FilterVar("I0001", "''", "S") & ""							
		iArrParam(5) = frm1.txtConDnType.alt					
		
		iArrField(0) = "ED15" & Parent.gColSep & "A.MINOR_CD"		
		iArrField(1) = "ED30" & Parent.gColSep & "A.MINOR_NM"		
    
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
		Case C_PopSoldToParty
			.txtConSoldToParty.value	= pvArrRet(0) 
			.txtConSoldToPartyNm.value	= pvArrRet(1)   			
			
		Case C_PopDnType
			.txtConDnType.value		= pvArrRet(0)
			.txtConDnTypeNm.value	= pvArrRet(1)			
			
		End Select
	End With

	SetConPopup = True		
	
End Function


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
Sub rdoConf_OnClick()
	lblTitle.innerHTML = "출고일"
	frm1.txtHConRdoConfFlag.value = frm1.rdoConf.value 

End Sub

Sub rdoNonConf_OnClick()
	lblTitle.innerHTML = "출고예정일"
	frm1.txtHConRdoConfFlag.value = frm1.rdoNonConf.value 
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
    If OldLeft <> NewLeft Then Exit Sub

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
	If Frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(Frm1.vspdData,NewTop) Then	    
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DbQuery
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


'========================================
Function OpenGIDtlRef()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(7)	
	
	On Error Resume Next
	
	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		frm1.txtConSoldToParty.focus
		Exit Function
	End IF
	
	If lgIsOpenPop = True Then Exit Function
	
	lgIsOpenPop = True
	
	With frm1
			
		.vspdData.Row	= .vspdData.ActiveRow
		
		arrParam(0)		=	""					'출하번호 

		.vspdData.Col	= GetKeyPos("A",1)
		
		If .vspdData.Text = "0" Then
		
			.vspdData.Col	= GetKeyPos("A",2) '납품처 
			arrParam(1)		= .vspdData.Text
	
			.vspdData.Col	= GetKeyPos("A",3) '납품처명 
			arrParam(2)		= .vspdData.Text
	
			.vspdData.Col	= GetKeyPos("A",4) '출하형태 
			arrParam(3)		= .vspdData.Text

			.vspdData.Col	= GetKeyPos("A",5) '출하형태명 
			arrParam(4)		= .vspdData.Text
		
		ElseIf .vspdData.Text = "1" Then	'소계일때 
			
			.vspdData.Col	= GetKeyPos("A",2) 
			arrParam(1)		= .vspdData.Text
	
			.vspdData.Col	= GetKeyPos("A",3) 
			arrParam(2)		= .vspdData.Text
			
			arrParam(3)		= ""
			arrParam(4)		= ""
			
		ElseIf .vspdData.Text = "2" Then	'합계일때 
			
			arrParam(1)		= ""
			arrParam(2)		= ""
			arrParam(3)		= ""
			arrParam(4)		= ""
						
		End If
		
		arrParam(5)		= .txtConFromDt.text
		arrParam(6)		= .txtConToDt.text									
																	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>출고현황조회(납품처)</font></td>
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
									<TD CLASS="TD5" NOWRAP>납품처</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSoldToParty" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="납품처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSoldToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSoldToParty) ">
															<INPUT TYPE=TEXT NAME="txtConSoldToPartyNm" SIZE=20 tag="14" ALT="납품처명"></TD>
									<TD CLASS="TD5" NOWRAP>출하형태</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConDnType" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="출하형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSoldToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopDnType) ">
															<INPUT TYPE=TEXT NAME="txtConDnTypeNm" SIZE=20 tag="14" ALT="출하형태명"></TD>
															
								</TR>
								<TR>
									<TD CLASS="TD5" id="lblTitle" NOWRAP>출고일</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<script language =javascript src='./js/s4115qa4_OBJECT1_txtConFromDt.js'></script>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<script language =javascript src='./js/s4115qa4_OBJECT2_txtConToDt.js'></script>
											</TD>
										</TR>										
									</TABLE>							        
							        </TD>	
							        <TD CLASS=TD5 NOWRAP>출고여부</TD>
									<TD CLASS=TD6 NOWRAP>										
										<input type=radio CLASS="RADIO" name="rdoConfFlag" id="rdoConf" value="Y" tag = "11X" checked>
											<label for="rdoConf">출고</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoConfFlag" id="rdoNonConf" value="N" tag = "11X">
											<label for="rdoNonConf">미출고</label>
									</TD>	
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
									<script language =javascript src='./js/s4115qa4_vspdData_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>    
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtHConFromDt"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConToDt"			tag="24" TABINDEX="-1">				
<INPUT TYPE=HIDDEN NAME="txtHConRdoConfFlag"    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConSoldToParty"    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConDnType"			tag="24" TABINDEX="-1">						
<INPUT TYPE=HIDDEN NAME="txtHlgPageNo"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectListDT"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgTailList"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectList"		tag="24" TABINDEX="-1">				
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
