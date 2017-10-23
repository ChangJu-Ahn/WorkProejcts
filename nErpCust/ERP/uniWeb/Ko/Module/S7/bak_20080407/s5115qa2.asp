<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
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
'=======================================================================================================
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

Const BIZ_PGM_ID 		= "S5115QB2.asp" 
Const C_MaxKey          = 20				                         

<!-- #Include file="../../inc/lgvariables.inc" -->	

'========================================================================================================
'=										User-defind Variables
'========================================================================================================
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
Const C_PopBizArea		=	0										
Const C_PopBillType		=	1

'========================================================================================================	
Sub InitVariables()
    lgStrPrevKey     = ""
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                     
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0    
End Sub


'========================================================================================================
Sub SetDefaultVal()
	Frm1.txtConFromDt.Text		= cstr(FirstDateOfDB)
	Frm1.txtConToDt.Text		= cstr(ToDateOfDB)
End Sub							
	
									
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("s5115qa2","S","A", "V20030523", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
    Call SetSpreadLock()       
End Sub


'========================================================================================================
Sub SetSpreadLock()
    With frm1
		.vspdData.ReDraw = False
		'------ Developer Coding part (Start ) --------------------------------------------------------------		
		ggoSpread.SpreadLock 1 , -1		
		'------ Developer Coding part (End   ) -------------------------------------------------------------- 
		.vspdData.ReDraw = True
    End With    
End Sub


'========================================================================================================
Sub Form_Load()
    Call LoadInfTB19029				                                           

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolBar("1100000000001111")	  
   
    Frm1.txtConFromDt.Focus
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
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
	

'========================================================================================================
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


'========================================================================================================
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


'========================================================================================================
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


'========================================================================================================
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


'========================================================================================================
Function DbQuery() 
    On Error Resume Next                                                        
    Err.Clear                                                                    

    DbQuery = False
    
	Call LayerShowHide(1)

    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    With frm1
        If lgIntFlgMode  <> parent.OPMD_UMODE Then								

			'원래는 Get방식이나 조건부가 많으면 POST방식으로 넘김 
			.txtHConFromDt.value	= .txtConFromDt.text
			.txtHConToDt.value		= .txtConToDt.text									
			.txtHConBizArea.value	= .txtConBizArea.value	
			.txtHConBillType.value	= .txtConBillType.value
												
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


'========================================================================================================
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

'========================================================================================================
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

'========================================================================================================
Function OpenConPopup(ByVal pvIntWhere)
	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere
	'사업장 
	Case C_PopBizArea									
		iArrParam(1) = "B_BIZ_AREA"									
		iArrParam(2) = Trim(frm1.txtConBizArea.value)				
		iArrParam(3) = ""											
		iArrParam(4) = ""											
		iArrParam(5) = frm1.txtConBizArea.alt						
		
		iArrField(0) = "ED15" & Parent.gColSep & "BIZ_AREA_CD"		
		iArrField(1) = "ED30" & Parent.gColSep & "BIZ_AREA_NM"		
    
	    iArrHeader(0) = frm1.txtConBizArea.alt					
	    iArrHeader(1) = frm1.txtConBizAreaNm.alt					

		frm1.txtConBizArea.focus 
	'매출채권형태 
	Case C_PopBillType	
		iArrParam(1) = "S_BILL_TYPE_CONFIG"								
		iArrParam(2) = Trim(frm1.txtConBillType.value)			
		iArrParam(3) = ""											
'		iArrParam(4) = "EXCEPT_FLAG <> " & FilterVar("Y", "''", "S") & "  AND USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND EXPORT_FLAG <> " & FilterVar("Y", "''", "S") & "  AND AS_FLAG = " & FilterVar("N", "''", "S") & " "							
		iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S")  & " "
		iArrParam(5) = frm1.txtConBillType.alt					
		
		iArrField(0) = "ED15" & Parent.gColSep & "BILL_TYPE"		
		iArrField(1) = "ED30" & Parent.gColSep & "BILL_TYPE_NM"		
    
	    iArrHeader(0) = frm1.txtConBillType.alt					
	    iArrHeader(1) = frm1.txtConBillTypeNm.alt					

		frm1.txtConBillType.focus 		
		
	End Select
	
	iArrParam(0) = iArrParam(5)										<%' 팝업 명칭 %> 
	
	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If iArrRet(0) = "" Then
		Exit Function
	Else
		OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	End If	

End Function


'========================================================================================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)
	SetConPopup = False

	With frm1
		Select Case pvIntWhere
		Case C_PopBizArea
			.txtConBizArea.value	= pvArrRet(0) 
			.txtConBizAreaNm.value	= pvArrRet(1)   			
			
		Case C_PopBillType
			.txtConBillType.value	= pvArrRet(0)
			.txtConBillTypeNm.value = pvArrRet(1)			
			
		End Select
	End With

	SetConPopup = True		
	
End Function


'========매출채권상세====================================================================================
Function OpenBillDtl()
	Dim iCalledAspName
	Dim arrRet
	Dim arrParam(8)	
	Dim iCnt
	
	On Error Resume Next

	If lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call DisplayMsgBox("900002", "X", "X", "X")
		frm1.txtConBpCd.focus
		Exit Function
	End IF
	
	If IsOpenPop = True Then Exit Function
	
	For iCnt=0 to 8
		arrParam(iCnt) = ""
	Next

	arrParam(0) = frm1.txtHConFromDt.value			
	arrParam(1) = frm1.txtHConToDt.value
	arrParam(5) = frm1.txtHConRdoConfFlag.value 	 

	frm1.vspdData.row = frm1.vspddata.activerow
	
	frm1.vspdData.Col = GetKeyPos("A",1)			
	If frm1.vspdData.Text = "0" Then	
		
		frm1.vspdData.Col = GetKeyPos("A",2)			' 사업장 
		arrParam(2) = frm1.vspdData.Text
		
		frm1.vspdData.Col = GetKeyPos("A",3)			' 사업장명 
		arrParam(6) = frm1.vspdData.Text	
		
		frm1.vspdData.Col = GetKeyPos("A",4)			' 매출채권형태 
		arrParam(4) = frm1.vspdData.Text
		
		frm1.vspdData.Col = GetKeyPos("A",5)			' 매출채권형태명 
		arrParam(8) = frm1.vspdData.Text	

	ElseIf frm1.vspdData.Text = "1" Then				' 소계일때		
				
		frm1.vspdData.Col = GetKeyPos("A",2)			' 사업장 
		arrParam(2) = frm1.vspdData.Text
		
		frm1.vspdData.Col = GetKeyPos("A",3)			' 사업장명 
		arrParam(6) = frm1.vspdData.Text
	
	End If	

	IsOpenPop = True
   
	iCalledAspName = AskPRAspName("s5116pa5")	
	if Trim(iCalledAspName) = "" then
		IntRetCD = DisplayMsgBox("900040",parent.VB_INFORMATION, "s5116pa5", "x")
		IsOpenPop = False
		exit Function
	end if

	arrRet = window.showModalDialog(iCalledAspName,Array(window.parent,arrParam),"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

End Function


'==================================================================================
Sub PopZAdoConfigGrid()
  If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
     Exit Sub
  End If

  Call OpenOrderBy("A")  
End Sub


'========================================================================================================
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


'========================================================================================================
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
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    Call SetSpreadColumnValue("A",frm1.vspdData,Col,Row)		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)	    
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If	
End Sub


'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub    

'========================================================================================================
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


'========================================================================================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConFromDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConFromDt.Focus
	End If
End Sub


'========================================================================================================
Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
       Frm1.txtConToDt.Action = 7
       Call SetFocusToDocument("M")	
       Frm1.txtConToDt.Focus
	End If
End Sub

'========================================================================================================
Sub txtConFromDt_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery()
	End If   
End Sub


'========================================================================================================
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권현황조회(사업장)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenBillDtl">매출채권상세</A></TD>
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
									<TD CLASS="TD5" NOWRAP>매출채권일자</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtConFromDt CLASS=FPDTYYYYMMDD tag="12X1" ALT="조회기간시작" Title=FPDATETIME id=OBJECT1></OBJECT>');</SCRIPT>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtConToDt CLASS=FPDTYYYYMMDD tag="11X1" ALT="조회기간끝" Title=FPDATETIME id=OBJECT2></OBJECT>');</SCRIPT>
											</TD>
										</TR>										
									</TABLE>							        
							        </TD>	
							        <TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBizArea" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBizArea) ">
															<INPUT TYPE=TEXT NAME="txtConBizAreaNm" SIZE=20 tag="14" ALT="사업장명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>매출채권형태</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBillType" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="매출채권형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBillType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBillType) ">
															<INPUT TYPE=TEXT NAME="txtConBillTypeNm" SIZE=20 tag="14" ALT="매출채권형태명"></TD>							
							        <TD CLASS=TD5 NOWRAP>확정여부</TD>
									<TD CLASS=TD6 NOWRAP>										
										<input type=radio CLASS="RADIO" name="rdoConfFlag" id="rdoConf" value="Y" tag = "11X" checked>
											<label for="rdoConf">확정</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoConfFlag" id="rdoNonConf" value="N" tag = "11X">
											<label for="rdoNonConf">미확정</label>
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
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vspdData> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
<INPUT TYPE=HIDDEN NAME="txtHConBizArea"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConBillType"		tag="24" TABINDEX="-1">
			
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
