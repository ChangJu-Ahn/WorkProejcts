<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 영업관리 
'*  2. Function Name        : 영업실적관리 
'*  3. Program ID           : SD513QA5
'*  4. Program Name         : 매출채권조회(매출채권일2)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/04/07
'*  8. Modified date(Last)  : 2003/06/10
'*  9. Modifier (First)     : kangsuhwan
'* 10. Modifier (Last)      : Hwang Seongbae
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

<Script Language="VBScript">
Option Explicit
	
' External ASP File
'========================================
Const BIZ_PGM_ID 		= "SD513QB5.asp"

' Constant variables 
'========================================
Const C_MaxKey          = 20

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim lgIsOpenPop                                          

Dim lgStrColorFlag

Dim ToDateOfDB

ToDateOfDB = UNIConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat,parent.gDateFormat)

Dim lgStartRow
Dim lgEndRow

Const C_PopSalesGrp1		=	0										
Const C_PopSalesGrp2		=	1
Const C_PopSoldToParty1		=	2
Const C_PopSoldToParty2		=	3										
Const C_PopBillType			=	4

'========================================	
Sub InitVariables()
    lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                      'Indicates that current mode is Create mode
    lgSortKey        = 1
    
    Call SetToolBar("1100000000001111")										
End Sub

'========================================
Sub SetDefaultVal()
	Frm1.txtConFromDt.Text		= UNIGetFirstDay(ToDateOfDB, Parent.gDateFormat)
	Frm1.txtConToDt.Text		= ToDateOfDB
End Sub							
	
'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================
Sub InitSpreadSheet()

    Call SetZAdoSpreadSheet("SD513QA5","S","A", "V20030425_", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
    Call SetSpreadLock()   
     
End Sub

'========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock 1 , -1
End Sub

'========================================
Sub Form_Load()

    Call LoadInfTB19029				                                           '⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables														
	Call SetDefaultVal	
	Call InitSpreadSheet()
    
    Call SetFocusToDocument("M")	
    frm1.txtConFromDt.focus
End Sub

'========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================
Function FncQuery() 

    On Error Resume Next                                                          
    Err.Clear                                                                     
	
    FncQuery = False

    If Not chkField(Document, "1") Then Exit Function

	If ValidDateCheck(frm1.txtConFromDt, frm1.txtConToDt) = False Then Exit Function
	
    Call ggoOper.ClearField(Document, "2")
    
    Call InitVariables
    
    If DbQuery Then Exit Function

    If Err.number = 0 Then FncQuery = True

    Set gActiveElement = document.ActiveElement  
End Function

'========================================
Function FncPrint()

    On Error Resume Next                                                          
    Err.Clear                                                                     

    FncPrint = False                                                              
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

	Call Parent.FncExport(parent.C_MULTI)

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

	Call Parent.FncFind(parent.C_MULTI, True)

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

    If Err.number = 0 Then
       FncExit = True                                                             
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQuery() 

	Dim strVal
	Dim strBillConfFlag
	Dim strExceptFlag
    On Error Resume Next                                                          
    Err.Clear                                                                     

    DbQuery = False
    
	Call LayerShowHide(1)
    
    With frm1

        strVal = BIZ_PGM_ID
        
		If frm1.rdoBillConfFlag(0).checked Then
			strBillConfFlag=""
		ElseIf frm1.rdoBillConfFlag(1).checked Then
			strBillConfFlag="Y"
		ElseIf frm1.rdoBillConfFlag(2).checked Then
			strBillConfFlag="N"
		End If
		
		If frm1.rdoExceptFlag(0).checked Then
			strExceptFlag=""
		ElseIf frm1.rdoExceptFlag(1).checked Then
			strExceptFlag="Y"
		ElseIf frm1.rdoExceptFlag(2).checked Then
			strExceptFlag="N"
		End If
		
        If lgIntFlgMode  <> parent.OPMD_UMODE Then									'☜: This means that it is first search
				frm1.txtHdnConFrDt.value		= Trim(frm1.txtConFromDt.Text)			<%'☜: 비지니스 처리 ASP의 상태 %>
				frm1.txtHdnConToDt.value		= Trim(frm1.txtConToDt.Text)	
				frm1.txtHdnConSalesGrpCd.value	= Trim(frm1.txtConSalesGrpCd1.value)	
				frm1.txtHdnConBpCd.value		= Trim(frm1.txtConBpCd1.value)	
				frm1.txtHdnConBillTypeCd.value	= Trim(frm1.txtConBillTypeCd.value)	
				frm1.txtHdnConBillConfFlag.value= Trim(strBillConfFlag)	
				frm1.txtHdnExceptFlag.value		= Trim(strExceptFlag)	
				frm1.txtHdnlgSelectListDT.value = GetSQLSelectListDataType("A")			 
				frm1.txtHdnlgTailList.value		= MakeSQLGroupOrderByList("A")
				frm1.txtHdnlgSelectList.value	= EnCoding(GetSQLSelectList("A"))				
				frm1.txtHdnPageNo.value			= lgPageNo		
		Else
				frm1.txtHdnConFrDt.value		= Trim(frm1.txtHConFromDt.Text)			<%'☜: 비지니스 처리 ASP의 상태 %>
				frm1.txtHdnConToDt.value		= Trim(frm1.txtHConToDt.Text)	
				frm1.txtHdnConSalesGrpCd.value	= Trim(frm1.txtHConSalesGrpCd1.value)	
				frm1.txtHdnConBpCd.value		= Trim(frm1.txtHConBpCd1.value)	
				frm1.txtHdnConBillTypeCd.value	= Trim(frm1.txtHConBillTypeCd.value)	
				frm1.txtHdnConBillConfFlag.value= Trim(frm1.rdoHBillConfFlag.value)	
				frm1.txtHdnExceptFlag.value		= Trim(frm1.rdoHExceptFlag.value)	
				frm1.txtHdnlgSelectListDT.value = GetSQLSelectListDataType("A")			 
				frm1.txtHdnlgTailList.value		= MakeSQLGroupOrderByList("A")
				frm1.txtHdnlgSelectList.value	= EnCoding(GetSQLSelectList("A"))				
				frm1.txtHdnPageNo.value			= lgPageNo		
        End If    
        lgStartRow = frm1.vspdData.MaxRows + 1										'포멧팅 적용하는 시작Row

		Call ExecMyBizASP(frm1, BIZ_PGM_ID)			
		
    End With

    If Err.number = 0 Then
       DbQuery = True																
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQueryOk()												

    On Error Resume Next                                                          
    Err.Clear                                                                     

    lgIntFlgMode     = parent.OPMD_UMODE										  '⊙: Indicates that current mode is Update mode

    Call SetQuerySpreadColor
    Call SetToolBar("1100000000011111")
	
	frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   

End Function

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

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case pvIntWhere
	'영업그룹#1
	Case C_PopSalesGrp1	
		iArrParam(1) = "B_SALES_GRP"								
		iArrParam(2) = Trim(frm1.txtConSalesGrpCd1.value)			
		iArrParam(3) = ""											
		iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "							
		iArrParam(5) = frm1.txtConSalesGrpCd1.alt					
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"		
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"		
    
	    iArrHeader(0) = frm1.txtConSalesGrpCd1.alt					
	    iArrHeader(1) = frm1.txtConSalesGrpNm1.alt					

		frm1.txtConSalesGrpCd1.focus 
	'지급처#1
	Case C_PopSoldToParty1											
		iArrParam(1) = "B_BIZ_PARTNER"								
		iArrParam(2) = Trim(frm1.txtConBpCd1.value)			
		iArrParam(3) = ""											
'		iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND BP_TYPE IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"	
		iArrParam(4) = "BP_TYPE IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"	
		iArrParam(5) = frm1.txtConBpCd1.alt					
		
		iArrField(0) = "ED15" & Parent.gColSep & "BP_CD"			
		iArrField(1) = "ED30" & Parent.gColSep & "BP_NM"			
    
	    iArrHeader(0) = frm1.txtConBpCd1.alt				
	    iArrHeader(1) = frm1.txtConBpNm1.alt				

		frm1.txtConBpCd1.focus
	'영업조직 
	Case C_PopBillType	
		iArrParam(1) = "S_BILL_TYPE_CONFIG"								
		iArrParam(2) = Trim(frm1.txtConBillTypeCd.value)			
		iArrParam(3) = ""											
		iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "							
		iArrParam(5) = frm1.txtConBillTypeCd.alt					
		
		iArrField(0) = "ED15" & Parent.gColSep & "BILL_TYPE"		
		iArrField(1) = "ED30" & Parent.gColSep & "BILL_TYPE_NM"		
    
	    iArrHeader(0) = frm1.txtConBillTypeCd.alt					
	    iArrHeader(1) = frm1.txtConBillTypeNm.alt					

		frm1.txtConBillTypeCd.focus 
		
	End Select
	
	iArrParam(0) = iArrParam(5)										' 팝업 명칭 
	
	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
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
		Case C_PopSalesGrp1
			.txtConSalesGrpCd1.value = pvArrRet(0) 
			.txtConSalesGrpNm1.value = pvArrRet(1)   			
			
		Case C_PopSoldToParty1
			.txtConBpCd1.value = pvArrRet(0)
			.txtConBpNm1.value = pvArrRet(1)

		Case C_PopBillType
			frm1.txtConBillTypeCd.value = pvArrRet(0) 
			frm1.txtConBillTypeNm.value = pvArrRet(1)  
		End Select
	End With

	SetConPopup = True		
	
End Function

'========================================
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

	If CheckRunningBizProcess Then  Exit Sub
    
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권조회(매출채권일2)</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="30" align=right></td>
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
									<TD CLASS="TD5" NOWRAP>매출채권일</TD>									
									<TD CLASS="TD6" NOWRAP>							        
									<TABLE CELLSPACING=0 CELLPADDING=0>
										<TR>
											<TD>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtConFromDt CLASS=FPDTYYYYMMDD tag="12X1" ALT="시작일" Title=FPDATETIME id=OBJECT1></OBJECT>');</SCRIPT>
											</TD>
											<TD>
											&nbsp;~&nbsp;
											</TD>
											<TD>
											<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtConToDt CLASS=FPDTYYYYMMDD tag="12X1" ALT="종료일" Title=FPDATETIME id=OBJECT2></OBJECT>');</SCRIPT>
											</TD>
										</TR>
									</TABLE>							        
							        </TD>
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSalesGrpCd1" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesGrp1) ">
															<INPUT TYPE=TEXT NAME="txtConSalesGrpNm1" SIZE=20 tag="14" ALT="영업그룹명"></TD>
	                            </TR>	
	                            <TR>
									<TD CLASS="TD5" NOWRAP>주문처</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBpCd1" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBpCd" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSoldToParty1) ">
															<INPUT TYPE=TEXT NAME="txtConBpNm1" SIZE=20 tag="14" ALT="주문처명"></TD>
									<TD CLASS="TD5" NOWRAP>매출채권형태</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBillTypeCd" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="매출채권형태"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBillType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBillType) ">
															<INPUT TYPE=TEXT NAME="txtConBillTypeNm" SIZE=20 tag="14" ALT="매출채권형태명"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>확정여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoBillConfFlag" id="rdoAll" value="A" tag = "11X" checked>
											<label for="rdoAll">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoBillConfFlag" id="rdoConf" value="S" tag = "11X">
											<label for="rdoConf">확정</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoBillConfFlag" id="rdoNonConf" value="D" tag = "11X">
											<label for="rdoNonConf">미확정</label>
									</TD>
									<TD CLASS=TD5 NOWRAP>예외여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoExceptFlag" id="rdoAll1" value="A" tag = "11X" checked>
											<label for="rdoAll1">전체</label>&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoExceptFlag" id="rdoExcept" value="Y" tag = "11X">
											<label for="rdoExcept">예외</label>&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoExceptFlag" id="rdoNormal" value="N" tag = "11X">
											<label for="rdoNormal">정상</label>
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
      <TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtHConFromDt"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConToDt"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConSalesGrpCd1"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConBpCd1"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConBillTypeCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="rdoHBillConfFlag"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="rdoHExceptFlag"		tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHdnConFrDt"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdnConToDt"			tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdnConSalesGrpCd"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdnConBpCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdnConBillTypeCd"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdnConBillConfFlag"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdnExceptFlag"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdnlgSelectListDT"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdnlgTailList"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdnlgSelectList"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHdnPageNo"			tag="24" TABINDEX="-1">

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
