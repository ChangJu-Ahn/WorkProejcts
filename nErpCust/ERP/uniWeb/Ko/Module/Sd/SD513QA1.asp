<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 영업실적관리 
'*  3. Program ID           : SD513QA1
'*  4. Program Name         : 매출채권조회(매출채권일1)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/04/25
'*  8. Modified date(Last)  : 2003/06/09
'*  9. Modifier (First)     : 
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
Const BIZ_PGM_ID 		= "SD513QB1.asp"

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

Const C_PopBizArea		=	0
Const C_PopSalesGrp		=	1
Const C_PopSalesType	=	2
Const C_PopSoldToParty	=	3
Const C_PopPayer		=	4										
Const C_PopBillToParty	=	5

'========================================
Sub InitVariables()

    lgPageNo		 = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                      'Indicates that current mode is Create mode
    lgSortKey        = 1

    Call SetToolBar("1100000000001111")										
    
End Sub

'========================================
Sub SetDefaultVal()

	With frm1
		.txtConFromDt.Text	= UNIGetFirstDay(ToDateOfDB, Parent.gDateFormat)
		.txtConToDt.Text	= ToDateOfDB
		.rdoPostfiFlagAll.checked = True
		.txtConRdoBillConfFlag.value = .rdoPostfiFlagAll.value   
		.rdoExceptFlagAll.checked = True
		.txtConRdoBillConfFlag1.value = .rdoExceptFlagAll.value   
	End With
	
End Sub							
	
'========================================
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================
Sub InitSpreadSheet()

    Call SetZAdoSpreadSheet("SD513QA1","S","A", "V20030408", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
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

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncQuery = False                                                              '⊙: Processing is NG
    
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

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
	Call Parent.FncPrint()                                                        '☜: Protect system from crashing

    If Err.number = 0 Then
       FncPrint = True                                                            '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(parent.C_MULTI)
    
    If Err.number = 0 Then
       FncExcel = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(parent.C_MULTI, True)
    
    If Err.number = 0 Then
       FncFind = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncExit()

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG

    If Err.number = 0 Then
       FncExit = True                                                             '⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQuery() 

	Dim strVal

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbQuery = False
    
	Call LayerShowHide(1)

    strVal = BIZ_PGM_ID
        
    If lgIntFlgMode  <> parent.OPMD_UMODE Then									'☜: This means that it is first search
			
		'원래는 Get방식이나 조건부가 많으면 POST방식으로 넘김 
		With frm1

			.txtHConFromDt.value	= .txtConFromDt.text
			.txtHConToDt.value	= .txtConToDt.text
			
			.txtHConBizAreaCd.value		= .txtConBizAreaCd.value
			.txtHConSalesGrpCd.value	= .txtConSalesGrpCd.value
			.txtHConSalesTypeCd.value	= .txtConSalesTypeCd.value
			.txtHConSoldToPartyCd.value	= .txtConSoldToPartyCd.value
			.txtHConPayerCd.value		= .txtConPayerCd.value
			.txtHConBillToPartyCd.value	= .txtConBillToPartyCd.value
			.txtHConRdoBillConfFlag.value	= .txtConRdoBillConfFlag.value
			.txtHConRdoBillConfFlag1.value	= .txtConRdoBillConfFlag1.value
							
			.txtHlgSelectListDT.value	= GetSQLSelectListDataType("A") 
			.txtHlgTailList.value		= MakeSQLGroupOrderByList("A")
			.txtHlgSelectList.value		= EnCoding(GetSQLSelectList("A"))
		End With	
    End If    

    lgStartRow = frm1.vspdData.MaxRows + 1										'포멧팅 적용하는 시작Row

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)

    If Err.number = 0 Then
       DbQuery = True																'⊙: Processing is OK
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQueryOk()												

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    lgIntFlgMode     = parent.OPMD_UMODE										  '⊙: Indicates that current mode is Update mode
    
    Call SetQuerySpreadColor
    Call SetToolBar("1100000000011111")
	
	frm1.vspdData.Focus
    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case pvIntWhere
	'사업장 
	Case C_PopBizArea									
		iArrParam(1) = "B_BIZ_AREA"									<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtConBizAreaCd.value)				<%' Code Condition%>
		iArrParam(3) = ""											<%' Name Cindition%>
		iArrParam(4) = ""											<%' Where Condition%>
		iArrParam(5) = frm1.txtConBizAreaCd.alt						<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & Parent.gColSep & "BIZ_AREA_CD"		<%' Field명(0)%>
		iArrField(1) = "ED30" & Parent.gColSep & "BIZ_AREA_NM"		<%' Field명(1)%>
    
	    iArrHeader(0) = frm1.txtConBizAreaCd.alt					<%' Header명(0)%>
	    iArrHeader(1) = frm1.txtConBizAreaNm.alt					<%' Header명(1)%>

		frm1.txtConBizAreaCd.focus 
	'영업그룹 
	Case C_PopSalesGrp	
		iArrParam(1) = "B_SALES_GRP"								<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtConSalesGrpCd.value)			<%' Code Condition%>
		iArrParam(3) = ""											<%' Name Cindition%>
		iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "							<%' Where Condition%>
		iArrParam(5) = frm1.txtConSalesGrpCd.alt					<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"		<%' Field명(0)%>
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"		<%' Field명(1)%>
    
	    iArrHeader(0) = frm1.txtConSalesGrpCd.alt					<%' Header명(0)%>
	    iArrHeader(1) = frm1.txtConSalesGrpNm.alt					<%' Header명(1)%>

		frm1.txtConSalesGrpCd.focus 
	'판매유형 
	Case C_PopSalesType
		iArrParam(1) = "B_MINOR"									<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtConSalesTypeCd.value)			<%' Code Condition%>
		iArrParam(3) = ""											<%' Name Cindition%>
		iArrParam(4) = "MAJOR_CD = " & FilterVar("S0001", "''", "S") & ""							<%' Where Condition%>
		iArrParam(5) = frm1.txtConSalesTypeCd.alt					<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"			<%' Field명(0)%>
		iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"			<%' Field명(1)%>
    
	    iArrHeader(0) = frm1.txtConSalesTypeCd.alt					<%' Header명(0)%>
	    iArrHeader(1) = frm1.txtConSalesTypeNm.alt					<%' Header명(1)%>

		frm1.txtConSalesTypeCd.focus
	'주문처 
	Case C_PopSoldToParty											
		iArrParam(1) = "B_BIZ_PARTNER"								<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtConSoldToPartyCd.value)			<%' Code Condition%>
		iArrParam(3) = ""											<%' Name Cindition%>
'		iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND BP_TYPE IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"	<%' Where Condition%>
		iArrParam(4) = "BP_TYPE IN (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"	<%' Where Condition%>
		iArrParam(5) = frm1.txtConSoldToPartyCd.alt					<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & Parent.gColSep & "BP_CD"			<%' Field명(0)%>
		iArrField(1) = "ED30" & Parent.gColSep & "BP_NM"			<%' Field명(1)%>
    
	    iArrHeader(0) = frm1.txtConSoldToPartyCd.alt				<%' Header명(0)%>
	    iArrHeader(1) = frm1.txtConSoldToPartyNm.alt				<%' Header명(1)%>

		frm1.txtConSoldToPartyCd.focus
	'수금처		
	Case C_PopPayer
		iArrParam(1) = "B_BIZ_PARTNER_FTN PF, B_BIZ_PARTNER PA"						<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtConPayerCd.value)								<%' Code Condition%>
		iArrParam(3) = ""															<%' Name Cindition%>
'		iArrParam(4) = "PF.USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND PF.PARTNER_FTN= " & FilterVar("SPA", "''", "S") & "" _
'						& "AND PA.BP_CD = PF.PARTNER_BP_CD AND PA.BP_TYPE <=" & FilterVar("CS", "''", "S") & ""	<%' Where Condition%>

		iArrParam(4) = "PF.PARTNER_FTN= " & FilterVar("SPA", "''", "S") & "" _
						& "AND PA.BP_CD = PF.PARTNER_BP_CD AND PA.BP_TYPE <=" & FilterVar("CS", "''", "S") & ""	<%' Where Condition%>
		iArrParam(5) = frm1.txtConPayerCd.alt										<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & Parent.gColSep & "PA.BP_CD"			<%' Field명(0)%>
		iArrField(1) = "ED30" & Parent.gColSep & "PA.BP_NM"			<%' Field명(1)%>
    
	    iArrHeader(0) = frm1.txtConPayerCd.alt						<%' Header명(0)%>
	    iArrHeader(1) = frm1.txtConPayerNm.alt						<%' Header명(1)%>

		frm1.txtConPayerCd.focus
	'발행처 
	Case C_PopBillToParty
		iArrParam(1) = "B_BIZ_PARTNER_FTN PF, B_BIZ_PARTNER PA"						<%' TABLE 명칭 %>
		iArrParam(2) = Trim(frm1.txtConBillToPartyCd.value)							<%' Code Condition%>
		iArrParam(3) = ""															<%' Name Cindition%>
'		iArrParam(4) = "PF.USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND PF.PARTNER_FTN = " & FilterVar("SBI", "''", "S") & "" _
'						& "AND PA.BP_CD = PF.PARTNER_BP_CD AND PA.BP_TYPE <= " & FilterVar("CS", "''", "S") & ""	<%' Where Condition%>
		iArrParam(4) = "PF.PARTNER_FTN = " & FilterVar("SBI", "''", "S") & "" _
						& "AND PA.BP_CD = PF.PARTNER_BP_CD AND PA.BP_TYPE <= " & FilterVar("CS", "''", "S") & ""	<%' Where Condition%>
		iArrParam(5) = frm1.txtConBillToPartyCd.alt									<%' TextBox 명칭 %>
		
		iArrField(0) = "ED15" & Parent.gColSep & "PA.BP_CD"			<%' Field명(0)%>
		iArrField(1) = "ED30" & Parent.gColSep & "PA.BP_NM"			<%' Field명(1)%>
    
	    iArrHeader(0) = frm1.txtConBillToPartyCd.alt				<%' Header명(0)%>
	    iArrHeader(1) = frm1.txtConBillToPartyNm.alt				<%' Header명(1)%>

		frm1.txtConBillToPartyCd.focus

	End Select
	
	iArrParam(0) = iArrParam(5)										<%' 팝업 명칭 %> 
	
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

		Case C_PopBizArea
			.txtConBizAreaCd.value = pvArrRet(0) 
			.txtConBizAreaNm.value = pvArrRet(1)   			
		Case C_PopSalesGrp
			.txtConSalesGrpCd.value = pvArrRet(0)
			.txtConSalesGrpNm.value = pvArrRet(1)
		Case C_PopSalesType
			.txtConSalesTypeCd.value = pvArrRet(0) 
			.txtConSalesTypeNm.value = pvArrRet(1)  
		Case C_PopSoldToParty
			frm1.txtConSoldToPartyCd.value = pvArrRet(0) 
			frm1.txtConSoldToPartyNm.value = pvArrRet(1)   
		Case C_PopPayer
			frm1.txtConPayerCd.value = pvArrRet(0) 
			frm1.txtConPayerNm.value = pvArrRet(1) 
		Case C_PopBillToParty
			frm1.txtConBillToPartyCd.value = pvArrRet(0) 
			frm1.txtConBillToPartyNm.value = pvArrRet(1)  

		End Select
	End With

	SetConPopup = True		
	
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
Sub rdoPostfiFlagAll_OnClick()
	frm1.txtConRdoBillConfFlag.value = frm1.rdoPostfiFlagAll.value 
End Sub

'========================================
Sub rdoPostfiFlagNo_OnClick()
	frm1.txtConRdoBillConfFlag.value = frm1.rdoPostfiFlagNo.value 
End Sub

'========================================
Sub rdoPostfiFlagYes_OnClick()
	frm1.txtConRdoBillConfFlag.value = frm1.rdoPostfiFlagYes.value 
End Sub

'========================================
Sub rdoExceptFlagAll_OnClick()
	frm1.txtConRdoBillConfFlag1.value = frm1.rdoExceptFlagAll.value 
End Sub

'========================================
Sub rdoExceptFlagNo_OnClick()
	frm1.txtConRdoBillConfFlag1.value = frm1.rdoExceptFlagNo.value 
End Sub

'========================================
Sub rdoExceptFlagYes_OnClick()
	frm1.txtConRdoBillConfFlag1.value = frm1.rdoExceptFlagYes.value 
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권조회(매출채권일1)</font></td>
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
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSalesGrpCd" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesGrp) ">
															<INPUT TYPE=TEXT NAME="txtConSalesGrpNm" SIZE=20 tag="14" ALT="영업그룹명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>판매유형</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSalesTypeCd" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="판매유형"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesType)  ">
															<INPUT TYPE=TEXT NAME="txtConSalesTypeNm" SIZE=20 tag="14" ALT="판매유형명"></TD>
									<TD CLASS="TD5" NOWRAP>주문처</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSoldToPartyCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="주문처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSoldToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSoldToParty) ">
															<INPUT TYPE=TEXT NAME="txtConSoldToPartyNm" SIZE=20 tag="14" ALT="주문처명"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>수금처</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConPayerCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="수금처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPayer" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopPayer) ">
															<INPUT TYPE=TEXT NAME="txtConPayerNm" SIZE=20 tag="14" ALT="수금처명"></TD>
									<TD CLASS="TD5" NOWRAP>발행처</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBillToPartyCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="발행처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBillToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBillToParty) ">
															<INPUT TYPE=TEXT NAME="txtConBillToPartyNm" SIZE=20 tag="14" ALT="발행처명"></TD>
								</TR>
								<TR>
							        </TD>
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBizArea) ">
															<INPUT TYPE=TEXT NAME="txtConBizAreaNm" SIZE=20 tag="14" ALT="사업장명"></TD>
									<TD CLASS=TD5 NOWRAP>확정여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoPostfiFlag" id="rdoPostfiFlagAll" value="%" tag = "11" checked>
											<label for="rdoPostfiFlagAll">전체</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoPostfiFlag" id="rdoPostfiFlagYes" value="Y" tag = "11">
											<label for="rdoPostfiFlagYes">확정</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoPostfiFlag" id="rdoPostfiFlagNo" value="N" tag = "11">
											<label for="rdoPostfiFlagNo">미확정</label></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>예외여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoExceptFlag" id="rdoExceptFlagAll" value="%" tag = "11" checked>
											<label for="rdoExceptFlagAll">전체</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoExceptFlag" id="rdoExceptFlagYes" value="Y" tag = "11">
											<label for="rdoExceptFlagYes">예외</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoExceptFlag" id="rdoExceptFlagNo" value="N" tag = "11">
											<label for="rdoExceptFlagNo">정상</label></TD>								
									<TD CLASS="TD5" NOWRAP>&nbsp;</TD>
									<TD CLASS="TD6" NOWRAP>&nbsp;</TD>
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

<INPUT TYPE=HIDDEN NAME="txtConRdoBillConfFlag" tag="14" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtConRdoBillConfFlag1" tag="14" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConFromDt"   tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConToDt"		tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConBizAreaCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConSalesGrpCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConSalesTypeCd"	tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConSoldToPartyCd"  tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConPayerCd"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConBillToPartyCd"	tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHConRdoBillConfFlag"    tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConRdoBillConfFlag1"    tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHlgSelectListDT"	tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgTailList"		tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHlgSelectList"		tag="24" TABINDEX="-1">				

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
