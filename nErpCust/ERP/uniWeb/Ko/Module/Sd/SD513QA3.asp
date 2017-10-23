<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : 영업관리 
'*  2. Function Name        : 영업실적관리 
'*  3. Program ID           : SD513QA3
'*  4. Program Name         : 부실채권조회(수금처)
'*  5. Program Desc         : 
'*  6. Comproxy List        :
'*  7. Modified date(First) : 2003/04/03
'*  8. Modified date(Last)  : 2003/06/09
'*  9. Modifier (First)     : kang su hwan
'* 10. Modifier (Last)      : Hwang Seong Bae
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
Const BIZ_PGM_ID 		= "SD513QB3.asp"

' Constant variables 
'========================================
Const C_MaxKey          = 20

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim lgIsOpenPop                                          

Dim lgStartRow
Dim lgEndRow

Const C_PopBizArea		=	0		'사업장								
Const C_PopSalesGrp		=	1		'영업그룹 
Const C_PopPayer		=	2		'수금처 
Const C_PopBpLevelType	=	3		'거래처등급							

Dim lgStrColorFlag

'========================================	
Sub InitVariables()

    lgPageNo	     = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                      'Indicates that current mode is Create mode
    lgSortKey        = 1

    Call SetToolBar("1100000000001111")
    
End Sub

'========================================
Sub SetDefaultVal()
	
End Sub							
	
'========================================
Sub LoadInfTB19029()

	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "S","NOCOOKIE","QA") %>                                 
	<% Call LoadBNumericFormatA("Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================
Sub InitSpreadSheet()

    Call SetZAdoSpreadSheet("SD513QA3","S","A", "V20030404", parent.C_SORT_DBAGENT,frm1.vspdData,C_MaxKey, "X", "X")
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
    
    frm1.txtConBizAreaCd.focus
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
       FncFind = True
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function FncExit()

    On Error Resume Next
    Err.Clear

	FncExit = True
End Function

'========================================
Function DbQuery() 

	Dim iStrVal

    On Error Resume Next
    Err.Clear

    DbQuery = False
    
	Call LayerShowHide(1)
    
    If lgIntFlgMode  <> parent.OPMD_UMODE Then
		With frm1
			iStrVal = BIZ_PGM_ID & "?BizAreaCd=" & Trim(.txtConBizAreaCd.value) & _
									"&SalesGrpCd=" & Trim(.txtConSalesGrpCd.value) & _
									"&PayerCd=" & Trim(.txtConPayerCd.value) & _
									"&BpLevelTypeCd=" & Trim(.txtConBpLevelTypeCd.value) & _
									"&lgSelectListDT=" & GetSQLSelectListDataType("A") & _
									"&lgTailList="     & MakeSQLGroupOrderByList("A") & _
									"&lgSelectList="   & EnCoding(GetSQLSelectList("A"))				
		End With	
    End If    

    lgStartRow = frm1.vspdData.MaxRows + 1										'포멧팅 적용하는 시작Row

    Call RunMyBizASP(MyBizASP, iStrVal)
		
    If Err.number = 0 Then
       DbQuery = True
    End If   

    Set gActiveElement = document.ActiveElement   

End Function

'========================================
Function DbQueryOk()												

    On Error Resume Next
    Err.Clear

    lgIntFlgMode     = parent.OPMD_UMODE
	
	Call SetQuerySpreadColor
    Call SetToolBar("1100000000011111")
	
	frm1.vspdData.Focus
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

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Select Case pvIntWhere
	
	'사업장 
	Case C_PopBizArea									
		iArrParam(1) = "B_BIZ_AREA"									' TABLE 명칭 
		iArrParam(2) = Trim(frm1.txtConBizAreaCd.value)				' Code Condition
		iArrParam(3) = ""											' Name Cindition
		iArrParam(4) = ""											' Where Condition
		iArrParam(5) = frm1.txtConBizAreaCd.alt						' TextBox 명칭 
		
		iArrField(0) = "ED15" & Parent.gColSep & "BIZ_AREA_CD"		' Field명(0)
		iArrField(1) = "ED30" & Parent.gColSep & "BIZ_AREA_NM"		' Field명(1)
    
	    iArrHeader(0) = frm1.txtConBizAreaCd.alt					' Header명(0)
	    iArrHeader(1) = frm1.txtConBizAreaNm.alt					' Header명(1)

		frm1.txtConBizAreaCd.focus 
	'영업그룹 
	Case C_PopSalesGrp	
		iArrParam(1) = "B_SALES_GRP"
		iArrParam(2) = Trim(frm1.txtConSalesGrpCd.value)
		iArrParam(3) = ""
		iArrParam(4) = "USAGE_FLAG = " & FilterVar("Y", "''", "S") & " "
		iArrParam(5) = frm1.txtConSalesGrpCd.alt
		
		iArrField(0) = "ED15" & Parent.gColSep & "SALES_GRP"
		iArrField(1) = "ED30" & Parent.gColSep & "SALES_GRP_NM"
    
	    iArrHeader(0) = frm1.txtConSalesGrpCd.alt
	    iArrHeader(1) = frm1.txtConSalesGrpNm.alt

		frm1.txtConSalesGrpCd.focus 
	'수금처		
	Case C_PopPayer
		iArrParam(1) = "B_BIZ_PARTNER_FTN PF, B_BIZ_PARTNER PA"
		iArrParam(2) = Trim(frm1.txtConPayerCd.value)
		iArrParam(3) = ""
		iArrParam(4) = "PF.USAGE_FLAG = " & FilterVar("Y", "''", "S") & "  AND PF.PARTNER_FTN= " & FilterVar("SPA", "''", "S") & "" _
						& "AND PA.BP_CD = PF.PARTNER_BP_CD AND PA.BP_TYPE <=" & FilterVar("CS", "''", "S") & ""
		iArrParam(5) = frm1.txtConPayerCd.alt
		
		iArrField(0) = "ED15" & Parent.gColSep & "PA.BP_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "PA.BP_NM"
    
	    iArrHeader(0) = frm1.txtConPayerCd.alt
	    iArrHeader(1) = frm1.txtConPayerNm.alt

		frm1.txtConPayerCd.focus

	'거래처등급 
	Case C_PopBpLevelType
		iArrParam(1) = "B_MINOR"
		iArrParam(2) = Trim(frm1.txtConBpLevelTypeCd.value)
		iArrParam(3) = ""
		iArrParam(4) = "MAJOR_CD = " & FilterVar("B9010", "''", "S") & ""
		iArrParam(5) = frm1.txtConBpLevelTypeCd.alt
		
		iArrField(0) = "ED15" & Parent.gColSep & "MINOR_CD"
		iArrField(1) = "ED30" & Parent.gColSep & "MINOR_NM"
    
	    iArrHeader(0) = frm1.txtConBpLevelTypeCd.alt
	    iArrHeader(1) = frm1.txtConBpLevelTypeNm.alt

		frm1.txtConBpLevelTypeCd.focus		

		
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
		Case C_PopBizArea
			.txtConBizAreaCd.value = pvArrRet(0) 
			.txtConBizAreaNm.value = pvArrRet(1)   			
			
		Case C_PopSalesGrp
			.txtConSalesGrpCd.value = pvArrRet(0)
			.txtConSalesGrpNm.value = pvArrRet(1)

		Case C_PopPayer
			.txtConPayerCd.value = pvArrRet(0) 
			.txtConPayerNm.value = pvArrRet(1) 

		Case C_PopBpLevelType
			frm1.txtConBpLevelTypeCd.value = pvArrRet(0) 
			frm1.txtConBpLevelTypeNm.value = pvArrRet(1)  			
			
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


</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>부실채권조회(수금처)</font></td>
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
									<TD CLASS="TD5" NOWRAP>사업장</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConBizAreaCd" SIZE=10 MAXLENGTH=10 tag="11NXXU" ALT="사업장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBizArea" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBizArea) ">
															<INPUT TYPE=TEXT NAME="txtConBizAreaNm" SIZE=20 tag="14" ALT="사업장명"></TD>
									<TD CLASS="TD5" NOWRAP>영업그룹</TD>
									<TD CLASS="TD6" NOWRAP>	<INPUT TYPE=TEXT NAME="txtConSalesGrpCd" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="영업그룹"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopSalesGrp) ">
															<INPUT TYPE=TEXT NAME="txtConSalesGrpNm" SIZE=20 tag="14" ALT="영업그룹명"></TD>
	                            </TR>		                            
								<TR>
									<TD CLASS="TD5" NOWRAP>수금처</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtConPayerCd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="수금처"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPayer" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopPayer) ">
														   <INPUT TYPE=TEXT NAME="txtConPayerNm" SIZE=20 tag="14" ALT="수금처명"></TD>
									<TD CLASS="TD5" NOWRAP>거래처등급</TD>
									<TD CLASS="TD6" NOWRAP> <INPUT TYPE=TEXT NAME="txtConBpLevelTypeCd" SIZE=10 MAXLENGTH=5 tag="11NXXU" ALT="거래처등급"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConBpLevelType" align=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenConPopUp(C_PopBpLevelType) ">
														    <INPUT TYPE=TEXT NAME="txtConBpLevelTypeNm" SIZE=20 tag="14" ALT="거래처등급명"></TD>
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
									<script language =javascript src='./js/sd513qa3_vspdData_vspdData.js'></script>
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

</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
