<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5111QA2
'*  4. Program Name         : 영업조직별 월 매출실적조회 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/08/09
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : 
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
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s5111qb2.asp"

' Constant variables 
'========================================
Const C_MaxKey          = 2                                           

Const C_PopSalesOrg	= 1

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim IsOpenPop  

Dim lgBlnOpenedFlag
Dim	lgBlnSalesOrgChg

Dim iDBSYSDate
Dim EndDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'========================================
Function InitVariables()
	lgPageNo         = ""
    lgIntFlgMode     = Parent.OPMD_CMODE                          
    lgSortKey        = 1   

    Call SetToolbar("11000000000011")										
	lgBlnSalesOrgChg = False								' 주문처 변경여부 
End Function

'========================================
Sub SetDefaultVal()
	Dim	iStrYear
	
	iStrYear = Left(UniConvDateToYYYYMM(EndDate, Parent.gDateFormat, "-"), 4)
	With frm1
		.cboYear.value = Left(iStrYear, 4)
		.cboQueryData.value = "B"
		.cboSalesOrgLvl.value = 1
		.cboYear.focus
	End With
	lgBlnFlgChgValue = False
End Sub

'==========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "QA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S5111QA2","S","A","V20021107", Parent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 
End Sub

'========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock 1 , -1
End Sub	

'========================================
Function OpenConPopup(ByVal pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	OpenConPopup = False
	
	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case pvIntWhere
	Case C_PopSalesOrg												
		iArrParam(1) = "B_SALES_ORG "						
		iArrParam(2) = Trim(frm1.txtSalesOrg.value)			
		iArrParam(3) = ""									
		iArrParam(4) = "LVL = " & frm1.cboSalesOrgLvl.value	
		iArrParam(5) = "영업조직"							
			
		iArrField(0) = "SALES_ORG"							
		iArrField(1) = "SALES_ORG_NM"								    
		iArrHeader(0) = "영업조직"						
		iArrHeader(1) = "영업조직명"						

	End Select
 
	iArrParam(0) = iArrParam(5)							

	iArrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	frm1.txtSalesOrg.focus
	
	If iArrRet(0) <> "" Then OpenConPopup = SetConPopup(iArrRet,pvIntWhere)
	
End Function

'========================================
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenOrderByPopup("A")
End Sub

'========================================
Function OpenOrderByPopup(ByVal pSpdNo)
	Dim arrRet
	
	On Error Resume Next 
	
	If IsOpenPop = True Then Exit Function
	IsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'========================================
Function SetConPopup(Byval pvArrRet,ByVal pvIntWhere)

	SetConPopup = False

	Select Case pvIntWhere
	Case C_PopSalesOrg
		frm1.txtSalesOrg.value = pvArrRet(0) 
		frm1.txtSalesOrgNm.value = pvArrRet(1)   
	End Select

	SetConPopup = True

End Function

'========================================
Sub InitComboBox()
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Call CommonQueryRs("n_year ","s_year","usage = " & FilterVar("Y", "''", "S") & "  ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboYear, lgF0, lgF0, Chr(11))
    Call CommonQueryRs("minor_cd, minor_nm ","b_minor","major_cd = " & FilterVar("S0016", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboSalesOrgLvl, lgF0, lgF1, Chr(11))
	Call SetCombo(frm1.cboQueryData, "B", "매출")
	Call SetCombo(frm1.cboQueryData, "T", "세금계산서")
End Sub

'========================================
Sub Form_Load()
    Call LoadInfTB19029											  
    
    'Html에서 tag 숫자가 1과 2로 시작하는 부분 각각Format
    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	
	Call ggoOper.LockField(Document, "N")                         
    
    Call InitComboBox()
	Call InitVariables											  
	Call SetDefaultVal	
	Call InitSpreadSheet()
	
	lgBlnOpenedFlag = True
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Function txtSalesOrg_OnChange()
	Dim iStrCode
	
	With frm1
		iStrCode = Trim(.txtSalesOrg.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", .cboSalesOrgLvl.value, "" & FilterVar("SO", "''", "S") & "", C_PopSalesOrg) Then
				.txtSalesOrg.value = ""
				.txtSalesOrgNm.value = ""
				.txtSalesOrg.focus
			Else
				.cboYear.focus
			End If
			txtSalesOrg_OnChange = False
		Else
			.txtSalesOrgNm.value = ""
		End If
	End With
	lgBlnSalesOrgChg = False
End Function

'========================================
Function cboSalesOrgLvl_OnChange()
	With frm1
		.txtSalesOrg.value = ""
		.txtSalesOrgNm.value = ""
		If .cboSalesOrgLvl.value = "" Then
			ggoOper.SetReqAttr .txtSalesOrg , "Q"
			.btnSalesOrg.disabled = True
		Else
			ggoOper.SetReqAttr .txtSalesOrg , "D"
			.btnSalesOrg.disabled = False
		End If
	End With
End Function

'========================================
Function txtSalesOrg_OnKeyDown()
	lgBlnSalesOrgChg = True
	lgBlnFlgChgValue = True
End Function

'========================================
Function ChkValidityQueryCon()
	Dim iStrCode

	ChkValidityQueryCon = True
	If lgBlnSalesOrgChg Then
		iStrCode = Trim(frm1.txtSalesOrg.value)
		If iStrCode <> "" Then
			iStrCode = " " & FilterVar(iStrCode, "''", "S") & ""
			If Not GetCodeName(iStrCode, "default", "default", "default", frm1.cboSalesOrgLvl.value, "" & FilterVar("SO", "''", "S") & "", C_PopSalesOrg) Then
				frm1.txtSalesOrg.value = ""
				frm1.txtSalesOrgNm.value = ""
'				Call DisplayMsgBox("970000", "X", frm1.txtSalesOrg.alt, "X")
				frm1.txtSalesOrg.focus
				ChkValidityQueryCon = False
				Exit Function
			End If
		Else
			frm1.txtSalesOrgNm.value = ""
		End If
		lgBlnSalesOrgChg	= False
	End If

End Function

'========================================
Function GetCodeName(ByVal pvStrArg1, ByVal pvStrArg2, ByVal pvStrArg3, ByVal pvStrArg4, ByVal pvIntArg5, ByVal pvStrFlag, ByVal pvIntWhere)

	Dim iStrSelectList, iStrFromList, iStrWhereList
	Dim iStrRs
	Dim iArrRs(2), iArrTemp
	
	GetCodeName = False

	iStrSelectList = " * "
	iStrFromList = " dbo.ufn_s_GetCodeName (" & pvStrArg1 & ", " & pvStrArg2 & ", " & pvStrArg3 & ", " & pvStrArg4 & ", " & pvIntArg5 & ", " & pvStrFlag & ") "
	iStrWhereList = ""
	
	Err.Clear
	
	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList,iStrRs) Then
		iArrTemp = Split(iStrRs, Chr(11))
		iArrRs(0) = iArrTemp(1)
		iArrRs(1) = iArrTemp(2)
		iArrRs(2) = iArrTemp(3)
		GetCodeName = SetConPopup(iArrRs, pvIntWhere)
	Else
		' 관련 Popup Display
		If lgBlnOpenedFlag Then GetCodeName = OpenConPopup(pvIntWhere)
	End if
End Function

'==========================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	
	Call SetPopupMenuItemInf("00000000001")
	
	gMouseClickStatus = "SPC"

    ggoSpread.Source = frm1.vspdData
    Set gActiveSpdSheet = frm1.vspdData
        
    If Row = 0 Then
		frm1.vspdData.ReDraw = False
		frm1.vspdData.OperationMode = 0

        If lgSortKey = 1 Then
            ggoSpread.SSSort Col, lgSortKey
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
		frm1.vspdData.ReDraw = True
	Else
		frm1.vspdData.ReDraw = False		
		frm1.vspdData.OperationMode = 3
		frm1.vspdData.ReDraw = True
    End If
  
End Sub

'========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub    

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then
	    Exit Sub
	End If

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgPageNo <> "" Then
			If DbQuery = False Then
				Exit Sub
			End if
		End If
	End If

End Sub

'========================================
Sub vspdData_Keypress(KeyAscii)
	If KeyAscii = 13	Then Call MainQuery()
End Sub

'========================================
Function FncQuery() 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               
   
    If Not chkField(Document, "1") Then Exit Function

	' 조회조건 유효값 check
	If 	lgBlnFlgChgValue Then
		If Not ChkValidityQueryCon Then	Exit Function
	End If

    Call ggoOper.ClearField(Document, "2")	         						
	
    Call InitVariables
    
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
End Function

'========================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_MULTI)
End Function

'========================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI , False)
End Function

'========================================
Function FncSplitColumn()
    
    Dim ACol
    Dim ARow
    Dim iRet
    Dim iColumnLimit
    
    iColumnLimit  = frm1.vspddata.maxcols
   
    ACol = Frm1.vspdData.ActiveCol
    ARow = Frm1.vspdData.ActiveRow

    If ACol > iColumnLimit Then
		'◎ Frm1없으면 frm1삭제 
		Frm1.vspdData.Col = iColumnLimit	:	Frm1.vspdData.Row = 0
		iRet = DisplayMsgBox("900030", "X", Trim(frm1.vspdData.Text), "X")
		Exit Function
    End If   

    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_NONE
    
    ggoSpread.Source = Frm1.vspdData
    
    ggoSpread.SSSetSplit(ACol)    
    
    Frm1.vspdData.Col = ACol
    Frm1.vspdData.Row = ARow
    
    Frm1.vspdData.Action = 0    
    
    Frm1.vspdData.ScrollBars = Parent.SS_SCROLLBAR_BOTH
    
End Function

'========================================
Function FncExit()
    FncExit = True
End Function

'========================================
Function DbQuery() 

	Err.Clear														
	DbQuery = False													
	
	If LayerShowHide(1) = False Then
		Exit Function
	End If
	
	Dim strVal
	
    With frm1
		strVal = BIZ_PGM_ID & "?txtHMode=" & Parent.UID_M0001
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			' Scroll시 
			strVal = strVal & "&txtYear=" & Trim(.txtHYear.value)
			strVal = strVal & "&txtQueryData=" & Trim(.txtHQueryData.value)
			strVal = strVal & "&txtSalesOrgLvl=" & Trim(.txtHSalesOrgLvl.value)
			strVal = strVal & "&txtSalesOrg=" & Trim(.txtHSalesOrg.value)
		Else
			' 처음 조회시 
			strVal = strVal & "&txtYear=" & Trim(.cboYear.value)
			strVal = strVal & "&txtQueryData=" & Trim(.cboQueryData.value)
			strVal = strVal & "&txtSalesOrgLvl=" & Trim(.cboSalesOrgLvl.value)
			strVal = strVal & "&txtSalesOrg=" & Trim(.txtSalesOrg.value)
		End If

        strVal = strVal & "&lgPageNo="		 & lgPageNo					 
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & Replace(Replace(MakeSQLGroupOrderByList("A"),"1?", "" & FilterVar("총계", "''", "S") & ""),"2?","" & FilterVar("소계", "''", "S") & "")
		strVal = strVal & "&lgSelectList="   & Replace(Replace(EnCoding(GetSQLSelectList("A")),"1?", "" & FilterVar("총계", "''", "S") & ""),"2?","" & FilterVar("소계", "''", "S") & "")    						'☜: Select list
	End With
	
	Call RunMyBizASP(MyBizASP, strVal)
    DbQuery = True    

End Function

'=========================================
Function DbQueryOk()

	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True
		If lgIntFlgMode <> Parent.OPMD_UMODE Then
			frm1.vspdData.Row = 1
			Call vspdData_Click(1, 1)
		    Call SetToolbar("11000000000111")
		End If
		lgIntFlgMode = Parent.OPMD_UMODE
	Else
		frm1.cboYear.focus
	End If

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
					<TD CLASS="CLSLTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>영업조직별월매출실적조회</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* >&nbsp;</TD>
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
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>매출년도</TD>
	                        		<TD CLASS="TD6" NOWRAP>
                						<SELECT Name="cboYear" ALT="매출년도" CLASS ="cbonormal" tag="12"><OPTION></OPTION></SELECT>
		                    		</TD>
									<TD CLASS="TD5" NOWRAP>조회기준</TD>
	                        		<TD CLASS="TD6" NOWRAP>
                						<SELECT Name="cboQueryData" ALT="조회기준" CLASS ="cbonormal" tag="12"><OPTION></OPTION></SELECT>
		                    		</TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>영업조직레벨</TD>
	                        		<TD CLASS="TD6" NOWRAP>
                						<SELECT Name="cboSalesOrgLvl" ALT="영업조직레벨" CLASS ="cbonormal" tag="12"><OPTION></OPTION></SELECT>
		                    		</TD>
									<TD CLASS=TD5 NOWRAP>영업조직</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSalesOrg" ALT="영업조직" TYPE="Text" MAXLENGTH="10" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSalesOrg" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopup C_PopSalesOrg">&nbsp;<INPUT NAME="txtSalesOrgNm" TYPE="Text" MAXLENGTH="20" SIZE=25 tag="14"></TD>
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
										<script language =javascript src='./js/s5111qa2_vaSpread1_vspdData.js'></script>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX ="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHYear" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHQueryData" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesOrgLvl" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSalesOrg" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
