<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procurement
'*  2. Function Name        : 
'*  3. Program ID           : m5211qa1
'*  4. Program Name         : 선적현황조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Jin-hyun Shin
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : 
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- '******************************************  1.1 Inc 선언   ****************************************** !-->
<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  =====================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 공통 Include   =====================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit					

<!-- #Include file="../../inc/lgvariables.inc" -->	
Dim lgIsOpenPop                                          
Dim lgMark                                                
Dim IscookieSplit 

Dim lgSaveRow    
                                       
Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"

EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID 		= "m5211qb1_KO441.asp"                     
Const BIZ_PGM_JUMP_ID1 	= "m5211ma1"                         
Const BIZ_PGM_JUMP_ID2 	= "m5211qa2"   
Const Major_Cd_Incoterms= "B9006"
Const C_MaxKey          = 22					             

'==========================================  setCookie()  ======================================
Function setCookie_01()

	if frm1.vspdData.maxrows > 0 then

		frm1.vspdData.row = frm1.vspdData.ActiveRow
		frm1.vspdData.col =  GetKeyPos("A", 3)

		WriteCookie "BlNo", Trim(frm1.vspdData.Text)
	end if
	
	Call PgmJump(BIZ_PGM_JUMP_ID1)

End Function

Function setCookie_02()

Dim strCfmFlg

	if frm1.vspdData.maxrows > 0 then
		if frm1.rdoCfmFlg0.checked then
			strCfmFlg = ""
		elseif frm1.rdoCfmFlg1.checked then
			strCfmFlg = "Y"
		else
			strCfmFlg = "N"
		end if		

		frm1.vspdData.row = frm1.vspdData.ActiveRow
		frm1.vspdData.col =  GetKeyPos("A", 3)

		WriteCookie "BlNo", Trim(frm1.vspdData.Text)
		WriteCookie "txtBeneficiaryCd", Trim(frm1.txtBeneficiaryCd.Value)
		WriteCookie "txtIncotermsCd", Trim(frm1.txtIncotermsCd.Value)
		WriteCookie "txtPurGrpCd", Trim(frm1.txtPurGrpCd.Value)
		WriteCookie "rdoCfmFlg", strCfmFlg
		WriteCookie "txtBlIssueFrDt", frm1.txtBlIssueFrDt.Text
		WriteCookie "txtBlIssueToDt", frm1.txtBlIssueToDt.Text
		WriteCookie "txtLoadingFrDt", frm1.txtLoadingFrDt.Text
		WriteCookie "txtLoadingToDt", frm1.txtLoadingToDt.Text
	end if
	
	Call PgmJump(BIZ_PGM_JUMP_ID2)

End Function


Function GetCookies()

	Dim strCfmFlg
	Dim strQueryFlg

	if ReadCookie("BlNo") <> "" then
		strQueryFlg					= ReadCookie("BlNo")
		frm1.txtBeneficiaryCd.Value	= ReadCookie("txtBeneficiaryCd")
		frm1.txtPurGrpCd.Value		= ReadCookie("txtPurGrpCd")
		frm1.txtIncotermsCd.Value	= ReadCookie("txtIncotermsCd")
		strCfmFlg					= ReadCookie("rdoCfmFlg")
		frm1.txtBlIssueFrDt.Text	= ReadCookie("txtBlIssueFrDt")
		frm1.txtBlIssueToDt.Text	= ReadCookie("txtBlIssueToDt")
		frm1.txtLoadingFrDt.Text	= ReadCookie("txtLoadingFrDt")
		frm1.txtLoadingToDt.Text	= ReadCookie("txtLoadingToDt")
		
		if	strCfmFlg = "" then
			frm1.rdoCfmFlg0.checked = true
		elseif strCfmFlg = "Y" then
			frm1.rdoCfmFlg1.checked = true
		else
			frm1.rdoCfmFlg2.checked = true
		end if	

		WriteCookie "BlNo",""
		WriteCookie "txtBeneficiaryCd",""
		WriteCookie "txtPurGrpCd",""
		WriteCookie "txtIncotermsCd",""
		WriteCookie "txtBlIssueFrDt",""
		WriteCookie "txtBlIssueToDt",""
		WriteCookie "txtLoadingFrDt",""
		WriteCookie "txtLoadingToDt",""
	end if
	
	if strQueryFlg <> "" then Call dbQuery

End Function

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgStrPrevKey     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
    lgIntFlgMode = parent.OPMD_CMODE   
    lgPageNo         = ""
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()

	frm1.txtBlIssueFrDt.Text	= StartDate
	frm1.txtBlIssueToDt.Text	= EndDate
	frm1.txtLoadingFrDt.Text	= StartDate
	frm1.txtLoadingToDt.Text	= EndDate 

	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPurGrpCd, "Q") 
		frm1.txtPurGrpCd.Tag = left(frm1.txtPurGrpCd.Tag,1) & "4" & mid(frm1.txtPurGrpCd.Tag,3,len(frm1.txtPurGrpCd.Tag))
        frm1.txtPurGrpCd.value = lgPGCd
	End If
	
 End Sub
'===========================  InitComboBox()  ============================================
Sub InitComboBox()
	Call SetCombo(frm1.cboPrcFlg, "T", "진단가")
	Call SetCombo(frm1.cboPrcFlg, "F", "가단가")
End Sub
'===========================  LoadInfTB19029()  ============================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("Q", "M", "NOCOOKIE", "QA") %>
<% Call LoadBNumericFormatA("Q", "M", "NOCOOKIE", "QA") %>
End Sub
'======================= 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("M5211QA101","S","A","V20030319",parent.C_SORT_DBAGENT,frm1.vspdData, _
									C_MaxKey, "X","X")
								
    Call SetSpreadLock 
 
End Sub
'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()    
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadLockWithOddEvenRowColor()
End Sub

'------------------------------------------  OpenBeneficiary()  -------------------------------------------------
'	Name : OpenBeneficiary()
'	Description : Supplier PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenBeneficiary()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "수출자"					
	arrParam(1) = "B_Biz_Partner"				
	arrParam(2) = Trim(frm1.txtBeneficiaryCd.Value)		
'	arrParam(3) = Trim(frm1.txtBpNm.Value)		
'	arrParam(4) = "BP_TYPE <> 'C'"				
    arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "	
	arrParam(5) = "수출자"					
	
    arrField(0) = "BP_CD"						
    arrField(1) = "BP_NM"						
    
    arrHeader(0) = "수출자"					
    arrHeader(1) = "수출자명"				
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtBeneficiaryCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtBeneficiaryCd.Value = arrRet(0)
		frm1.txtBeneficiaryNm.Value = arrRet(1)
		frm1.txtBeneficiaryCd.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenIncoterms()  -------------------------------------------------
'	Name : OpenIncoterms()
'	Description : OpenIncoterms PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenIncoterms()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	arrParam(0) = "가격조건"					
	arrParam(1) = "B_Minor"			
	arrParam(2) = Trim(frm1.txtIncotermsCd.Value)	
'	arrParam(3) = Trim(frm1.txtPoTypeNm.Value)	
	arrParam(4) = "Major_Cd=  " & FilterVar(Major_Cd_Incoterms , "''", "S") & ""
	arrParam(5) = "가격조건"					
	
    arrField(0) = "Minor_Cd"						
    arrField(1) = "Minor_Nm"						
        
    arrHeader(0) = "가격조건"					
    arrHeader(1) = "가격조건명"					
    
    arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtIncotermsCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtIncotermsCd.Value = arrRet(0)
		frm1.txtIncotermsNm.Value = arrRet(1)
		frm1.txtIncotermsCd.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenPurGrp()  -------------------------------------------------
'	Name : OpenPurGrp()
'	Description : PurGrp PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function
    If frm1.txtPurGrpCd.className = "protected" Then Exit Function
    
	lgIsOpenPop = True

	arrParam(0) = "구매그룹"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtPurGrpCd.Value)
'	arrParam(3) = Trim(frm1.txtPurGrpNm.Value)	
	
	arrParam(4) = ""
	arrParam(5) = "구매그룹"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "구매그룹"		
    arrHeader(1) = "구매그룹명"
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPurGrpCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtPurGrpCd.Value = arrRet(0)
		frm1.txtPurGrpNm.Value = arrRet(1)
		frm1.txtPurGrpCd.focus
		Set gActiveElement = document.activeElement
	End If	

End Function 

'------------------------------------  PopZAdoConfigGrid()  ----------------------------------------------
Sub PopZAdoConfigGrid()
	If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
		Exit Sub
	End If
	
	Call OpenGroupPopup("A")
End Sub

'------------------------------------  OpenGroupPopup()  ----------------------------------------------
Function OpenGroupPopup(ByVal pSpdNo)

	Dim arrRet

	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
   
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()

    Call LoadInfTB19029		

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")   	
 	Call InitVariables														
    Call GetValue_ko441()
	Call SetDefaultVal	
	Call InitSpreadSheet()
    Call SetToolbar("1100000000001111")		
	Call GetCookies()
    frm1.txtBeneficiaryCd.focus

    
End Sub
'===========================  Form_QueryUnload()  ============================================
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub
'===========================  vspdData_MouseDown()  ============================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'===========================  FncSplitColumn()  ============================================
Function FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Function
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Function

'===========================  OCX_EVENT()  ============================================
Sub txtBlIssueFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBlIssueFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBlIssueFrDt.focus
	End If
End Sub

Sub txtBlIssueToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtBlIssueToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtBlIssueToDt.focus
	End If
End Sub

Sub txtBlIssueFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

Sub txtBlIssueToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub


Sub txtLoadingFrDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtLoadingFrDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLoadingFrDt.focus
	End If
End Sub

Sub txtLoadingToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtLoadingToDt.Action = 7
		Call SetFocusToDocument("M")
		frm1.txtLoadingToDt.focus
	End If
End Sub

Sub txtLoadingFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub

Sub txtLoadingToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call FncQuery()
End Sub
'===========================  vspdData_GotFocus()  ============================================
Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub
'===========================  vspdData_DblClick()  ============================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If frm1.vspdData.MaxRows > 0 Then
		If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
		'	Call CookiePage(1)
		End If
	End If
End Function
'===========================  vspdData_Click()  ============================================	
Sub vspdData_Click(ByVal Col, ByVal Row)
    
    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    Call SetPopupMenuItemInf("00000000001")
    
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
           ggoSpread.SSSort, lgSortKey
           lgSortKey = 2
        Else
            ggoSpread.SSSort, lgSortKey
            lgSortKey = 1
        End If    
    End If
    
	If Row < 1 Then Exit Sub

	IscookieSplit = ""
    
End Sub
'===========================  vspdData_TopLeftChange()  ============================================	
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
 	If OldLeft <> NewLeft Then
 	    Exit Sub
 	End If
    
 	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
 		If lgPageNo <> "" Then		                                                    '다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
 			If DbQuery = False Then
 				Exit Sub
 			End if
 		End If
 	End If
 End Sub

'*******************************  5.1 Toolbar(Main)에서 호출되는 Function *******************************
Function FncQuery() 

    FncQuery = False                                            
    
    Err.Clear                                                   

    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables 											
	
	with frm1
		if (UniConvDateToYYYYMMDD(.txtBlIssueFrDt.text,gDateFormat,"") > UniConvDateToYYYYMMDD(.txtBlIssueToDt.text,gDateFormat,"")) And Trim(.txtBlIssueFrDt.text) <> "" And Trim(.txtBlIssueToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","B/L접수일", "X")	
			Exit Function
		End if   
		if (UniConvDateToYYYYMMDD(.txtLoadingFrDt.text,gDateFormat,"") > UniConvDateToYYYYMMDD(.txtLoadingToDt.text,gDateFormat,"")) And Trim(.txtLoadingFrDt.text) <> "" And Trim(.txtLoadingToDt.text) <> "" then	
			Call DisplayMsgBox("17a003", "X","선적일", "X")	
			Exit Function
		End if   
	End with

    If DbQuery = False Then Exit Function

    FncQuery = True													
	Set gActiveElement = document.activeElement
End Function
'===========================  FncSave()  ============================================	
Function FncSave()     
End Function
'===========================  FncPrint()  ============================================	
Function FncPrint() 
    Call parent.FncPrint()
    Set gActiveElement = document.activeElement
End Function
'===========================  FncExcel()  ============================================	
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.activeElement
End Function
'===========================  FncFind()  ============================================	
Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False) 
    Set gActiveElement = document.activeElement                           
End Function
'===========================  FncExit()  ============================================	
Function FncExit()
    FncExit = True
    Set gActiveElement = document.activeElement
End Function
'===========================  DbQuery()  ============================================
Function DbQuery() 
	Dim strVal
	Dim strCfmFlg

    DbQuery = False
    
    Err.Clear                                                       

	If  LayerShowHide(1) = False Then
      	Exit Function
    End If

    
    With frm1
	
		if .rdoCfmFlg0.checked then
			strCfmFlg = ""
		elseif .rdoCfmFlg1.checked then
			strCfmFlg = "Y"
		else
			strCfmFlg = "N"
		end if
	  
	If lgIntFlgMode = parent.OPMD_UMODE Then	
		strVal = BIZ_PGM_ID & "?txtBpCd=" & FilterVar(Trim(.hdnBeneficiaryCd.value),"","SNM")
	    strVal = strVal & "&txtIncotermsCd=" & FilterVar(Trim(.hdnIncotermsCd.value),"","SNM")
	    strVal = strVal & "&txtPurGrpCd=" & FilterVar(Trim(.hdnPurGrpCd.value),"","SNM")
    	strVal = strVal & "&txtBlFrDt=" & Trim(.hdnBlIssueFrDt.value)
    	strVal = strVal & "&txtBlToDt=" & Trim(.hdnBlIssueToDt.value)
    	strVal = strVal & "&txtLoadingFrDt=" & Trim(.hdnLoadingFrDt.value)    	
    	strVal = strVal & "&txtLoadingToDt=" & Trim(.hdnLoadingToDt.value)
    	strVal = strVal & "&txtCfmFlg=" & FilterVar(Trim(.hdnstrCfmFlg.value),"","SNM")
       
	else
		strVal = BIZ_PGM_ID & "?txtBpCd=" & FilterVar(Trim(.txtBeneficiaryCd.value),"","SNM")
	    strVal = strVal & "&txtIncotermsCd=" & FilterVar(Trim(.txtIncotermsCd.value),"","SNM")
	    strVal = strVal & "&txtPurGrpCd=" & FilterVar(Trim(.txtPurGrpCd.value),"","SNM")
    	strVal = strVal & "&txtBlFrDt=" & Trim(.txtBlIssueFrDt.Text)
    	strVal = strVal & "&txtBlToDt=" & Trim(.txtBlIssueToDt.Text)
    	strVal = strVal & "&txtLoadingFrDt=" & Trim(.txtLoadingFrDt.Text)    	
    	strVal = strVal & "&txtLoadingToDt=" & Trim(.txtLoadingToDt.Text)
    	strVal = strVal & "&txtCfmFlg=" & FilterVar(Trim(strCfmFlg),"","SNM")

	end if
		strVal = strVal & "&lgPageNo="   & lgPageNo         
		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
        strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))

        strVal = strVal & "&gBizArea=" & lgBACd 
        strVal = strVal & "&gPlant=" & lgPLCd 
        strVal = strVal & "&gPurGrp=" & lgPGCd 
        strVal = strVal & "&gPurOrg=" & lgPOCd  
		
        Call RunMyBizASP(MyBizASP, strVal)							
        
    End With
    
    DbQuery = True
    Call SetToolbar("1100000000011111")								

End Function
'===========================  DbQueryOk()  ============================================
Function DbQueryOk()												

	lgBlnFlgChgValue = False
    lgSaveRow        = 1
	lgIntFlgMode = parent.OPMD_UMODE  

	Call vspdData_Click(1,1)
	If frm1.vspdData.MaxRows > 0 Then
		frm1.vspdData.Focus
	Else
		frm1.txtBeneficiaryCd.focus
	End If						

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>B/L현황</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH="*" align=right><!--<button name="btnAutoSel" class="clsmbtn" ONCLICK="OpenGroupPopup()">정렬순서</button></td>-->
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
									<TD CLASS="TD5" NOWRAP>수출자</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="수출자" NAME="txtBeneficiaryCd"  SIZE=10 MAXLENGTH=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSpplCd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBeneficiary()">
														   <INPUT TYPE=TEXT NAME="txtBeneficiaryNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>B/L접수일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m5211qa1_fpDateTime2_txtBlIssueFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m5211qa1_fpDateTime2_txtBlIssueToDt.js'></script>
												</td>
											</tr>
										</table>
							         </TD>				   
								</TR>					   
								<TR>
									<TD CLASS="TD5" NOWRAP>가격조건</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT ALT="가격조건"  NAME="txtIncotermsCd" SIZE=10 LANG="ko" MAXLENGTH=5 tag="1XNXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPoType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenIncoterms() ">
														   <INPUT TYPE=TEXT NAME="txtIncotermsNm" SIZE=20 tag="14"></TD>
									<TD CLASS="TD5" NOWRAP>선적일</TD>
									<TD CLASS="TD6" NOWRAP>
										<table cellspacing=0 cellpadding=0>
											<tr>
												<td>
													<script language =javascript src='./js/m5211qa1_fpDateTime2_txtLoadingFrDt.js'></script>
												</td>
												<td>~</td>
												<td>
													<script language =javascript src='./js/m5211qa1_fpDateTime2_txtLoadingToDt.js'></script>
												</td>
											</tr>
										</table>
							         </TD>
	                            </TR>
	                            <TR>
									<TD CLASS="TD5" NOWRAP>구매그룹</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT Alt="구매그룹" NAME="txtPurGrpCd" SIZE=10 MAXLENGTH=4  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPurGrp" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPurGrp()">
														   <INPUT TYPE=TEXT NAME="txtPurGrpNm" SIZE=20 tag="14"></TD>	
									<TD CLASS="TD5" NOWRAP>확정여부</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg0" CLASS="RADIO" checked tag="11"><label for="rdoCfmFlg0">&nbsp;전체&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg1" CLASS="RADIO" tag="11"><label for="rdoCfmFlg1">&nbsp;확정&nbsp;</label>
														   <INPUT TYPE=radio AlT="확정여부" NAME="rdoCfmFlg" ID="rdoCfmFlg2" CLASS="RADIO" checked tag="11"><label for="rdoCfmFlg2">&nbsp;미확정&nbsp;&nbsp;</label></TD>
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
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/m5211qa1_vaSpread1_vspdData.js'></script>
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
    <TR HEIGHT="20">
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH="*" ALIGN="RIGHT"><a ONCLICK="VBSCRIPT:setCookie_02()">B/L상세조회</a> | <a ONCLICK="VBSCRIPT:setCookie_01()">B/L등록</a></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
    </TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="hdnBeneficiaryCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIncotermsCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPurGrpCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBlIssueFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBlIssueToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnLoadingFrDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnLoadingToDt" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnstrCfmFlg" tag="24">
</FORM>
    <DIV ID="MousePT" NAME="MousePT">
    <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
    </DIV>
</BODY>
</HTML>
