<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'**********************************************************************************************
'*  1. Module Name          : 영업
'*  2. Function Name        : 출하관리
'*  3. Program ID           : s4511aa1
'*  4. Program Name         : 수주참조
'*  5. Program Desc         : 출하요청등록을 위한 수주참조
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2005/12/12
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Ahn Junesun
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2002/12/13 Include 성능향상 강준구
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE>{{수주참조}}</TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" --> 
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentA.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliDBAgentVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMA_KO441.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/incImage.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit

' External ASP File
'========================================
Const BIZ_PGM_ID 		= "s4511ab1.asp"                              

' Constant variables 
'========================================
Const C_MaxKey          = 15                                          

Const C_PopPlant		= 1			' 공장
Const C_PopShipToParty	= 2			' 납품처
Const C_PopMovType		= 3			' 출하형태
Const C_PopSoType		= 4			' 출하형태
Const C_PopSalesGrp		= 5			' 영업그룹
Const C_PopSoNo			= 6			' 수주번호

' Common variables 
'========================================
<!-- #Include file="../../inc/lgvariables.inc" -->	

' User-defind Variables
'========================================
Dim lgIsOpenPop
Dim lgStrAllocInvFlag				' 재고할당 사용여부 (Y:사용, N:사용하지 않음)
Dim arrReturn						'☜: Return Parameter Group
Dim arrParam
Dim arrParent
ArrParent = window.dialogArguments
Set PopupParent  = ArrParent(0)
top.document.title = PopupParent.gActivePRAspName

Dim lgCurDate, lgStartDate

'------ 서버의 오늘.
lgCurDate = UNIConvDateAtoB("<%=GetSvrDate%>", PopupParent.gServerDateFormat, PopupParent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜(7일전으로 설정) ------
lgStartDate = UnIDateAdd("D", -7, lgCurDate, PopupParent.gDateFormat)

'========================================
Function InitVariables()
	lgPageNo         = ""
    lgBlnFlgChgValue = False	                           
    lgIntFlgMode     = PopupParent.OPMD_CMODE                          
    lgSortKey        = 1   
        
    lgIsOpenPop = False
    Redim arrReturn(0)        
    Self.Returnvalue = arrReturn     
End Function

'========================================
Sub SetDefaultVal()
	on error resume next
	frm1.txtConFrDlvyDt.Text = lgStartDate
	frm1.txtConToDlvyDt.Text = lgCurDate
	frm1.txtConPlantCd.value = PopupParent.gPlant
						
	frm1.txtConPlantCd.focus
	
	If lgPLCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtConPlantCd, "Q") 
        	frm1.txtConPlantCd.value = lgPLCd
	End If
	
	If lgSGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtConSalesGrp, "Q") 
        	frm1.txtConSalesGrp.value = lgSGCd
	End If
		  
End Sub

'========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "Q", "S", "NOCOOKIE", "PA") %>
End Sub

'========================================
Sub InitSpreadSheet()
	Call SetZAdoSpreadSheet("S4511AA1","S","A","V20051218", PopupParent.C_SORT_DBAGENT, frm1.vspdData, C_MaxKey, "X", "X" )
	Call SetSpreadLock 	  
End Sub

Sub SetSpreadLock()
	ggoSpread.SpreadLockWithOddEvenRowColor()
	frm1.vspdData.OperationMode = 3
End Sub	

'========================================
Function OKClick()

	Dim intColCnt
		
	If frm1.vspdData.ActiveRow > 0 Then	
		
		Redim arrReturn(C_MaxKey - 1)
		frm1.vspdData.Row = frm1.vspdData.ActiveRow
			
		For intColCnt = 0 To C_MaxKey - 1
			frm1.vspdData.Col = GetKeyPos("A",intColCnt + 1)
			arrReturn(intColCnt) = frm1.vspdData.Text
		Next	
					
	End If
		
	Self.Returnvalue = arrReturn
	Self.Close()
	
End Function

'========================================
Function CancelClick()
	Redim arrReturn(0)
	arrReturn(0) = ""
	Self.Returnvalue = arrReturn
	Self.Close()
End Function

' 조회조건 Popup
'=========================================
Function OpenConPopUp(Byval pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgIsOpenPop Then Exit Function

	lgIsOpenPop = True

	With frm1
		Select Case pvIntWhere
			Case C_PopPlant		'공장
			
			If frm1.txtConPlantCd.className = "protected" Then
			lgIsOpenPop = FALSE
			 Exit Function
			End If
			
				iArrParam(1) = "dbo.B_PLANT"									
				iArrParam(2) = Trim(.txtConPlantCd.value)				
				iArrParam(3) = ""										
				iArrParam(4) = ""										
				
				iArrField(0) = "ED15" & PopupParent.gColSep & "PLANT_CD"
				iArrField(1) = "ED30" & PopupParent.gColSep & "PLANT_NM"
							
				iArrHeader(0) = .txtConPlantCd.alt						
				iArrHeader(1) = .txtConPlantNm.alt					
	
				.txtConPlantCd.focus

			Case C_PopShipToParty	'납품처
				iArrParam(1) = "dbo.B_BIZ_PARTNER BP INNER JOIN dbo.B_COUNTRY CT ON (CT.COUNTRY_CD = BP.CONTRY_CD)"								
				iArrParam(2) = Trim(.txtConShipToParty.value)			
				iArrParam(3) = ""											
				iArrParam(4) = "BP.BP_TYPE IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND EXISTS (SELECT * FROM B_BIZ_PARTNER_FTN BPF WHERE BPF.PARTNER_BP_CD = BP.BP_CD AND BPF.PARTNER_FTN = " & FilterVar("SSH", "''", "S") & ")"						
	
				iArrField(0) = "ED15" & PopupParent.gColSep & "BP.BP_CD"
				iArrField(1) = "ED30" & PopupParent.gColSep & "BP.BP_NM"
				iArrField(2) = "ED10" & PopupParent.gColSep & "BP.CONTRY_CD"
				iArrField(3) = "ED20" & PopupParent.gColSep & "CT.COUNTRY_NM"
    
				iArrHeader(0) = .txtConShipToParty.alt					
				iArrHeader(1) = .txtConShipToPartyNm.alt					
				iArrHeader(2) = "{{국가}}"
				iArrHeader(3) = "{{국가명}}"

				.txtConShipToParty.focus
			
			Case C_PopMovType	'출하형태
				iArrParam(1) = "dbo.B_MINOR MN "		
				iArrParam(2) = Trim(.txtConMovType.value)					
				iArrParam(3) = ""											
				iArrParam(4) = "MN.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND EXISTS (SELECT * FROM dbo.S_SO_TYPE_CONFIG ST WHERE	ST.MOV_TYPE = MN.MINOR_CD) "			
				
				iArrField(0) = "ED15" & PopupParent.gColSep & "MN.MINOR_CD"
				iArrField(1) = "ED30" & PopupParent.gColSep & "MN.MINOR_NM"
				
				iArrHeader(0) = .txtConMovType.alt							
				iArrHeader(1) = .txtConMovTypeNm.alt	
				
				frm1.txtConMovType.focus

			' 수주형태
			Case C_PopSoType												
				iArrParam(1) = "dbo.S_SO_TYPE_CONFIG"
				iArrParam(2) = Trim(.txtConSOType.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & "  AND REL_DN_FLAG = " & FilterVar("Y", "''", "S") & " "
					
				iArrField(0) = "ED15" & PopupParent.gColSep & "SO_TYPE"
				iArrField(1) = "ED30" & PopupParent.gColSep & "SO_TYPE_NM"
    
			    iArrHeader(0) = .txtConSOType.Alt
			    iArrHeader(1) = .txtConSOTypeNm.Alt
				    
			    .txtConSOType.focus

			' 영업그룹
			Case C_PopSalesGrp
			
			If frm1.txtConSalesGrp.className = "protected" Then
			lgIsOpenPop = FALSE
			 Exit Function
			End If
															
				iArrParam(1) = "dbo.B_SALES_GRP"
				iArrParam(2) = Trim(.txtConSalesGrp.value)
				iArrParam(3) = ""
				iArrParam(4) = "USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
					
				iArrField(0) = "ED15" & PopupParent.gColSep & "SALES_GRP"
				iArrField(1) = "ED30" & PopupParent.gColSep & "SALES_GRP_NM"
    
			    iArrHeader(0) = .txtConSalesGrp.Alt
			    iArrHeader(1) = .txtConSalesGrpNm.Alt
				    
			    .txtConSalesGrp.focus

			Case C_PopSoNo
				iArrParam(1) = "S_SO_HDR SH, B_BIZ_PARTNER SP, B_SALES_GRP SG"
				iArrParam(2) = Trim(.txtConSONo.value)
				iArrParam(3) = ""
				
				' 재고할당을 사용여부
				If lgStrAllocInvFlag = "N" Then
					iArrParam(4) = "SH.SOLD_TO_PARTY = SP.BP_CD AND SH.SALES_GRP = SG.SALES_GRP AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & "  AND SH.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM S_SO_DTL SD WHERE SD.SO_NO = SH.SO_NO AND SD.SO_QTY + SD.BONUS_QTY > SD.REQ_QTY + SD.REQ_BONUS_QTY) "
				Else
					iArrParam(4) = "SH.SOLD_TO_PARTY = SP.BP_CD AND SH.SALES_GRP = SG.SALES_GRP AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & "  AND SH.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  AND EXISTS (SELECT * FROM S_SO_SCHD SC WHERE SC.SO_NO = SH.SO_NO AND SC.ALLC_QTY + SC.ALLC_BONUS_QTY > SC.REQ_QTY + SC.REQ_BONUS_QTY) "
				End If
				iArrParam(5) = "{{수주번호}}"

				iArrField(0) = "ED12" & PopupParent.gColSep & "SH.SO_NO"
				iArrField(1) = "ED10" & PopupParent.gColSep & "SH.SOLD_TO_PARTY"
				iArrField(2) = "ED20" & PopupParent.gColSep & "SP.BP_NM"
				iArrField(3) = "DD10" & PopupParent.gColSep & "SH.SO_DT"
				iArrField(4) = "ED15" & PopupParent.gColSep & "SG.SALES_GRP_NM"
				iArrField(5) = "ED10" & PopupParent.gColSep & "SH.PAY_METH"
				
				iArrHeader(0) = "{{수주번호}}"
				iArrHeader(1) = "{{주문처}}"
				iArrHeader(2) = "{{주문처명}}"
				iArrHeader(3) = "{{수주일}}"
				iArrHeader(4) = "{{영업그룹명}}"
				iArrHeader(5) = "{{결제방법}}"
				
				.txtConSONo.focus
		End Select
	End With
	
	iArrParam(0) = iArrHeader(0)							' 팝업 Title
	iArrParam(5) = iArrHeader(0)							' 조회조건 명칭

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConPopUp = SetConPopUp(iArrRet,pvIntWhere)
	End If	
	
End Function

'========================================
Function SetConPopUp(ByVal pvArrRet,ByVal pvIntWhere)

	With frm1
		Select Case pvIntWhere
			Case C_PopPlant
				.txtConPlantCd.value = pvArrRet(0)
				.txtConPlantNm.value = pvArrRet(1) 

			Case C_PopShipToParty
				.txtConShipToParty.value = pvArrRet(0)
				.txtConShipToPartyNm.value = pvArrRet(1) 

			Case C_PopMovType
				.txtConMovType.value = pvArrRet(0)
				.txtConMovTypeNm.value = pvArrRet(1) 

			Case C_PopSoType
				.txtConSOType.value = pvArrRet(0)
				.txtConSOTypeNm.value = pvArrRet(1) 

			Case C_PopSalesGrp
				.txtConSalesGrp.value = pvArrRet(0)
				.txtConSalesGrpNm.value = pvArrRet(1) 

			Case C_PopSoNo
				.txtConSoNo.value = pvArrRet(0)
		End Select
	End With

End Function

'========================================
Function OpenSortPopup()
	Dim arrRet
	
	On Error Resume Next
	
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData("A"),gMethodText),"dialogWidth=" & PopupParent.SORTW_WIDTH & "px; dialogHeight=" & PopupParent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData("A",arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

'========================================
Sub Form_Load()
    Call LoadInfTB19029													'⊙: Load table , B_numeric_format
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
	Call ggoOper.LockField(Document, "N")                         '⊙: Lock  Suitable  Field
    
	Call InitVariables	
	Call GetValue_ko441()										  
	Call SetDefaultVal	
	Call InitSpreadSheet()
	Call MM_preloadImages("../../../CShared/image/Query.gif","../../../CShared/image/OK.gif","../../../CShared/image/Cancel.gif")
	Call GetAllocInvFlag()			' 재고할당여부 Fetch
End Sub

'========================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub

'========================================
Function vspdData_DblClick(ByVal Col, ByVal Row)
	If Frm1.vspdData.ActiveRow > 0 Then	Call OKClick
End Function

'========================================
Function vspdData_KeyPress(KeyAscii)
     On Error Resume Next
     If KeyAscii = 13 And frm1.vspdData.ActiveRow > 0 Then    'Frm1없으면 frm1삭제
        Call OKClick()
     ElseIf KeyAscii = 27 Then
        Call CancelClick()
     End If
End Function

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
	
	If OldLeft <> NewLeft Then Exit Sub

	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크	
		If CheckRunningBizProcess Then Exit Sub
		If lgPageNo <> "" Then Call DbQuery
	End If		 
End Sub

'========================================
Sub txtConFrDlvyDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConFrDlvyDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtConFrDlvyDt.Focus
	End If
End Sub

'========================================
Sub txtConToDlvyDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConToDlvyDt.Action = 7
		Call SetFocusToDocument("P")
		frm1.txtConToDlvyDt.Focus
	End If
End Sub


'========================================
Sub txtConFrDlvyDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

'========================================
Sub txtConToDlvyDt_KeyPress(KeyAscii)
	On Error Resume Next
    If KeyAscii = 27 Then
       Call CancelClick()
    Elseif KeyAscii = 13 Then
       Call FncQuery()
    End if
End Sub

'========================================
Function FncQuery() 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               

	If Not chkField(Document, "1") Then Exit Function
	
	If ValidDateCheck(frm1.txtConFrDlvyDt, frm1.txtConToDlvyDt) = False Then Exit Function
   
    Call ggoOper.ClearField(Document, "2")	         						
    Call InitVariables 														
    
	If DbQuery = False Then Exit Function									

    FncQuery = True		
    
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
		If lgIntFlgMode = PopupParent.OPMD_UMODE Then		
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001					
			strVal = strVal & "&txtConPlantCd=" & Trim(.HPlantCd.value)				
			strVal = strVal & "&txtConSalesGrp=" & Trim(.HSalesGrp.value)				
			strVal = strVal & "&txtConSoNo=" & Trim(.HSoNo.value)				
			strVal = strVal & "&txtConShipToParty=" & Trim(.HShipToParty.value)
			strVal = strVal & "&txtConMovType=" & Trim(.HMovType.value)
			strVal = strVal & "&txtConSOType=" & Trim(.HSOType.value)
			strVal = strVal & "&txtConFrDlvyDt=" & Trim(.HFrDlvyDt.value)
			strVal = strVal & "&txtConToDlvyDt=" & Trim(.HToDlvyDt.value)				
        Else
			strVal = BIZ_PGM_ID & "?txtMode=" & PopupParent.UID_M0001			
			strVal = strVal & "&txtConPlantCd=" & Trim(.txtConPlantCd.value)				
			strVal = strVal & "&txtConSalesGrp=" & Trim(.txtConSalesGrp.value)				
			strVal = strVal & "&txtConSoNo=" & Trim(.txtConSoNo.value)
			strVal = strVal & "&txtConShipToParty=" & Trim(.txtConShipToParty.value)
			strVal = strVal & "&txtConMovType=" & Trim(.txtConMovType.value)
			strVal = strVal & "&txtConSOType=" & Trim(.txtConSOType.value)
			strVal = strVal & "&txtConFrDlvyDt=" & Trim(.txtConFrDlvyDt.text)
			strVal = strVal & "&txtConToDlvyDt=" & Trim(.txtConToDlvyDt.Text)				
		End If
			
        strVal = strVal & "&lgPageNo="			& lgPageNo						
        strVal = strVal & "&lgStrAllocInvFlag="	& lgStrAllocInvFlag						
		strVal = strVal & "&lgSelectListDT="	& GetSQLSelectListDataType("A")			 
		strVal = strVal & "&lgTailList="		& MakeSQLGroupOrderByList("A")
		strVal = strVal & "&lgSelectList="		& EnCoding(GetSQLSelectList("A"))

        Call RunMyBizASP(MyBizASP, strVal)		    						
        
    End With
    
    DbQuery = True    

End Function

'=========================================
Function DbQueryOk()
	If frm1.vspdData.MaxRows > 0 Then
		lgIntFlgMode = PopupParent.OPMD_UMODE
		frm1.vspdData.Focus
		frm1.vspdData.SelModeSelected = True		
	End If

End Function

' 재고할당 여부를 Fetch한다.
'=========================================
Sub GetAllocInvFlag()
	Dim iStrSelectList, iStrFromList, iStrWhereList, iStrRs
	Dim iArrRs 

	iStrSelectList = "REFERENCE"
	iStrFromList = "dbo.B_CONFIGURATION"
	iStrWhereList = "MAJOR_CD = " & FilterVar("S0017", "''", "S") & " AND MINOR_CD = " & FilterVar("A", "''", "S") & "  AND SEQ_NO = 1 "

	If CommonQueryRs2by2(iStrSelectList,iStrFromList,iStrWhereList, iStrRs) Then
		iArrRs = Split(iStrRs, Chr(11))
		lgStrAllocInvFlag = iArrRs(1)
	Else
		err.Clear
		lgStrAllocInvFlag = "N"
	End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">

<TABLE <%=LR_SPACE_TYPE_20%>>
	<TR>
		<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
	</TR>
	<TR>
		<TD HEIGHT=20 WIDTH=100%>
			<FIELDSET CLASS="CLSFLD">
				<TABLE <%=LR_SPACE_TYPE_40%>>
					<TR>
						<TD CLASS=TD5 NOWRAP>{{공장}}</TD>
						<TD CLASS=TD6><INPUT NAME="txtConPlantCd" TYPE="Text" Alt="{{공장}}" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConPlant" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopPlant">&nbsp;<INPUT NAME="txtConPlantNm" TYPE="Text" SIZE=25 tag="14" Alt="{{공장명}}"></TD>
						<TD CLASS="TD5" NOWRAP>{{납기일}}</TD>
						<TD CLASS="TD6" NOWRAP>
							<TABLE CELLSPACING=0 CELLPADDING=0>
								<TR>
									<TD>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtConFrDlvyDt" Alt="{{납기시작일}}" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
									<TD>
										&nbsp;~&nbsp;
									</TD>
									<TD>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtConToDlvyDt" Alt="{{납기종료일}}" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME"></OBJECT>');</SCRIPT>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>{{납품처}}</TD>
						<TD CLASS=TD6><INPUT NAME="txtConShipToParty" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="11XXXU" Alt="{{납품처}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConShipToParty" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopShipToParty">&nbsp;<INPUT NAME="txtConShipToPartyNm" TYPE="Text" SIZE=25 tag="14" Alt="{{납품처명}}"></TD>
						<TD CLASS=TD5 NOWRAP>{{S/O번호}}</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtConSoNo" SIZE=34 MAXLENGTH=18 TAG="11XXXU" ALT="{{S/O번호}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSoNo" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSoNo">&nbsp;</TD>
					</TR>
					<TR>
						<TD CLASS=TD5 NOWRAP>{{출하형태}}</TD>
						<TD CLASS=TD6><INPUT TYPE=TEXT NAME="txtConMovType" SIZE=10 MAXLENGTH=3 TAG="11XXXU" ALT="{{출하형태}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConMovType" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopMovType">&nbsp;<INPUT TYPE=TEXT NAME="txtConMovTypeNm" SIZE=25 TAG="14" Alt="{{출하형태명}}"></TD>
						<TD CLASS=TD5 NOWRAP>{{수주형태}}</TD>
						<TD CLASS=TD6 NOWRAP><INPUT TYPE=TEXT NAME="txtConSOType" SIZE=10 MAXLENGTH=10 TAG="11XXXU" ALT="{{수주형태}}"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSOType" align=top TYPE="BUTTON" OnClick="vbscript:OpenConPopUp C_PopSoType">&nbsp;<INPUT TYPE=TEXT NAME="txtConSOTypeNm" SIZE=25 TAG="14" Alt="{{수주형태명}}"></TD>
					</TR>					
					<TR>
						<TD CLASS=TD5 NOWRAP>{{영업그룹}}</TD>
						<TD CLASS=TD6><INPUT NAME="txtConSalesGrp" TYPE="Text" Alt="{{영업그룹}}" MAXLENGTH=4 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnConSalesGrp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopSalesGrp">&nbsp;<INPUT NAME="txtConSalesGrpNm" TYPE="Text" SIZE=25 tag="14" Alt="{{영업그룹명}}"></TD>	
						<TD CLASS=TD5 NOWRAP> </TD>
						<TD CLASS=TD6>&nbsp;</TD>
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
						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TAG="23" id=vaSpread TITLE="SPREAD"> <PARAM NAME="MaxRows" Value=0> <PARAM NAME="MaxCols" Value=0> <PARAM NAME="ReDraw" VALUE=0> </OBJECT>');</SCRIPT>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
    <TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
    </TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=70% NOWRAP>     <IMG SRC="../../../CShared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Query.gif',1)"  ONCLICK="FncQuery()"   ></IMG>
							                  <IMG SRC="../../../CShared/image/zpConfig_d.gif" Style="CURSOR: hand" ALT="Config" NAME="Config" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/zpConfig.gif',1)" OnClick="OpenSortPopup()" ></IMG>
					</TD>
					<TD WIDTH=30% ALIGN=RIGHT><IMG SRC="../../../CShared/image/ok_d.gif"     Style="CURSOR: hand" ALT="OK"     NAME="pop1"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/OK.gif',1)"     ONCLICK="OkClick()"    ></IMG>
							                  <IMG SRC="../../../CShared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2"   onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../../CShared/image/Cancel.gif',1)" ONCLICK="CancelClick()"></IMG>					</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX ="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="HSoNo" tag="24">
<INPUT TYPE=HIDDEN NAME="HShipToParty" tag="24">
<INPUT TYPE=HIDDEN NAME="HMovType" tag="24">
<INPUT TYPE=HIDDEN NAME="HSOType" tag="24">
<INPUT TYPE=HIDDEN NAME="HFrPromiseDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HToPromiseDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HFrDlvyDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HToDlvyDt" tag="24">
<INPUT TYPE=HIDDEN NAME="HPlantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="HSalesGrp" tag="24">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>
