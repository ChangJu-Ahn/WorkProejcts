<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111MA6
'*  4. Program Name         : 일괄출고처리 
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2003/07/01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwang Seongbae
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
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                               

Dim iDBSYSDate
Dim EndDate, StartDate

'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)

Const BIZ_PGM_ID = "s4111mb6.asp"											'☆: Head Query 비지니스 로직 ASP명 

Const C_PopPlant		= 1			' 공장 
Const C_PopDnType		= 2			' 출하형태 
Const C_PopShipToParty	= 3			' 납품처 

'☆: Spread Sheet의 Column별 상수 
Dim C_Select
Dim C_DnNo					' 출하번호 
Dim C_PromiseDt				' 출고예정일 
Dim C_ShipToParty			' 납품처 
Dim C_ShipToPartyNm			' 납품처명 
Dim C_MovType				' 출하형태 
Dim C_MovTypeNm				' 출하형태명 
Dim C_ExceptDnFlag			' 예외출고여부 

'=========================================
Dim lgBlnFlgChgValue           ' Variable is for Dirty flag
Dim lgIntFlgMode               ' Variable is for Operation Status
Dim lgBlnCfmChecked				' 작업유형 '출고처리' 선택여부 

Dim lgSortKey
Dim lgStrPrevKey

Dim lgBlnOpenPop						

'=========================================
Sub initSpreadPosVariables()
	C_Select		= 1
	C_DnNo			= 2
	C_PromiseDt		= 3
	C_ShipToParty	= 4
	C_ShipToPartyNm	= 5
	C_MovType		= 6
	C_MovTypeNm		= 7
	C_ExceptDnFlag	= 8
End Sub

'=========================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE            
    lgBlnFlgChgValue = False                    	
    lgStrPrevKey = ""   

    Call SetToolbar("11000000000011")										'⊙: 버튼 툴바 제어 
End Sub

'=========================================
Sub SetDefaultVal()
	lgBlnFlgChgValue = False
	lgBlnCfmChecked = True
	With frm1
		.rdoWorkTypeCfm.checked = True
		.txtConFromDt.Text = EndDate
		.txtConToDt.Text = EndDate
		.txtActualGIDt.Text = EndDate
		.txtConPlant.focus
	End With
End Sub

'=========================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
	<% Call LoadBNumericFormatA( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'=========================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    

	With ggoSpread
		.Source = frm1.vspdData
		.Spreadinit "V20030701",,parent.gAllowDragDropSpread    
	    
		frm1.vspdData.ReDraw = false
			
		frm1.vspdData.MaxCols = C_ExceptDnFlag + 1											'☜: 최대 Columns의 항상 1개 증가시킴	    
		frm1.vspdData.MaxRows = 0
	
		Call GetSpreadColumnPos("A")

					   'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)
		.SSSetCheck		C_Select,			"선택",			10,,,true
	    .SSSetEdit		C_DnNo,				"출하번호",		20,,,,2
		.SSSetDate		C_PromiseDt,		"출고예정일",	15,2,Parent.gDateFormat    
	    .SSSetEdit		C_ShipToParty,		"납품처",		15,,,,2
	    .SSSetEdit		C_ShipToPartyNm,	"납품처명",		20,,,,2
	    .SSSetEdit		C_MovType,			"출하형태",		15,,,,2
	    .SSSetEdit		C_MovTypeNm,		"출하형태명",	20,,,,2
	    .SSSetEdit		C_ExceptDnFlag,		"",	0
		
	    Call .SSSetColHidden(C_ExceptDnFlag, frm1.vspdData.MaxCols, True)
	    
	    Call SetSpreadLock
    End With
    
	frm1.vspdData.ReDraw = True

End Sub

'=========================================
Sub SetSpreadLock()
	ggoSpread.SpreadLock 1 , -1
	ggoSpread.SpreadUnLock	C_Select, -1, C_Select
End Sub

' 에러 발생시 해당 위치로 Focus이동 
'=========================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow

           If Not Frm1.vspdData.ColHidden Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If
       Next
    End If   
End Sub

' 조회조건 Popup
'=========================================
Function OpenConPopUp(Byval pvIntWhere)

	Dim iArrRet
	Dim iArrParam(5), iArrField(6), iArrHeader(6)

	If lgBlnOpenPop Then Exit Function

	lgBlnOpenPop = True

	With frm1
		Select Case pvIntWhere
			Case C_PopPlant		'공장 
				iArrParam(1) = "dbo.B_PLANT"									
				iArrParam(2) = Trim(.txtConPlant.value)				
				iArrParam(3) = ""										
				iArrParam(4) = ""										
				
				iArrField(0) = "ED15" & Parent.gColSep & "PLANT_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "PLANT_NM"
							
				iArrHeader(0) = .txtConPlant.alt						
				iArrHeader(1) = .txtConPlantNm.alt					
	
				.txtConPlant.focus

			Case C_PopDnType	'출하형태 
				iArrParam(1) = "dbo.B_MINOR MN "		
				iArrParam(2) = Trim(.txtConDnType.value)					
				iArrParam(3) = ""											
				iArrParam(4) = "MN.MAJOR_CD = " & FilterVar("I0001", "''", "S") & " AND EXISTS (SELECT * FROM dbo.S_SO_TYPE_CONFIG ST WHERE	ST.MOV_TYPE = MN.MINOR_CD) "			
				
				iArrField(0) = "ED15" & Parent.gColSep & "MN.MINOR_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "MN.MINOR_NM"
				
				iArrHeader(0) = .txtConDnType.alt							
				iArrHeader(1) = .txtConDnTypeNm.alt	
				
				frm1.txtConDnType.focus

			Case C_PopShipToParty	'납품처 
				iArrParam(1) = "dbo.B_BIZ_PARTNER BP INNER JOIN dbo.B_COUNTRY CT ON (CT.COUNTRY_CD = BP.CONTRY_CD)"								
				iArrParam(2) = Trim(.txtConShipToParty.value)			
				iArrParam(3) = ""											
				iArrParam(4) = "BP.BP_TYPE IN (" & FilterVar("C", "''", "S") & " , " & FilterVar("CS", "''", "S") & ") AND EXISTS (SELECT * FROM B_BIZ_PARTNER_FTN BPF WHERE BPF.PARTNER_BP_CD = BP.BP_CD AND BPF.PARTNER_FTN = " & FilterVar("SSH", "''", "S") & ")"						
	
				iArrField(0) = "ED15" & Parent.gColSep & "BP.BP_CD"
				iArrField(1) = "ED30" & Parent.gColSep & "BP.BP_NM"
				iArrField(2) = "ED10" & Parent.gColSep & "BP.CONTRY_CD"
				iArrField(3) = "ED20" & Parent.gColSep & "CT.COUNTRY_NM"
    
				iArrHeader(0) = .txtConShipToParty.alt					
				iArrHeader(1) = .txtConShipToPartyNm.alt					
				iArrHeader(2) = "국가"
				iArrHeader(3) = "국가명"

				.txtConShipToParty.focus
			
		End Select
	End With
	
	iArrParam(0) = iArrHeader(0)							' 팝업 Title
	iArrParam(5) = iArrHeader(0)							' 조회조건 명칭 

	iArrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(iArrParam, iArrField, iArrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	lgBlnOpenPop = False

	If iArrRet(0) <> "" Then
		OpenConPopUp = SetConPopUp(iArrRet,pvIntWhere)
	End If	
	
End Function

'========================================
Function SetConPopUp(ByVal pvArrRet,ByVal pvIntWhere)

	With frm1
		Select Case pvIntWhere
			Case C_PopPlant
				.txtConPlant.value = pvArrRet(0)
				.txtConPlantNm.value = pvArrRet(1) 

			Case C_PopDnType
				.txtConDnType.value = pvArrRet(0)
				.txtConDnTypeNm.value = pvArrRet(1) 

			Case C_PopShipToParty
				.txtConShipToParty.value = pvArrRet(0)
				.txtConShipToPartyNm.value = pvArrRet(1) 

		End Select
	End With

End Function

'=====================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_Select		= iCurColumnPos(1)
			C_DnNo			= iCurColumnPos(2)
			C_PromiseDt		= iCurColumnPos(3)
			C_ShipToParty	= iCurColumnPos(4)			
			C_ShipToPartyNm	= iCurColumnPos(5)	
			C_MovType		= iCurColumnPos(6)
			C_MovTypeNm		= iCurColumnPos(7)
			C_ExceptDnFlag	= iCurColumnPos(8)
    End Select    
End Sub

'==========================================================================================================
Function CookiePage(Byval pvKubun)

	On Error Resume Next
	Const CookieSplit = 4877      <%'Cookie Split String : CookiePage Function Use%>
	Dim iStrTemp, iArrVal

	With frm1
		If pvKubun = 1 Then
			WriteCookie CookieSplit , .txtConPlant.value & Parent.gColSep & .txtPromiseDt.Text & Parent.gColSep & _
									  .txtConMovType.value & Parent.gColSep & .txtConShipToPartylue
		ElseIf pvKubun = 0 Then
			iStrTemp = ReadCookie(CookieSplit)
			
			If Trim(Replace(iStrTemp, parent.gColSep, "")) = "" Then Exit Function
			
			iArrVal = Split(iStrTemp, Parent.gColSep)

			.txtConPlant.value			= iArrVal(0)
			.txtConFromDt.Text			= iArrVal(1)
			.txtConToDt.Text			= iArrVal(1)
			.txtConDnType.value			= iArrVal(2)
			.txtConShipToParty.value	= iArrVal(3)
			WriteCookie CookieSplit , ""
		End If
	End With
End Function

'========================================
Sub Form_Load()
	Call LoadInfTB19029              '⊙: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat,Parent.gComNum1000, Parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field
	Call InitSpreadSheet

	Call SetDefaultVal
	Call CookiePage(0)
	Call InitVariables
End Sub

'========================================
Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

'========================================
Sub rdoWorkTypeCfm_OnClick()
	If Not lgBlnCfmChecked Then
		lgBlnCfmChecked = True
		idDtTitle.innerHTML = "출고예정일"
	End If
End Sub

'========================================
Sub rdoWorkTypeCancel_OnClick()
	If lgBlnCfmChecked Then
		lgBlnCfmChecked = False
		idDtTitle.innerHTML = "실제출고일"
	End If
End Sub

'========================================
Sub txtConFromDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConFromDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtConFromDt.focus
	End If
End Sub

'========================================
Sub txtConToDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtConToDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtConToDt.focus
	End If
End Sub

'========================================
Sub txtActualGIDt_DblClick(Button)
	If Button = 1 Then
		frm1.txtActualGiDt.Action = 7
		Call SetFocusToDocument("M")	
		frm1.txtActualGiDt.focus
	End If
End Sub

'========================================
Sub txtConFromDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()	
End Sub

'========================================
Sub txtConToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

' 전체선택 
'========================================
Sub chkSelectAll_onClick()
	Dim iStrOldValue
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub

	ggoSpread.Source = frm1.vspdData	
	With frm1.vspdData
		.Row = 1			:	.Row2 = .MaxRows
		
		' 전체선택 
		If frm1.chkSelectAll.checked Then
			' Row Header 설정(수정)
			.Col = 0			:	.Col2 = 0
			.Clip = Replace(.Clip, vbCrLf, ggoSpread.UpdateFlag & vbCrLf)
			
			' 선택버튼의 선택여부 설정 
			.Col = C_Select		:	.Col2 = C_Select
			.Clip = Replace(.Clip, "0", "1")
			
		' 전체선택 취소 
		Else
			' Row Header 설정(수정)
			.Col = 0			:	.Col2 = 0
			.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")

			.Col = C_SELECT		:	.Col2 = C_SELECT
			.Clip = Replace(.Clip, "1", "0")
		End if
	End With

	' Active Cell 설정	
	Call SetActiveCell(frm1.vspdData,C_Select, 1,"M","X","X")
End Sub
'=====================================================
Sub chkVatFlag_OnClick()

	On Error Resume Next

	Select Case frm1.chkVatFlag.checked
		Case True
			frm1.chkARflag.checked = True  
	End Select

	lgBlnFlgChgValue = True

	If Err.number <> 0 Then Err.Clear
 
End Sub

'=====================================================
Sub chkARflag_OnClick()

	On Error Resume Next

	Select Case frm1.chkVatFlag.checked
		Case True
			frm1.chkVatFlag.checked = False
	End Select

	lgBlnFlgChgValue = True

	If Err.number <> 0 Then Err.Clear
 
End Sub

'========================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If lgIntFlgMode = Parent.OPMD_CMODE Then Exit Sub

	ggoSpread.Source = frm1.vspdData
	
	If Row > 0 Then
		Select Case Col
		Case C_Select
			If ButtonDown = 0 then					'---선택이 취소된 경우				
				frm1.vspddata.row = Row
				Call FncCancel()
			Else									'--- 선택된 경우 
				ggoSpread.UpdateRow Row	
			End If			
		End Select
	End If

End Sub

'=======================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	Call SetPopupMenuItemInf("1101111111")
	
    gMouseClickStatus = "SPC"
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then 
		Exit Sub
	End If  
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    	

End Sub

'========================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'========================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    If Col <= C_SELECT Or NewCol <= C_SELECT Then
        Cancel = True
        Exit Sub
    End If
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================
Sub vspdData_TopLeftChange(ByVal OldLeft ,ByVal OldTop ,ByVal NewLeft ,ByVal NewTop )
    
    If OldLeft <> NewLeft Then
       Exit Sub
    End If

    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKey <> "" Then                         
			If CheckRunningBizProcess Then Exit Sub
			Call DisableToolBar(Parent.TBC_QUERY)
            Call DbQuery()
        End If
    End if
End Sub

'=====================================================
Function FncQuery() 
    
    Dim IntRetCD 
        
    FncQuery = False                                                        
    
    Err.Clear                                                               
	
    If Not chkField(Document, "1") Then Exit Function

	If ValidDateCheck(frm1.txtConFromDt, frm1.txtConToDt) = False Then Exit Function

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

    Call ggoSpread.ClearSpreadData()
	frm1.chkSelectAll.checked = False
    Call InitVariables															

    Call DbQuery																<%'☜: Query db data%>

    FncQuery = True																
        
End Function

'=====================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")
    Call SetDefaultVal
    Call InitVariables															

    FncNew = True																

End Function

'=====================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         
    
    Err.Clear                                                               
    
	ggoSpread.Source = frm1.vspdData	
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001", "X", "X", "X")		
	    Exit Function
    End If

	' 출고처리시에 입력필수 항목 Check
	If frm1.txtHConPostFlag.value = "Y" Then
	    If Not chkField(Document, "2")Then Exit Function
    End If

    CAll DbSave
    
    FncSave = True                                                          
    
End Function

'=====================================================
Function FncCancel() 
	If frm1.vspdData.MaxRows < 1 Then Exit Function
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.EditUndo
End Function

'=====================================================
Function FncPrint() 
	Call parent.FncPrint()
End Function

'=====================================================
Function FncExcel() 
	Call parent.FncExport(Parent.C_SINGLEMULTI)
End Function

'=====================================================
Function FncFind() 
	Call parent.FncFind(Parent.C_SINGLEMULTI, False)
End Function

'=====================================================
Sub FncSplitColumn()
    
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
    
End Sub

'=====================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'=====================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    
	Call ggoSpread.ReOrderingSpreadData()
	
End Sub

'=====================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False

	ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")   '☜ 바뀐부분 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'=====================================================
Function DbQuery() 

    On Error Resume Next                                                          
    Err.Clear
    
	If LayerShowHide(1) = False Then
		Exit Function 
    End If
	  
	Dim iStrVal
	
    DbQuery = False
    
    With frm1
		
		' 재쿼리시(Scrollbar)
		iStrVal = BIZ_PGM_ID & "?txtMode="			& Parent.UID_M0001							
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			iStrVal = iStrVal & "&txtConPlant="			& Trim(.txtHConPlant.value)			
			iStrVal = iStrVal & "&txtConFromDt="		& Trim(.txtHConFromDt.value)
			iStrVal = iStrVal & "&txtConToDt="			& Trim(.txtHConToDt.value)		
			iStrVal = iStrVal & "&txtConDnType="		& Trim(.txtHConDnType.value)			
			iStrVal = iStrVal & "&txtConShipToParty="	& Trim(.txtHConShipToParty.value)		
			iStrVal = iStrVal & "&txtConPostFlag="		& Trim(.txtHConPostFlag.value)
			iStrVal = iStrVal & "&lgStrPrevKey="		& lgStrPrevKey
		Else
			iStrVal = iStrVal & "&txtConPlant="			& Trim(.txtConPlant.value)			
			iStrVal = iStrVal & "&txtConFromDt="		& Trim(.txtConFromDt.text)
			iStrVal = iStrVal & "&txtConToDt="			& Trim(.txtConToDt.text)		
			iStrVal = iStrVal & "&txtConDnType="		& Trim(.txtConDnType.value)			
			iStrVal = iStrVal & "&txtConShipToParty="	& Trim(.txtConShipToParty.value)
			
			If .rdoWorkTypeCfm.checked Then
				iStrVal = iStrVal & "&txtConPostFlag=Y"
				
				' Column Title 변경 
				.vspdData.Row = 0	:	.vspdData.Col = C_PromiseDt		: .vspdData.Text = "출고예정일"
				Call ggoOper.SetReqAttr(.txtActualGiDt,"N")
				Call ggoOper.SetReqAttr(.chkArFlag,"D")
				Call ggoOper.SetReqAttr(.chkVatFlag,"D")
			Else
				iStrVal = iStrVal & "&txtConPostFlag=N"

				' Column Title 변경 
				.vspdData.Row = 0	:	.vspdData.Col = C_PromiseDt		: .vspdData.Text = "실제출고일"
				Call ggoOper.SetReqAttr(.txtActualGiDt,"Q")
				Call ggoOper.SetReqAttr(.chkArFlag,"Q")
				Call ggoOper.SetReqAttr(.chkVatFlag,"Q")
			End If
			
			iStrVal = iStrVal & "&lgStrPrevKey="
		End if
		
		If .chkBatchQuery.checked Then
			iStrVal = iStrVal & "&txtBatchQuery=Y"
		Else
			iStrVal = iStrVal & "&txtBatchQuery=N"
		End If
		
		iStrVal = iStrVal & "&txtLastRow=" & .vspdData.MaxRows
		
    End With

	Call RunMyBizASP(MyBizASP, iStrVal)											
               
    If Err.number = 0 Then	 
       DbQuery = True                                                           
    End If

    Set gActiveElement = document.ActiveElement    
    
End Function

'=====================================================
Function DbQueryOk()
    lgIntFlgMode = Parent.OPMD_UMODE
	lgBlnFlgChgValue = False
	Call SetToolbar("11001001000111")	

	frm1.vspdData.focus
End Function

'=====================================================
Function DbSave() 
	On Error Resume Next

    Err.Clear																

    DbSave = False      
    
	If LayerShowHide(1) = False Then Exit Function 

	Dim iIntRow
	Dim iArrColData
	Dim iStrIns
	
	Dim iColSep, iRowSep, iFormLimitByte, iChunkArrayCount
	Dim iLngCTotalvalLen		'버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장 

	Dim iTmpCBuffer				'현재의 버퍼 [수정,신규] 
	Dim iTmpCBufferCount		'현재의 버퍼 Position
	Dim iTmpCBufferMaxCount		'현재의 버퍼 Chunk Size

	' 속도 향상을 위해 Local 변수로 재정의 
	iColSep = parent.gColSep
	iRowSep = parent.gRowSep
	iFormLimitByte = parent.C_FORM_LIMIT_BYTE
	iChunkArrayCount = parent.C_CHUNK_ARRAY_COUNT
	
	iTmpCBufferMaxCount = iChunkArrayCount '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpCBufferCount = -1
	iLngCTotalvalLen = 0
	
	ReDim iTmpCBuffer(iTmpCBufferMaxCount)	'최기 버퍼의 설정[신규]
	Redim iArrColData(2)

	With frm1.vspdData
		'-----------------------
		'Data manipulate area
		'-----------------------
		For iIntRow = 1 To .MaxRows
			.Row = iIntRow	:	.Col = C_Select
			If .Text = "1" Then
				iArrColData(0) = iIntRow									' Row 번호 
				.Col = C_DnNo			: iArrColData(1) = Trim(.Text)		' 출고번호 
				.Col = C_ExceptDnFlag	: iArrColData(2) = Trim(.Text)		' 예외출고여부 
				
				iStrIns = Join(iArrColData, iColSep) & iRowSep

				If iLngCTotalvalLen + Len(iStrIns) >  iFormLimitByte Then			'한개의 form element에 넣을 Data 한개치가 넘으면 
					Call MakeTextArea("txtCSpread", iTmpCBuffer)
								
				   iTmpCBufferMaxCount = iChunkArrayCount			                ' 임시 영역 새로 초기화 
				   ReDim iTmpCBuffer(iTmpCBufferMaxCount)
				   iTmpCBufferCount = -1
				   iLngCTotalvalLen  = 0
				End If
							   
				iTmpCBufferCount = iTmpCBufferCount + 1
							  
				If iTmpCBufferCount > iTmpCBufferMaxCount Then                      ' 버퍼의 조정 증가치를 넘으면 
				   iTmpCBufferMaxCount = iTmpCBufferMaxCount + iChunkArrayCount		' 버퍼 크기 증성 
				   ReDim Preserve iTmpCBuffer(iTmpCBufferMaxCount)
				End If   
				iTmpCBuffer(iTmpCBufferCount) =  iStrIns         
				iLngCTotalvalLen = iLngCTotalvalLen + Len(iStrIns)
			End If
		Next
	End With

   ' 나머지 데이터 처리 
	If iTmpCBufferCount > -1 Then Call MakeTextArea("txtCSpread", iTmpCBuffer)

	With frm1
		.txtMode.value = Parent.UID_M0002
		
		' 후속 작업여부 설정(매출채권)
		If .chkArFlag.checked Then
			.txtHArflag.value = "Y"
		Else
			.txtHArflag.value = "N"
		End If
		
		' 후속 작업여부 설정(세금계산서)
		If .chkVatFlag.checked Then
			.txtHVatFlag.value = "Y"
		Else
			.txtHVatFlag.value = "N"
		End If
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	
    DbSave = True                                                           
    
End Function

'=====================================================
Function DbSaveOk()
	
	Call ggoSpread.ClearSpreadData()
    Call InitVariables
    Call RemovedivTextArea
    Call MainQuery()

End Function

'========================================
Sub MakeTextArea(ByVal pvStrName, ByRef prArrData)
	Dim iObjTEXTAREA		'동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Set iObjTEXTAREA = document.createElement("TEXTAREA")            '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
	iObjTEXTAREA.name = pvStrName
	iObjTEXTAREA.value = Join(prArrData,"")
	divTextArea.appendChild(iObjTEXTAREA)
End Sub

'========================================
Function RemovedivTextArea()
	Dim iIntIndex
	
	For iIntIndex = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
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
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>일괄출고처리</font></td>
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
									<TD CLASS=TD5 NOWRAP>작업유형</TD>
								    <TD CLASS=TD6><INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoWorkType" TAG="11X" VALUE="Y" CHECKED ID="rdoWorkTypeCfm"><LABEL FOR="rdoWorkTypeCfm">출고처리</LABEL>&nbsp;
								                  <INPUT TYPE="RADIO" CLASS="RADIO" NAME="rdoWorkType" TAG="11X" VALUE="N" ID="rdoWorkTypeCancel"><LABEL FOR="rdoWorkTypeCancel">출고처리취소</LABEL></TD>
									<TD CLASS=TD5 NOWRAP>일괄조회</TD>
									<TD CLASS=TD6>
										<INPUT TYPE=CHECKBOX NAME="chkBatchQuery" ID="chkBatchQuery" tag="11" Class="Check">
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>공장</TD>
									<TD CLASS=TD6><INPUT NAME="txtConPlant" TYPE="Text" Alt="공장" MAXLENGTH=4 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopPlant">&nbsp;<INPUT NAME="txtConPlantNm" TYPE="Text" MAXLENGTH="20" Alt="공장명" SIZE=25 tag="14"></TD>									
									<TD CLASS="TD5" ID=idDtTitle NOWRAP>출고예정일</TD>
									<TD CLASS="TD6">
										<TABLE CELLSPACING=0 CELLPADDING=0>
											<TR>
												<TD>
													<script language =javascript src='./js/s4111ma6_fpDateTime1_txtConFromDt.js'></script>
												</TD>
												<TD>
													&nbsp;~&nbsp;
												</TD>
												<TD>
													<script language =javascript src='./js/s4111ma6_fpDateTime2_txtConToDt.js'></script>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>출하형태</TD>
									<TD CLASS=TD6 ><INPUT NAME="txtConDnType" TYPE="Text" MAXLENGTH="3" SIZE=10 Alt="출하형태" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopDnType">&nbsp;<INPUT NAME="txtConDnTypeNm" TYPE="Text" Alt="출하형태명" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>납품처</TD>
									<TD CLASS=TD6><INPUT NAME="txtConShipToParty" TYPE="Text" Alt="납품처" MAXLENGTH=10 SiZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenPopUp" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConPopUp C_PopShipToParty">&nbsp;<INPUT NAME="txtConShipToPartyNm" TYPE="Text" MAXLENGTH="20" Alt="납품처명" SIZE=25 tag="14"></TD>									
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>실제출고일</TD>
									<TD CLASS=TD6 ><script language =javascript src='./js/s4111ma6_fpDateTime1_txtActualGIDt.js'></script></TD>
									<TD CLASS=TD5 NOWRAP>후속작업여부</TD>
									<TD CLASS=TD6 >
										<INPUT TYPE=CHECKBOX NAME="chkArFlag" tag="21" Class="Check"><LABEL ID="lblArFlag" FOR="chkArFlag">매출채권</LABEL>&nbsp;&nbsp;
										<INPUT TYPE=CHECKBOX NAME="chkVatFlag" tag="21" Class="Check"><LABEL ID="lblVatFlag" FOR="chkVatFlag">세금계산서</LABEL>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>전체선택</TD>
									<TD CLASS=TD6 >
										<INPUT TYPE=CHECKBOX NAME="chkSelectAll" ID="chkSelectAll" tag="21" Class="Check">
									</TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 ></TD>
								</TR>
								<TR>
									<TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
										<script language =javascript src='./js/s4111ma6_vaSpread1_vspdData.js'></script>
									</TD>
								</TR>
							</TABLE>					
						</DIV>						
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME></TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConPlant" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConFromDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConToDt" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConDnType" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConShipToParty" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConPostFlag" tag="24" TABINDEX="-1">

<INPUT TYPE=HIDDEN NAME="txtHArFlag" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHVatFlag" tag="24" TABINDEX="-1">
<P ID="divTextArea"></P>
</FORM>
  <DIV ID="MousePT" NAME="MousePT">
  	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
  </DIV>
</BODY>
</HTML>
