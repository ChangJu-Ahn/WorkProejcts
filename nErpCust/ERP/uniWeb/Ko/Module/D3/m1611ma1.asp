<%@ LANGUAGE="VBSCRIPT" %>
<!--
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m1611ma1.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 매입형태구성정보 ASP														*
'*  6. Comproxy List        : +																			*
'*  7. Modified date(First) : 2003/06/02																*
'*  8. Modified date(Last)  :																			*
'*  9. Modifier (First)     : Yoon Ji Young																*
'* 10. Modifier (Last)      : Kim Jin Ha																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"									*
'*                            this mark(☆) Means that "must change"									*
'* 13. History              : 1. 2000/03/20 : 화면 design												*
'*         2. 2000/07/12 : Coding ReStart																*
'*                          *
'********************************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit   
'=======================================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'=======================================================================================================================
Dim interface_Account

Const BIZ_PGM_ID = "m1611mb1.asp" 

Dim  M_IvType				'매입형태 
Dim  M_IvTypeNm				'매입형태명    
Dim  M_ImportFlg			'수입여부 
Dim  M_ExceptFlg			'예외여부 
Dim  M_RetFlg				'반품여부 
'==== 2005.06.22 재고반영여부 추가 ==========
Dim	 M_StockFlg				'재고반영여부 
'==== 2005.06.22 재고반영여부 추가 ==========
Dim  M_UsageFlg				'사용여부 
Dim  M_TransType			'회계처리형태 
Dim  M_TransTypePopUp       '회계처리형태 
Dim  M_TransTypeNm			'회계처리형태명 

Dim lsBtnProtectedRow
Dim lgQuery
Dim lgCopyRow

Dim gblnWinEvent    
'=======================================================================================================================
Sub InitVariables()
    lgIntFlgMode = parent.OPMD_CMODE   
    lgBlnFlgChgValue = False    
    lgIntGrpCount = 0           

    lgStrPrevKey = ""
    lgLngCurRows = 0  
End Sub
'=======================================================================================================================
Sub initSpreadPosVariables()  
	M_IvType		= 1		'매입형태 
	M_IvTypeNm		= 2		'매입형태명    
	M_ImportFlg		= 3     '수입여부 
	M_ExceptFlg		= 4     '예외여부 
	M_RetFlg		= 5		'반품여부 
'==== 2005.06.22 재고반영여부 추가 ==========	
	M_StockFlg		= 6		'재고반영여부 
'==== 2005.06.22 재고반영여부 추가 ==========	
	M_UsageFlg		= 7		'사용여부 
	M_TransType     = 8     '회계처리형태 
	M_TransTypePopUp = 9    '회계처리형태 
	M_TransTypeNm   = 10		'회계처리형태명 
End Sub
'=======================================================================================================================
Sub SetDefaultVal()
	frm1.txtIvType.focus
	Set gActiveElement = document.activeElement
	Call SetToolbar("1110110100101111")
	interface_Account = GetSetupMod(parent.gSetupMod, "a")
End Sub
'=======================================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub
'=======================================================================================================================
 Sub InitSpreadSheet()
	
	Call initSpreadPosVariables() 
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.Spreadinit "V20030602",,parent.gAllowDragDropSpread       
	
	With frm1.vspdData

		.ReDraw = false
		
		.MaxCols = M_TransTypeNm + 1 
		.Col = .MaxCols
		.MaxRows = 0
		

		Call GetSpreadColumnPos("A")
		
		ggoSpread.SSSetEdit		M_IvType,		"매입형태", 15,,,5,2
		ggoSpread.SSSetEdit		M_IvTypeNm,		"매입형태명", 30,,,50
		ggoSpread.SSSetCheck	M_ImportFlg,	"수입여부", 15,,,true
		ggoSpread.SSSetCheck	M_ExceptFlg,	"예외여부", 15,,,true
		ggoSpread.SSSetCheck	M_RetFlg,		"반품여부", 15,,,true
		'==== 2005.06.22 재고반영여부 추가 ==========			
		ggoSpread.SSSetCheck	M_StockFlg,		"재고반영여부", 15,,,true
		'==== 2005.06.22 재고반영여부 추가 ==========	
		ggoSpread.SSSetCheck	M_UsageFlg,		"사용여부", 15,,,true
		ggoSpread.SSSetEdit		M_TransType,	"회계처리형태", 15,,,20,2
		ggoSpread.SSSetButton	M_TransTypePopUp
		ggoSpread.SSSetEdit		M_TransTypeNm,	"회계처리형태명", 20,,,,2
		
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		If interface_Account = "N" then
			Call ggoSpread.SSSetColHidden(.M_TransType, .M_TransType, True)
			Call ggoSpread.SSSetColHidden(.M_TransTypePopUp, .M_TransTypePopUp, True)
			Call ggoSpread.SSSetColHidden(.M_TransTypeNm, .M_TransTypeNm, True)
		End if
		
		Call ggoSpread.SSSetColHidden(M_RetFlg, M_RetFlg, True)
		Call ggoSpread.MakePairsColumn(M_TransType,M_TransTypePopUp)
		Call ggoSpread.SSSetSplit2(2)
		Call SetSpreadLock("", 0, -1, "")
		
		.ReDraw = true
	End With
    
End Sub
'=======================================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			M_IvType		= iCurColumnPos(1)		'매입형태 
			M_IvTypeNm		= iCurColumnPos(2)		'매입형태명    
			M_ImportFlg		= iCurColumnPos(3)		'수입여부 
			M_ExceptFlg		= iCurColumnPos(4)		'예외여부 
			M_RetFlg		= iCurColumnPos(5)		'반품여부 
			'==== 2005.06.22 재고반영여부 추가 ==========	
			M_StockFlg		= iCurColumnPos(6)		'재고반영여부 
			'==== 2005.06.22 재고반영여부 추가 ==========	
			M_UsageFlg		= iCurColumnPos(7)		'사용여부 
			M_TransType     = iCurColumnPos(8)		'회계처리형태 
			M_TransTypePopUp = iCurColumnPos(9)		'회계처리형태 
			M_TransTypeNm   = iCurColumnPos(10)		'회계처리형태명 
    End Select    
End Sub
'=======================================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
    With frm1
		ggoSpread.Source = .vspdData

		.vspdData.ReDraw = False    

		ggoSpread.SSSetrequired M_IvType, lRow, -1
		ggoSpread.SSSetrequired M_IvTypeNm, lRow, -1
		ggoSpread.SSSetrequired M_TransType, lRow, -1
		ggoSpread.SpreadLock M_TransTypeNm, lRow, -1
		ggoSpread.SSSetProtected frm1.vspdData.MaxCols, lRow, -1
		      
		.vspdData.ReDraw = True
    
    End With
  
End Sub
'=======================================================================================================================
Sub SetSpreadColor(ByVal pvStarRow, Byval pvEndRow)
    ggoSpread.Source = frm1.vspdData
    With ggoSpread
		frm1.vspdData.ReDraw = False
		.SSSetrequired M_IvType, pvStarRow,pvEndRow 
		.SSSetrequired M_IvTypeNm, pvStarRow,pvEndRow
		.SSSetrequired M_TransType, pvStarRow, pvEndRow
		'==== 2005.06.22 재고반영여부 추가 ==========
		.SSSetProtected M_Stockflg, pvStarRow, pvEndRow
		'==== 2005.06.22 재고반영여부 추가 ==========
		.SSSetProtected M_TransTypeNm, pvStarRow, pvEndRow
		.SSSetProtected frm1.vspdData.MaxCols, pvStarRow, pvEndRow
		frm1.vspdData.ReDraw = True
	End With

End Sub
'=======================================================================================================================
Function OpenIvType()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	arrParam(0) = "매입형태"   
	arrParam(1) = "M_IV_TYPE"              
	arrParam(2) = UCase(Trim(frm1.txtIvType.value))
	' arrParam(3) = Trim(frm1.txtIvTypeNm.value)        
	arrParam(4) = ""          
	arrParam(5) = "매입형태"       
	 
	arrField(0) = "IV_TYPE_CD"       
	arrField(1) = "IV_TYPE_NM"       
	    
	arrHeader(0) = "매입형태"      
	arrHeader(1) = "매입형태명"      

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False

	If arrRet(0) = "" Then
		frm1.txtIvType.focus	
		Set gActiveElement = document.activeElement
		Exit Function
	Else
		frm1.txtIvType.value = arrRet(0)
		frm1.txtiVTypeNm.value = arrRet(1)  
		frm1.txtIvType.focus	
		Set gActiveElement = document.activeElement
	End If 
End Function
'=======================================================================================================================
Sub changeException(ByVal curRow)
	Dim iCurExceptFlag
		
	Call frm1.vspdData.GetText(M_ExceptFlg,	curRow, iCurExceptFlag)  
		
	If iCurExceptFlag <> "1" then
		ggoSpread.spreadunlock  M_RetFlg, curRow, M_RetFlg, curRow
		'==== 2005.06.22 재고반영여부 추가 ==========
		ggoSpread.SSSetProtected M_StockFlg, curRow, curRow
		Call frm1.vspdData.SetText(M_StockFlg,	curRow, "0")
		'==== 2005.06.22 재고반영여부 추가 ==========
	Else
		ggoSpread.SSSetProtected M_RetFlg, curRow, curRow
		'==== 2005.06.22 재고반영여부 추가 ==========
		ggoSpread.spreadunlock M_StockFlg, curRow, curRow
		'==== 2005.06.22 재고반영여부 추가 ==========
		Call frm1.vspdData.SetText(M_RetFlg,	curRow, "")
	End if
End Sub
'=======================================================================================================================
Function OpenTypePopup()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True
	  
	frm1.vspdData.Row = frm1.vspdData.ActiveRow
	frm1.vspdData.Col = M_TransType
	  
	arrParam(0) = "회계처리형태"     
	arrParam(1) = "A_ACCT_TRANS_TYPE"    
	arrParam(2) = UCase(Trim(frm1.vspdData.Text))
	arrParam(3) = ""        
	arrParam(4) = "MO_CD = " & FilterVar("M", "''", "S") & " "      
	arrParam(5) = "회계처리형태"     
	 
	arrField(0) = "TRANS_TYPE"      
	arrField(1) = "TRANS_NM"      
	    
	arrHeader(0) = "회계처리형태"    
	arrHeader(1) = "회계처리형태명"    

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	 
	gblnWinEvent = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call frm1.vspdData.SetText(M_TransType,		frm1.vspdData.ActiveRow,	arrRet(0))
		Call frm1.vspdData.SetText(M_TransTypeNm,	frm1.vspdData.ActiveRow,	arrRet(1))
		Call vspdData_Change(frm1.vspdData.Col, frm1.vspdData.Row) 
	End If 
End Function
'=======================================================================================================================
Sub Form_Load()
	call LoadInfTB19029
    Call ggoOper.LockField(Document, "N")                                   
	Call SetDefaultVal 
	Call InitSpreadSheet
    Call InitVariables
End Sub
'=======================================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	
	IF lgIntFlgMode <> Parent.OPMD_UMODE Then
		Call SetPopupMenuItemInf("1001111111")
	Else
		Call SetPopupMenuItemInf("1101111111")
	End If
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
	
	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
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

	frm1.vspdData.Row = Row   
	
End Sub
'=======================================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub
'=======================================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    
'=======================================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub
'=======================================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'=======================================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
	CALL DbQueryOk()
End Sub
'=======================================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'=======================================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
   Dim iColumnName
    
	If Row <= 0 Then
		Exit Sub
	End If
	
	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End if
End Sub
'=======================================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	Dim strType 
 
	If lgQuery = true then Exit Sub
	If lgCopyRow = true then Exit Sub
	 
	If lsBtnProtectedRow >= Row Then Exit Sub

	frm1.vspdData.ReDraw = False
	With frm1.vspdData 
	 
		ggoSpread.Source = frm1.vspdData
		   
		If Row > 0 And Col = M_TransTypePopUp Then
			Call OpenTypePopup()
		End If
	 
	End With
		 
	Select Case Col
		Case M_ExceptFlg
			Call changeException(frm1.vspdData.ActiveRow)
	End Select 

	frm1.vspdData.ReDraw = True
 
End Sub
'=======================================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	    
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col

	Call CheckMinNumSpread(frm1.vspdData, Col, Row)        
End Sub
'=======================================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    '☜: 재쿼리 체크 
		If lgStrPrevKey <> "" Then							
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub
'=======================================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
	FncQuery = False                                
	    
	Err.Clear                                       

	ggoSpread.Source = frm1.vspdData
	    
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
	Call ggoOper.ClearField(Document, "2")   
	Call InitVariables        

	If Not ChkField(Document, "1") Then    
		Exit Function
	End If

	If DbQuery = False Then Exit Function

	FncQuery = True         
    Set gActiveElement = document.ActiveElement       
End Function
'=======================================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                  
    
    If ggoSpread.SSCheckChange = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"X","X") 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "1")          
    Call ggoOper.ClearField(Document, "2")          
    Call ggoOper.LockField(Document, "N")           
    Call SetDefaultVal
    Call InitVariables        

    FncNew = True         
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncDelete() 
    
    Exit Function
    Err.Clear                                       
    
    FncDelete = False        
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002","X","X","X")
        Exit Function
    End If
    
    If DbDelete = False Then                        
       Exit Function                                
    End If
    
    Call ggoOper.ClearField(Document, "1")          
    Call ggoOper.ClearField(Document, "2")          
    
    FncDelete = True                                
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim index 
    
	FncSave = False                                 

	ggoSpread.Source = frm1.vspdData    

	Err.Clear                                       
	    
	If ggoSpread.SSCheckChange = False Then
		IntRetCD = DisplayMsgBox("900001","X","X","X")
		Exit Function
	End If

	If interface_Account = "N" then
	
		for index=1 to frm1.vspdData.MaxRows
			frm1.vspdData.Row = index
			frm1.vspdData.Col = 0
			If frm1.vspdData.Text = ggoSpread.InsertFlag Or frm1.vspdData.Text = ggoSpread.UpdateFlag Or frm1.vspdData.Text = ggoSpread.DeleteFlag Then
				Call frm1.vspdData.SetText(M_TransType,	index, "*")
			End if
		Next
	End if
	 
	If Not ChkField(Document, "2") Then             
		Exit Function
	End If

	ggoSpread.Source = frm1.vspdData                
	If Not ggoSpread.SSDefaultCheck Then    
		Exit Function
	End If
	     
	If DbSave = False Then Exit Function

	FncSave = True                                  
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncCopy() 
	If frm1.vspdData.Maxrows < 1 Then Exit Function
	
	lgCopyRow = True
	frm1.vspdData.ReDraw = False
    ggoSpread.Source = frm1.vspdData
    ggoSpread.CopyRow 
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    Call frm1.vspdData.SetText(M_IvType,	frm1.vspdData.ActiveRow, "")
    frm1.vspdData.ReDraw = True
	lgCopyRow = False
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncCancel() 
	if frm1.vspdData.Maxrows < 1 then exit function
    ggoSpread.Source = frm1.vspdData 
    ggoSpread.EditUndo  
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
	Dim IntRetCD
    Dim imRow, iRow
    
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

    If IsNumeric(Trim(pvRowCnt)) Then
		imRow = CInt(pvRowCnt)
    Else
		imRow = AskSpdSheetAddRowCount()
		
		If imRow = "" Then
			Exit Function
		End if
    End If
    
	With frm1
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow, imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		
		For iRow = .vspdData.ActiveRow to .vspdData.ActiveRow + imRow -1
			Call frm1.vspdData.SetText(M_UsageFlg,	iRow,	"1")
		Next
		.vspdData.ReDraw = True

	End With
	
	If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    Dim iDelRowCnt, i
    if frm1.vspdData.Maxrows < 1 then exit function
    
    frm1.vspdData.focus
    ggoSpread.Source = frm1.vspdData 
	lDelRows = ggoSpread.DeleteRow
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData
    Call parent.FncPrint()
    Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_MULTI, False)
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function FncExit()
	Dim IntRetCD
	 
	FncExit = False
	 
	ggoSpread.Source = frm1.vspdData
	    
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"X","X")  
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	    
	FncExit = True
	Set gActiveElement = document.ActiveElement   
End Function
'=======================================================================================================================
Function DbQuery() 

	Err.Clear                                                   

	DbQuery = False                                             

	If LayerShowHide(1) = False Then
		Exit Function
	End If 
	     
	Dim strVal
	      
	If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001   
		strVal = strVal & "&txtIvType=" & Trim(frm1.hdnIvType.value)
		strVal = strVal & "&rdoUsageFlg=" & Trim(frm1.hdnUseflg.value)  
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001   
		strVal = strVal & "&txtIvType=" & Trim(frm1.txtIvType.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey  
			
		if frm1.rdoUsageFlgAll.checked = True then
			strVal = strVal & "&rdoUsageFlg=" & ""
		elseif frm1.rdoUsageFlgYes.checked = True then
			strVal = strVal & "&rdoUsageFlg=" & "Y"
		else
			strVal = strVal & "&rdoUsageFlg=" & "N"
		end if
	End If    
	 
	Call RunMyBizASP(MyBizASP, strVal)       
	 
	DbQuery = True            

End Function
'=======================================================================================================================
Function DbQueryOk()           
 
	Dim index
	
	lgIntFlgMode = parent.OPMD_UMODE         
	  
	Call ggoOper.LockField(Document, "Q")      
	Call SetToolbar("1110111100111111")
	    
	frm1.vspdData.ReDraw = False
	ggoSpread.spreadlock  M_Ivtype, 1, M_Ivtype, frm1.vspdData.MaxRows
	 
	For index = 1 To frm1.vspdData.MaxRows
		Call changeException(index)
	Next
	 
	lgQuery = False
	  
	frm1.vspdData.ReDraw = True
End Function
'=======================================================================================================================
Function DbSave() 

    Err.Clear             
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
	Dim PvArr
	Dim iColSep
	
    DbSave = False                                              
    
	If LayerShowHide(1) = False Then
	     Exit Function
	End If 
	
	With frm1
	 .txtMode.value = parent.UID_M0002
	   
	 lGrpCnt = 0    
	 strVal = ""
	 iColSep = parent.gColSep
	 ReDim PvArr(0)  
		 
	 For lRow = 1 To .vspdData.MaxRows
	   
		.vspdData.Row = lRow
		.vspdData.Col = 0
		
		Select Case .vspdData.Text
			Case ggoSpread.InsertFlag, ggoSpread.UpdateFlag, ggoSpread.DeleteFlag 
				
				If  .vspdData.Text = ggoSpread.InsertFlag then
					strVal = "C" & iColSep    '☜: C=Create
				ElseIf  .vspdData.Text = ggoSpread.UpdateFlag then
					strVal = "U" & iColSep    '☜: U=Update
				Else
					strVal = "D" & iColSep    '☜: D=Delete
				End if
					
				strVal = strVal & lRow & iColSep
	  
				.vspdData.Col = M_IvType:		strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep
				.vspdData.Col = M_IvTypeNm:		strVal = strVal & Trim(.vspdData.Text) & iColSep
	      		.vspdData.Col = M_ImportFlg	           
				
				if Trim(.vspdData.Text) = "1" then	
					strVal = strVal & "Y" & iColSep
				else
				strVal = strVal & "N" & iColSep
				end if

	           .vspdData.Col = M_ExceptFlg           
				if Trim(.vspdData.Text) = "1" then
					strVal = strVal & "Y" & iColSep
				else
				strVal = strVal & "N" & iColSep
				end if

				.vspdData.Col = M_RetFlg           
				if Trim(.vspdData.Text) = "1" then
					strVal = strVal & "Y" & iColSep
				else
				strVal = strVal & "N" & iColSep
				end if

				.vspdData.Col = M_UsageFlg
				if Trim(.vspdData.Text) = "1" then             
				strVal = strVal & "Y" & iColSep
				else              
				strVal = strVal & "N" & iColSep
				end if              

				.vspdData.Col = M_TransType 
				if interface_Account = "N" then
				 strVal = strVal & "*" & iColSep
				else
				 strVal = strVal & UCase(Trim(.vspdData.Text)) & iColSep
				end if
				
				'==== 2005.06.22 여분필드 추가 ==========					
				strVal = strVal & iColSep & iColSep & iColSep & iColSep			
				'==== 2005.06.22 여분필드 추가 ==========					
				
				'==== 2005.06.22 재고반영여부 추가 ==========					           
				.vspdData.Col = M_StockFlg 
				if Trim(.vspdData.Text) = "1" then
					strVal = strVal & "Y" & parent.gRowSep 
				else
					strVal = strVal & "N" & parent.gRowSep 
				end if
				'==== 2005.06.22 재고반영여부 추가 ==========
				
				ReDim Preserve PvArr(lGrpCnt)
				PvArr(lGrpCnt) = strVal
				lGrpCnt = lGrpCnt + 1 
	   End Select    
	 Next
	 
	 .txtMaxRows.value = lGrpCnt -1
	 .txtSpread.value = Join(PvArr, "")
	 
	 Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '☜: 비지니스 ASP 를 가동 
 
	End With
 
    DbSave = True                                                           '⊙: Processing is NG
    
End Function
'=======================================================================================================================
Function DbSaveOk()            
    Call InitVariables
    frm1.vspdData.MaxRows = 0
    
    Call MainQuery()
End Function
'=======================================================================================================================
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" --> 
</HEAD>
<!-- 
'#########################################################################################################
'            6. Tag부 
'######################################################################################################### 
-->
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매입형태</font></td>
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
									<TD CLASS=TD5 NOWRAP>매입형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtIvType" ALT="매입형태" TYPE="Text" MAXLENGTH=5 SIZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnIvType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenIvType()">
										<INPUT NAME="txtIvTypeNm" TYPE="Text" MAXLENGTH=30 SIZE=20  tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>사용여부</TD>
									<TD CLASS=TD6 NOWRAP><input type=radio CLASS="RADIO" name="rdoUsageFlg" id="rdoUsageFlgAll" value="A" tag = "11" checked><label for="rdoUsageFlgAll">전체</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoUsageFlg" id="rdoUsageFlgYes" value="Y" tag = "11"><label for="rdoUsageFlgYes">사용</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoUsageFlg" id="rdoUsageFlgNo" value="N" tag = "11"><label for="rdoUsageFlgNo">미사용</label></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>  
							<TR>
								<TD HEIGHT="100%">
									<script language =javascript src='./js/m1611ma1_I617720144_vspdData.js'></script>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<tr>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnUseflg" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV> 
</BODY>
</HTML>
