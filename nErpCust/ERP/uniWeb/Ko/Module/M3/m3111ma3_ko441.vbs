
Option Explicit				

Const BIZ_PGM_ID = "m3111mb3_ko441.asp"					

Dim C_CfmFlg
Dim C_PoNo
Dim C_PoType
Dim C_PoTypeNm
Dim C_PoDt
Dim C_PoAmt	
Dim C_Curr
Dim C_SupplierCd
Dim C_SupplierNm


Dim IsOpenPop          

'==========================================   Selection()  ======================================
'	Name : Selection()
'	Description : 일괄선택버튼의 Event 합수 
'=========================================================================================================
Sub Selection()
	Dim index,Count
	
	frm1.vspdData.ReDraw = false
	
	Count = frm1.vspdData.MaxRows 
	
	For index = 1 to Count
		frm1.vspdData.Row = index
		frm1.vspdData.Col = C_CfmFlg
		
		If frm1.vspdData.Text = "1" Then
			frm1.vspdData.Text = "0"
		Else
			frm1.vspdData.Text = "1"
		End If
		
		frm1.vspdData.Col = 0 
		
		If ggoSpread.UpdateFlag = frm1.vspdData.Text Then    
			frm1.vspdData.Text=""
	    Else
	    	ggoSpread.UpdateRow Index                         '편집모드는 수정모드로 
		End If
	Next 
	
	frm1.vspdData.ReDraw = true
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE  
    lgBlnFlgChgValue = False   
    lgIntGrpCount = 0 
    lgStrPrevKey = ""          
    lgLngCurRows = 0           
    frm1.vspdData.MaxRows = 0
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()
	frm1.txtPur_Grp.focus
	frm1.rdoConfirmFlg(1).checked = true
	Set gActiveElement = document.activeElement
	frm1.txtPur_Grp.Value = Parent.gPurGrp
	Call SetToolbar("1110000000001111")
	frm1.txtFrDt.Text = StartDate
	frm1.txtToDt.Text = EndDate
	frm1.btnSelect.disabled = True
	frm1.btnDisSelect.disabled = True
	If lgPGCd <> "" then 
		Call ggoOper.SetReqAttr(frm1.txtPur_Grp, "Q") 
        	frm1.txtPur_Grp.value = lgPGCd
	End If
End Sub

'======================================== 2.2.3 InitSpreadPosVariables() ========================================
Sub InitSpreadPosVariables()
	C_CfmFlg	= 1
	C_PoNo		= 2
	C_PoType	= 3
	C_PoTypeNm	= 4
	C_PoDt		= 5 
	C_PoAmt		= 6
	C_Curr		= 7
	C_SupplierCd= 8
	C_SupplierNm= 9
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ========================================
Sub InitSpreadSheet()
	Call InitSpreadPosVariables()
 
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021118",,Parent.gAllowDragDropSpread  

		.ReDraw = false	
	
		.MaxCols = C_SupplierNm + 1									
		.Col = .MaxCols:    .ColHidden = True						
		.MaxRows = 0

		Call GetSpreadColumnPos("A")

		ggoSpread.SSSetCheck C_CfmFlg, "확정여부",10,,,true
		ggoSpread.SSSetEdit	C_PoNo,"발주번호",18
		ggoSpread.SSSetEdit	C_PoType, "발주형태", 10
		ggoSpread.SSSetEdit	C_PoTypeNm, "발주형태명",15
		ggoSpread.SSSetDate	C_PoDt,"발주일", 10, 2, Parent.gDateFormat
		SetSpreadFloat	 	C_PoAmt, "발주금액", 15, 1, 2
		ggoSpread.SSSetEdit	C_Curr, "화폐", 10
		ggoSpread.SSSetEdit	C_SupplierCd, "공급처", 10
		ggoSpread.SSSetEdit	C_SupplierNm, "공급처명", 20
	
		Call SetSpreadLock 
		.ReDraw = true
    End With
End Sub

'================================== 2.2.4 SetSpreadLock() ==================================================
Sub SetSpreadLock()
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SpreadLock -1, -1
    ggoSpread.spreadUnlock C_CfmFlg, -1,C_CfmFlg, -1
    .vspdData.ReDraw = True

    End With
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
	Dim iCurColumnPos
	
	Select Case UCase(pvSpdNo)
		Case "A"
			ggoSpread.Source = frm1.vspdData
			
			Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

			C_CfmFlg	= iCurColumnPos(1)
			C_PoNo		= iCurColumnPos(2)
			C_PoType	= iCurColumnPos(3)
			C_PoTypeNm	= iCurColumnPos(4)
			C_PoDt		= iCurColumnPos(5)
			C_PoAmt		= iCurColumnPos(6)
			C_Curr		= iCurColumnPos(7)
			C_SupplierCd= iCurColumnPos(8)
			C_SupplierNm= iCurColumnPos(9)

	End Select
End Sub	

'------------------------------------------  OpenSupplier()  -------------------------------------------------
Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공급처"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & " "				
	arrParam(5) = "공급처"					
	
    arrField(0) = "BP_Cd"				
    arrField(1) = "BP_NM"				
    
    arrHeader(0) = "공급처"			
    arrHeader(1) = "공급처명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		frm1.txtSupplierCd.focus
	End If	
End Function

'------------------------------------------  OpenPurGrp()  -------------------------------------------------
Function OpenPurGrp()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function
	If frm1.txtPur_Grp.className = "protected" Then Exit Function

	IsOpenPop = True

	arrParam(0) = "구매그룹"			
	arrParam(1) = "B_PUR_GRP"			

	arrParam(2) = Trim(frm1.txtPur_Grp.Value)		
'	arrParam(3) = Trim(frm1.txtPur_Grp_Nm.Value)	
	
	arrParam(4) = ""								
	arrParam(5) = "구매그룹"						
	
    arrField(0) = "PUR_GRP"							
    arrField(1) = "PUR_GRP_NM"						
    
    arrHeader(0) = "구매그룹"					
    arrHeader(1) = "구매그룹명"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPur_Grp.focus
		Exit Function
	Else
		frm1.txtPur_Grp.Value		= arrRet(0)		
		frm1.txtPur_Grp_Nm.Value	= arrRet(1)		
		frm1.txtPur_Grp.focus
	End If	
End Function

'==========================================================================================
'   Event Name : btnPosting_OnClick()
'   Event Desc : 출고처리 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnSelect_OnClick()
	Dim i
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_CfmFlg
			frm1.vspdData.Row = i
			frm1.vspdData.value = 1
			Call vspdData_ButtonClicked(C_CfmFlg, i, 1)
		Next	
		
	End If
End Sub

'==========================================================================================
'   Event Name : btnPostCancel_OnClick()
'   Event Desc : 출고처리취소 버튼을 클릭할 경우 발생 
'==========================================================================================
Sub btnDisSelect_OnClick()
	Dim i
	If frm1.vspdData.Maxrows > 0 then
	    ggoSpread.Source = frm1.vspdData

	    For i = 1 to frm1.vspdData.Maxrows
			frm1.vspdData.Col = C_CfmFlg
			frm1.vspdData.Row = i
			frm1.vspdData.value = 0

			Call vspdData_ButtonClicked(C_CfmFlg, i, 0)
		Next	
	End If
End Sub

'==========================================================================================
'   Event Name : txtFrDt
'==========================================================================================
Sub txtFrDt_DblClick(Button)
	if Button = 1 then
		frm1.txtFrDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtFrDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : txtToDt
'==========================================================================================
Sub txtToDt_DblClick(Button)
	if Button = 1 then
		frm1.txtToDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtToDt.Focus
	End if
End Sub

'==========================================================================================
'   Event Name : OCX_KeyDown()
'==========================================================================================
Sub txtFrDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

Sub txtToDt_KeyDown(KeyCode, Shift)
	If KeyCode = 13	Then Call MainQuery()
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	If frm1.vspdData.MaxRows > 0 Then
		Call SetPopupMenuItemInf("0000111111")
	Else
		Call SetPopupMenuItemInf("0000111111")
	End If   

	If frm1.vspdData.MaxRows = 0 Then Exit Sub
	   
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col				'Sort in Ascending
			lgSortkey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey	'Sort in Descending
			lgSortkey = 1
		End If
		
		Exit Sub
	End If    		
End Sub

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

   If Button = 2 And gMouseClickStatus = "SPC" Then
      gMouseClickStatus = "SPCR"
   End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
	End If
	If frm1.vspddata.MaxRows=0 Then
		Exit Sub
	End If

End Sub
'========================================================================================
' Function Name : FncSplitColumn
'========================================================================================
Sub FncSplitColumn()
    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   
End Sub
'========================================================================================
' Function Name : PopSaveSpreadColumnInf
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub
'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    Call ggoSpread.ReOrderingSpreadData()
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 버튼 컬럼을 클릭할 경우 발생 
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	If Col = C_CfmFlg And Row > 0 Then
	    Select Case ButtonDown
	    Case 1
			ggoSpread.Source = frm1.vspdData
			if Trim(frm1.hdnrdoflg.value) = "Y" then
			    frm1.vspdData.Col = 0
			    frm1.vspdData.Row = Row 
			    frm1.vspdData.text = "" 
			else 
			    ggoSpread.UpdateRow Row
			end if
			lgBlnFlgChgValue = True		
	    Case 0
			ggoSpread.Source = frm1.vspdData
			if Trim(frm1.hdnrdoflg.value) = "N" then
			    frm1.vspdData.Col = 0
			    frm1.vspdData.Row = Row 
			    frm1.vspdData.text = "" 
			else 
			    ggoSpread.UpdateRow Row
			end if
			lgBlnFlgChgValue = False					
	    End Select
	End If
End Sub
'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	
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
'========================================================================================
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    FncQuery = False                                        
    Err.Clear                                               

	ggoSpread.Source = frm1.vspdData
	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    Call InitVariables
    														    
	with frm1
		if (UniConvDateToYYYYMMDD(.txtFrDt.text,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(.txtToDt.text,Parent.gDateFormat,"")) and Trim(.txtFrDt.text)<>"" and Trim(.txtToDt.text)<>"" then	
			Call DisplayMsgBox("17a003", "X","발주일", "X")			
			Exit Function
		End if   
	End with
	
    If DbQuery = False Then Exit Function

	Set gActiveElement = document.activeElement
    FncQuery = True											
End Function

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                          
    
    Err.Clear                                               
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
       
    End If
    
    Call ggoOper.ClearField(Document, "1")                  
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
	
    Call LockObjectField(frm1.txtFrDt, "O")
    Call LockObjectField(frm1.txtToDt, "O")
    
    Call InitVariables                                      
    Call SetDefaultVal
        
	Set gActiveElement = document.activeElement
    FncNew = True                                           
End Function

'========================================================================================
' Function Name : FncDelete
'========================================================================================
Function FncDelete() 
    Dim IntRetCD 
    
    FncDelete = False                                       
    
    Err.Clear                                               
        
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                      
        Call DisplayMsgBox("900002", "X", "X", "X")         
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
    Call ggoOper.ClearField(Document, "1")                  
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
    
	Set gActiveElement = document.activeElement
    FncDelete = True                                        
End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                         
    
    Err.Clear         
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                                      
    
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = False Then                 
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")   
        Exit Function
    End If
    
    ggoSpread.Source = frm1.vspdData                        
    If Not ggoSpread.SSDefaultCheck         Then            
       Exit Function
    End If
    
    If DbSave = False Then Exit Function
    
	Set gActiveElement = document.activeElement
    FncSave = True                                         
End Function

'========================================================================================
' Function Name : FncCancel
'========================================================================================
Function FncCancel() 
	if frm1.vspdData.Maxrows < 1	then exit function
	ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                     
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncExport(Parent.C_MULTI)							
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind()
	ggoSpread.Source = frm1.vspdData 
    Call parent.FncFind(Parent.C_MULTI , False)                    
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData
    
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")  
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	Set gActiveElement = document.activeElement
    FncExit = True
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim LngRow          
    Dim strTemp         
    Dim StrNextKey      

    DbQuery = False
    
    If LayerShowHide(1) = False Then Exit Function
    
    Err.Clear                                                    

	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = Parent.OPMD_UMODE Then
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtSupplier=" & .hdnSupplier.value
	    strVal = strVal & "&txtGroup=" & .hdnGroup.value
		strVal = strVal & "&txtFrDt=" & .hdnFrDt.value
		strVal = strVal & "&txtToDt=" & .hdnToDt.value
		strVal = strVal & "&txtCfmflg=" & .hdnCfmflg.value
	else
	    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001
	    strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	    strVal = strVal & "&txtSupplier=" & Trim(.txtSuppliercd.value)
	    strVal = strVal & "&txtGroup=" & Trim(.txtPur_grp.value)
		strVal = strVal & "&txtFrDt=" & Trim(.txtFrDt.text)
		strVal = strVal & "&txtToDt=" & Trim(.txtToDt.text)
		if .rdoConfirmFlg(0).checked = true then
			strVal = strVal & "&txtCfmflg=" & "Y"
		else
			strVal = strVal & "&txtCfmflg=" & "N"
		end if
	end if 

	Call RunMyBizASP(MyBizASP, strVal)		
        
    End With
    
    DbQuery = True
End Function
'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()						
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE				
        
	Call SetSpreadLock
	Call SetToolbar("11101000000111")
	frm1.btnSelect.disabled = False
	frm1.btnDisSelect.disabled = False
	If frm1.rdoConfirmFlg(0).checked = true Then
	   frm1.hdnrdoflg.value = "Y"
	Else
	   frm1.hdnrdoflg.value = "N"
	End If
End Function

Sub RemovedivTextArea()
	Dim ii

	For ii = 1 To divTextArea.children.length
	    divTextArea.removeChild(divTextArea.children(0))
	Next
End Sub

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal
    Dim lColSep,lRowSep

	Dim strCUTotalvalLen '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[수정,신규] 
	Dim strDTotalvalLen  '버퍼에 채워지는 양이 102399byte를 넘어 가는가를 체크하기위한 누적 데이타 크기 저장[삭제]

	Dim objTEXTAREA '동적인 HTML객체(TEXTAREA)를 만들기위한 임시 버퍼 

	Dim iTmpCUBuffer         '현재의 버퍼 [수정,신규] 
	Dim iTmpCUBufferCount    '현재의 버퍼 Position
	Dim iTmpCUBufferMaxCount '현재의 버퍼 Chunk Size

	Dim iTmpDBuffer          '현재의 버퍼 [삭제] 
	Dim iTmpDBufferCount     '현재의 버퍼 Position
	Dim iTmpDBufferMaxCount  '현재의 버퍼 Chunk Size
    Dim ii
	
    DbSave = False    
    
    If LayerShowHide(1) = False Then Exit Function
    
	With frm1
		.txtMode.value = Parent.UID_M0002
    
    lColSep = parent.gColSep
    lRowSep = parent.gRowSep
    '-----------------------
    'Data manipulate area
    '-----------------------
    lGrpCnt = 1
    strVal = ""
	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[수정,신규]
	iTmpDBufferMaxCount  = parent.C_CHUNK_ARRAY_COUNT '한번에 설정한 버퍼의 크기 설정[삭제]

	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount) '최기 버퍼의 설정[수정,신규]
	ReDim iTmpDBuffer(iTmpDBufferMaxCount)  '최기 버퍼의 설정[수정,신규]

	iTmpCUBufferCount = -1
	iTmpDBufferCount = -1

	strCUTotalvalLen = 0
	strDTotalvalLen  = 0

    '-----------------------
    'Data manipulate area
    '-----------------------
	For lRow = 1 To .vspdData.MaxRows
			
        If Trim(GetSpreadText(.vspdData,0,lRow,"X","X")) = ggoSpread.UpdateFlag Then
	   				
			strVal = "U" & lColSep
			If Trim(GetSpreadText(.vspdData,C_CfmFlg,lRow,"X","X")) <> 0 Then
				strVal = strVal & "Y" & lColSep
			Else
				strVal = strVal & "N" & lColSep
			End If
			strVal = strVal & Trim(GetSpreadText(.vspdData,C_PoNo,lRow,"X","X")) & lColSep
			strVal = strVal & lRow & lRowSep
				
			lGrpCnt = lGrpCnt + 1
		End If
		
		Select Case Trim(GetSpreadText(.vspdData,0,lRow,"X","X"))
		    Case ggoSpread.UpdateFlag
		         If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then  '한개의 form element에 넣을 Data 한개치가 넘으면 
		                            
		            Set objTEXTAREA = document.createElement("TEXTAREA")                 '동적으로 한개의 form element를 동저으로 생성후 그곳에 데이타 넣음 
		            objTEXTAREA.name = "txtCUSpread"
		            objTEXTAREA.value = Join(iTmpCUBuffer,"")
		            divTextArea.appendChild(objTEXTAREA)     
		 
		            iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT                  ' 임시 영역 새로 초기화 
		            ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
		            iTmpCUBufferCount = -1
		            strCUTotalvalLen  = 0
		         End If
		       
		         iTmpCUBufferCount = iTmpCUBufferCount + 1
		      
		         If iTmpCUBufferCount > iTmpCUBufferMaxCount Then                              '버퍼의 조정 증가치를 넘으면 
		            iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT '버퍼 크기 증성 
		            ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
		         End If   
		         iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
		         strCUTotalvalLen = strCUTotalvalLen + Len(strVal)
		End Select   
	Next

	If iTmpCUBufferCount > -1 Then   ' 나머지 데이터 처리 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	End If   

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
	
	End With
	
    DbSave = True                       
End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()						
	Call InitVariables()
	Call MainQuery()
End Function

