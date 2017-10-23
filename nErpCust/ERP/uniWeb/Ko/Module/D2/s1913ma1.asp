<%@ LANGUAGE="VBSCRIPT" %>
<%
'********************************************************************************************************
'*  1. Module Name          : 영업																		*
'*  2. Function Name        :																			*
'*  3. Program ID           : s1913ma1.asp																*
'*  4. Program Name         : 매출채권형태등록															*
'*  5. Program Desc         : 매출채권형태등록															*
'*  6. Comproxy List        :  																			*
'*  7. Modified date(First) : 2000/08/25																*
'*  8. Modified date(Last)  : 2001/12/18																*
'*  9. Modifier (First)     : Juvenile	 																*
'* 10. Modifier (Last)      : Sonbumyeol	
'* 11. Comment              :																			*
'* 12. Comment              : 2002/11/27 : Grid성능 적용, Kang Jun Gu
'* 13. Comment              : 2002/12/02 : Grid성능 추가 적용, Kang Jun Gu
'********************************************************************************************************
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
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                                             
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UniConvDateAToB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)
'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UNIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "s1913mb1.asp"												
Const gstrDNTypeMajor = "I0001"

Dim S_BillType	
Dim S_BillTypeNm
Dim S_ExceptFlg	
Dim S_ExportFlg	
Dim S_GiFlg		
Dim S_TransType	
Dim S_TransTypePopup
Dim S_TransTypeNm	
Dim S_UsageFlg		
Dim S_AsFlg		
Dim S_ChgFlg	

Dim lsBtnClickFlag 
Dim lsBtnProtectedRow

Dim lsQueryMode               
Dim lsFncCopyFlag 
 
Dim IsOpenPop						

'========================================================================================================
Sub initSpreadPosVariables()  
	S_BillType			= 1		'매출형태 
	S_BillTypeNm		= 2		'매출형태명 
	S_ExceptFlg			= 3		'예외여부 
	S_ExportFlg			= 4		'수출여부 
	S_GiFlg				= 5     '출하여부 
	S_TransType			= 6		'회계거래유형 
	S_TransTypePopup	= 7		'회계거래유형팝업버튼	
	S_TransTypeNm		= 8		'회계거래유형명   
	S_UsageFlg			= 9		'사용여부 
	S_AsFlg				= 10	'AS여부 
	S_ChgFlg			= 11	'Sort시 필요한 히든 칼럼 

End Sub

'===============================================================================================================
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   
    lgBlnFlgChgValue = False                    
    lsFncCopyFlag = False
    lgIntGrpCount = 0                           
    '---- Coding part--------------------------------------------------------------------
    lgStrPrevKey = ""
    lgLngCurRows = 0  
End Sub

'===============================================================================================================
Sub SetDefaultVal()
	frm1.txtBillType.focus
End Sub

'===============================================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub


'===============================================================================================================
Sub InitSpreadSheet()
	Call initSpreadPosVariables()    	
    ggoSpread.Source = frm1.vspdData
	With frm1.vspdData
		ggoSpread.Spreadinit "V20030425",,parent.gAllowDragDropSpread            	    
		.ReDraw = false
	    .MaxCols = S_ChgFlg									'☜: 최대 Columns의 항상 1개 증가시킴 
	    .Col = .MaxCols											'☜: 공통콘트롤 사용 Hidden Column
		.ColHidden = True
	    .MaxRows = 0
	    Call GetSpreadColumnPos("A")
		
	    
		ggoSpread.SSSetEdit S_BillType, "매출채권형태", 15,,,4, 2	
						'ColumnPosition		Header              Width	Align(0:L,1:R,2:C)  Row   Length  CharCase(0:L,1:N,2:U)	
	    ggoSpread.SSSetEdit S_BillTypeNm, "매출채권형태명", 20,		,					,	  30
		ggoSpread.SSSetCheck S_ExportFlg, "수출여부", 15,,,true
		ggoSpread.SSSetCheck S_ExceptFlg, "예외여부", 15,,,true
		ggoSpread.SSSetCheck S_GiFlg, "출하여부", 15,,,true
	    ggoSpread.SSSetEdit S_TransType, "회계거래유형", 15,,,20, 2
		ggoSpread.SSSetButton S_TransTypePopup
		ggoSpread.SSSetEdit S_TransTypeNm, "회계거래유형명", 20, 0
		ggoSpread.SSSetCheck S_UsageFlg, "사용여부", 15,,,true		

		ggoSpread.SSSetCheck S_AsFlg, "A/S여부", 15,,,true		
		
'		call ggoSpread.MakePairsColumn(S_BillType,S_BillTypeNm)
		call ggoSpread.MakePairsColumn(S_TransType,S_TransTypePopup)		
		Call ggoSpread.SSSetColHidden(S_ChgFlg,S_ChgFlg,True)		
		

		If GetSetupMod(Parent.gSetupMod, "A") = "N" Then
			Call ggoSpread.SSSetColHidden(S_TransType,S_TransType,True)		
			Call ggoSpread.SSSetColHidden(S_TransTypePopup,S_TransTypePopup,True)		
			Call ggoSpread.SSSetColHidden(S_TransTypeNm,S_TransTypeNm,True)		
		End If
		
	    Call SetSpreadLock("", 0, -1, "")
	    .ReDraw = True
    End With
End Sub


'===============================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
    With frm1
		ggoSpread.Source = .vspdData

		.vspdData.ReDraw = False    

		ggoSpread.SSSetrequired S_BillType, lRow, -1
		ggoSpread.SSSetrequired S_BillTypeNm, lRow, -1
		ggoSpread.SSSetrequired S_TransType, lRow, -1
		ggoSpread.SpreadLock S_TransTypeNm,lRow, -1
		.vspdData.ReDraw = True
    End With
End Sub


'===============================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False 
		ggoSpread.SSSetRequired S_BillType, pvStartRow,pvEndRow
		ggoSpread.SSSetRequired S_BillTypeNm, pvStartRow,pvEndRow
		ggoSpread.SSSetRequired S_TransType, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected S_TransTypeNm, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    End With
End Sub

'===============================================================================================================
Sub SetInsertSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
		.vspdData.ReDraw = False 
		ggoSpread.SSSetRequired S_BillType, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired S_BillTypeNm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired S_TransType, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected S_TransTypeNm, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
    End With
End Sub

'===============================================================================================================
Sub SetQuerySpreadColor(ByVal lRow)
	With frm1
	    ggoSpread.SSSetProtected S_BillType, lRow,lRow
		ggoSpread.SSSetRequired S_TransType, lRow, lRow
    End With
End Sub

'===============================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				S_BillType			= iCurColumnPos(1)		'매출형태 
				S_BillTypeNm		= iCurColumnPos(2)		'매출형태명 
				S_ExceptFlg			= iCurColumnPos(3)		'예외여부 
				S_ExportFlg			= iCurColumnPos(4)		'수출여부 
				S_GiFlg				= iCurColumnPos(5)     '출하여부 
				S_TransType			= iCurColumnPos(6)		'회계거래유형 
				S_TransTypePopup	= iCurColumnPos(7)		'회계거래유형팝업버튼	
				S_TransTypeNm		= iCurColumnPos(8)		'회계거래유형명   
				S_UsageFlg			= iCurColumnPos(9)		'사용여부 
				S_ASFlg				= iCurColumnPos(10)		'Sort시 필요한 히든 칼럼 
				S_ChgFlg			= iCurColumnPos(11)		'Sort시 필요한 히든 칼럼 
    End Select    
End Sub


'===============================================================================================================
Function OpenCondtionPopup()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "매출채권형태"						
	arrParam(1) = "S_BILL_TYPE_CONFIG"						
	arrParam(2) = Trim(frm1.txtBillType.value)  			
	arrParam(4) = ""										
	arrParam(5) = "매출채권형태"						
	
	arrField(0) = "BILL_TYPE"								
	arrField(1) = "BILL_TYPE_NM"							
    
	arrHeader(0) = "매출채권형태"						
	arrHeader(1) = "매출채권형태명"						

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCondtionPopup(arrRet)
	End If		
End Function

'===============================================================================================================
Function OpenTypePopup(Byval strCode)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "회계거래유형"
	arrParam(1) = "a_acct_trans_type"					
	arrParam(2) = strCode								
	arrParam(3) = ""									
	arrParam(4) = "mo_cd = " & FilterVar("S", "''", "S") & ""							
	arrParam(5) = "회계거래유형"					
	
	arrField(0) = "TRANS_TYPE"							
	arrField(1) = "TRANS_NM"							

	arrHeader(0) = "회계거래유형"					
	arrHeader(1) = "회계거래유형명"					

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetTypePopup(arrRet)
	End If	
	
End Function
'===============================================================================================================
Function SetCondtionPopup(Byval arrRet)
	frm1.txtBillType.value = arrRet(0)
	frm1.txtBillTypeNm.value = arrRet(1)  
	frm1.txtBillType.focus
End Function
'===============================================================================================================
Function SetTypePopup(Byval arrRet)
	With frm1
		.vspdData.Col = S_TransType
		.vspdData.Text = arrRet(0)
		.vspdData.Col = S_TransTypeNm
		.vspdData.Text = arrRet(1)

		Call vspdData_Change(.vspdData.Col, .vspdData.Row)		
	End With

	lgBlnFlgChgValue = True
End Function

'===============================================================================================================
Function SetRadio()
	frm1.rdoUsageFlgAll.checked = True
End Function

'===============================================================================================================
Sub Form_Load()

	Call InitVariables														
    Call LoadInfTB19029														
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)  '⊙: Format Contents  Field
    Call ggoOper.LockField(Document, "N")                                   
    
	Call InitSpreadSheet
	Call SetDefaultVal	
    Call SetToolbar("1110110100101111")										
End Sub

'===============================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	If lsQueryMode = False Then Exit Sub
	If lsFncCopyFlag = True Then Exit Sub
    ggoSpread.Source = frm1.vspdData

	lgBlnFlgChgValue = True

	With frm1.vspdData 
		If Row > 0 And Col = S_TransTypePopup Then
		    .Col = Col - 1
		    .Row = Row
		    Call OpenTypePopup(.Text)
		    Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")		    
		End If
	End With
	

End Sub
'===============================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True
End Sub

'===============================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	Call SetPopupMenuItemInf("1111111111")
    
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    If frm1.vspdData.MaxRows <= 0 Then                                                   
       Exit Sub
   	End If
   	
	If Row <= 0 Then
	    ggoSpread.Source = frm1.vspdData
	    If lgSortKey = 1 Then
	        ggoSpread.SSSort Col				'Sort in Ascending
	        lgSortKey = 2
	    Else
	        ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
	        lgSortKey = 1
	    End If
		 Exit Sub     
	End If   	 	    
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    	frm1.vspdData.Row = Row
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
End Sub

'===============================================================================================================

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
'===============================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
	
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
    	Exit Sub
    End If
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    
	
End Sub
'===============================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub


'===============================================================================================================

Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub


'===============================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'===============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    
		If CheckRunningBizProcess = True Then Exit Sub	
    	If lgPageNo <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub
'===============================================================================================================
Sub rdoUsageFlgAll_OnClick()
	frm1.txtRadio.value = frm1.rdoUsageFlgAll.value
End Sub

Sub rdoUsageFlgYes_OnClick()
	frm1.txtRadio.value = frm1.rdoUsageFlgYes.value
End Sub

Sub rdoUsageFlgNo_OnClick()
	frm1.txtRadio.value = frm1.rdoUsageFlgNo.value
End Sub

'===============================================================================================================
Function FncQuery() 
	Dim IntRetCD 
	
    Err.Clear  
    
    FncQuery = False


    ggoSpread.source = frm1.vspdData   
    If ggoSpread.SSCheckChange Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
         						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData
    Call InitVariables														


    If Not chkField(Document, "1") Then									
       Exit Function
    End If
 
	Call DbQuery                                                      
    
    If Err.number = 0 Then	
       FncQuery = True      
       
    End If
	
    Set gActiveElement = document.ActiveElement   
    
End Function

'===============================================================================================================
Function FncNew() 
    Dim IntRetCD 
    On Error Resume Next
    Err.Clear           
        
    FncNew = False                                                          
    

    ggoSpread.source = frm1.vspdData   
    If ggoSpread.SSCheckChange Then
        IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x") 
		
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    

    Call ggoOper.ClearField(Document, "A")                                      	
    Call ggoOper.LockField(Document, "N")                                       
    Call InitVariables															
	Call SetRadio
    Call SetToolbar("11101101001111")										
    Call SetDefaultVal

    If Err.number = 0 Then	
       FncNew = True                                                              
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'===============================================================================================================
Function FncDelete() 
    
    Exit Function
    
    On Error Resume Next
    Err.Clear           
    
    
    FncDelete = False														
    

    If lgIntFlgMode <> Parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
    

    If DbDelete = False Then                                                
       Exit Function                                                       
    End If
    
    Call ggoOper.ClearField(Document, "A")                                         
	
    If Err.number = 0 Then	
       FncDelete = True                                                           
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function

'===============================================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim i
    
    On Error Resume Next
    Err.Clear           
        
    FncSave = False                                                         

	
    ggoSpread.source = frm1.vspdData   
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")        
        Exit Function
    End If

	If GetSetupMod(Parent.gSetupMod, "A") = "N" Then
		For i=1 To frm1.vspdData.MaxRows
			frm1.vspdData.Row = i
			frm1.vspdData.Col = S_TransType 
			frm1.vspdData.text = "*"
		Next
	End If
    
    If Not chkField(Document, "2") Then     
       Exit Function
    End If

    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then                                          '☜: Check contents area
       Exit Function
    End If
    
    If DbSave = False Then                                                        '☜: Query db data
       Exit Function
    End If

    If Err.number = 0 Then	
       FncSave = True                                                             
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function

'========================================================================================
Function FncCopy() 
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If    
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetInsertSpreadColor .ActiveRow, .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
	lsFncCopyFlag = True    
    If Err.number = 0 Then	
       FncCopy = True                                                            
    End If

    Set gActiveElement = document.ActiveElement   		
	
End Function


'===============================================================================================================
Function FncCancel() 
	Dim iDx

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG    
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
        
    If Err.number = 0 Then	
       FncCancel = True                                                            
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'===============================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim lngRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG
    If IsNumeric(Trim(pvRowCnt)) Then
        imRow = CInt(pvRowCnt)
    Else
        imRow = AskSpdSheetAddRowCount()
        If imRow = "" Then
            Exit Function
        End If
    End If
        
    
	With frm1
		lsQueryMode = True
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData		

		ggoSpread.InsertRow,imRow

		For lngRow =  .vspdData.ActiveRow to .vspdData.ActiveRow + imRow - 1
			.vspdData.Col = S_UsageFlg 
			.vspdData.Row = lngRow
			.vspdData.Text = "1"  
		Next

		 SetInsertSpreadColor  .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

		.vspdData.ReDraw = True
		lgBlnFlgChgValue = True
   End With
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          
    End If   
	
    Set gActiveElement = document.ActiveElement      
End Function

'===============================================================================================================
Function FncDeleteRow() 
	Dim lDelRows
	
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncDeleteRow = False                                                          '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
	
	With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    lgBlnFlgChgValue = True
	
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            
    End If

    Set gActiveElement = document.ActiveElement       
End Function


'===============================================================================================================
Function FncPrint() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        

    If Err.number = 0 Then	 
       FncPrint = True                                                            
    End If

    Set gActiveElement = document.ActiveElement   
    

End Function

'===============================================================================================================
Function FncPrev() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrev = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------

    If Err.number = 0 Then	 
       FncPrev = True                                                             
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'===============================================================================================================
Function FncNext() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncNext = False                                                               '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
    If Err.number = 0 Then	 
       FncNext = True                                                             
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'===============================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call parent.FncExport(Parent.C_SINGLEMULTI)

    If Err.number = 0 Then	 
       FncExcel = True                                                            
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'===============================================================================================================
Function FncFind() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(Parent.C_SINGLEMULTI, False)

    If Err.number = 0 Then	 
       FncFind = True                                                             
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function


'===============================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub
'===============================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'===============================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call SetQuerySpreadColor (-1)
End Sub


'===============================================================================================================
Function FncExit()    
	Dim IntRetCD
	
	On Error Resume Next 
    Err.Clear 
    	
	FncExit = False

    ggoSpread.source = frm1.vspdData   
    If ggoSpread.SSCheckChange Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    If Err.number = 0 Then	 
       FncExit = True                                                             
    End If

    Set gActiveElement = document.ActiveElement   
End Function

'===============================================================================================================
Function DbDelete() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    DbDelete = False                                                              '☜: Processing is NG
	'------ Developer Coding part (Start)  -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	 
       DbDelete = True                                                             
    End If

    Set gActiveElement = document.ActiveElement   
End Function


'===============================================================================================================
Function DbDeleteOk()														
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	'------ Developer Coding part (Start)  -------------------------------------------------------------- 

	'------ Developer Coding part (End )   -------------------------------------------------------------- 

    Set gActiveElement = document.ActiveElement   
    
End Function


'===============================================================================================================
Function DbQuery() 

    On Error Resume Next
    Err.Clear           
    

	lsQueryMode = False    
    DbQuery = False                                                         

	
	If   LayerShowHide(1) = False Then
	     Exit Function 
	End If


	    
    Dim strVal
      
    If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001								
		strVal = strVal & "&txtBilltype=" & Trim(frm1.txtBilltype.value)
		strVal = strVal & "&rdoUsageFlg=" & Trim(frm1.txtRadio.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001									
		strVal = strVal & "&txtBilltype=" & Trim(frm1.txtBilltype.value)
		strVal = strVal & "&rdoUsageFlg=" & Trim(frm1.txtRadio.value)
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
    End If    
	
	Call RunMyBizASP(MyBizASP, strVal)												
	
    If Err.number = 0 Then	 
       DbQuery = True                                                             
    End If

    Set gActiveElement = document.ActiveElement   
	
	End Function
'===============================================================================================================
Function DbQueryOk()														
    
    On Error Resume Next
    Err.Clear           
    
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = Parent.OPMD_UMODE												
	lsQueryMode = True
    Dim LngRow

    ggoSpread.Source = frm1.vspdData

    Call SetToolbar("1110111100111111")					   
	Call SetDefaultVal()
	
	Set gActiveElement = document.ActiveElement   	
End Function

'===============================================================================================================
Function DbSave() 
 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
	
    On Error Resume Next
    Err.Clear           
    	
    DbSave = False                                                          
    
	
	If   LayerShowHide(1) = False Then
	     Exit Function 
	End If


	
	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
    
		lGrpCnt = 0    
		strVal = ""
		strDel = ""
    
		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col = 0
				Select Case .vspdData.Text
			        Case ggoSpread.InsertFlag							'☜: 신규 
						strVal = strVal & "C" & Parent.gColSep	& lRow & Parent.gColSep'☜: C=Create

						.vspdData.Col = S_BillType		
						strVal = strVal & FilterVar(UCase(Trim(.vspdData.Text)),"","SNM")   & Parent.gColSep
						
						.vspdData.Col = S_BillTypeNm		
						strVal = strVal & FilterVar(Trim(.vspdData.Text),"","SNM")   & Parent.gColSep      
				        
				        .vspdData.Col = S_ExceptFlg					        
				        If Trim(.vspdData.Text) = "1" Then
						strVal = strVal & "Y" & Parent.gColSep
				        Else
				        strVal = strVal & "N" & Parent.gColSep
				        End If

				        .vspdData.Col = S_ExportFlg	
				        If Trim(.vspdData.Text) = "1" Then	            
				        strVal = strVal & "Y" & Parent.gColSep
				        Else		            
				        strVal = strVal & "N" & Parent.gColSep
				        End If
				        
				        .vspdData.Col = S_GiFlg	
				        If Trim(.vspdData.Text) = "1" Then	            
				        strVal = strVal & "Y" & Parent.gColSep
				        Else		            
				        strVal = strVal & "N" & Parent.gColSep
				        End If

						.vspdData.Col = S_TransType			
				        strVal = strVal & FilterVar(Trim(.vspdData.Text),"","SNM")   & Parent.gColSep 

				        .vspdData.Col = S_UsageFlg
				        If Trim(.vspdData.Text) = "1" Then	            
				        strVal = strVal & "Y" & Parent.gColSep
				        Else		            
				        strVal = strVal & "N" & Parent.gColSep
				        End If		            
						
				        .vspdData.Col = S_AsFlg
				        If Trim(.vspdData.Text) = "1" Then	            
				        strVal = strVal & "Y" & Parent.gRowSep
				        Else		            
				        strVal = strVal & "N" & Parent.gRowSep
				        End If		            

				        lGrpCnt = lGrpCnt + 1 
					

					Case ggoSpread.UpdateFlag							'☜: 수정 
						strVal = strVal & "U" & Parent.gColSep	& lRow & Parent.gColSep'☜: U=Update
						
						.vspdData.Col = S_BillType		
						strVal = strVal & FilterVar(Trim(.vspdData.Text),"","SNM") & Parent.gColSep
						
						.vspdData.Col = S_BillTypeNm		
						strVal = strVal & FilterVar(Trim(.vspdData.Text),"","SNM") & Parent.gColSep      
				        
				        .vspdData.Col = S_ExceptFlg					        
				        If Trim(.vspdData.Text) = "1" Then
						strVal = strVal & "Y" & Parent.gColSep
				        Else
				        strVal = strVal & "N" & Parent.gColSep
				        End If

				        .vspdData.Col = S_ExportFlg	
				        If Trim(.vspdData.Text) = "1" Then	            
				        strVal = strVal & "Y" & Parent.gColSep
				        Else		            
				        strVal = strVal & "N" & Parent.gColSep
				        End if
				        
				        .vspdData.Col = S_GiFlg	
				        If Trim(.vspdData.Text) = "1" Then	            
				        strVal = strVal & "Y" & Parent.gColSep
				        Else		            
				        strVal = strVal & "N" & Parent.gColSep
				        End If

						.vspdData.Col = S_TransType			
				        strVal = strVal & FilterVar(Trim(.vspdData.Text),"","SNM") & Parent.gColSep 
				        
				        .vspdData.Col = S_UsageFlg
				        If Trim(.vspdData.Text) = "1" Then	            
				        strVal = strVal & "Y" & Parent.gColSep
				        Else		            
				        strVal = strVal & "N" & Parent.gColSep
				        End If		            

				        .vspdData.Col = S_AsFlg
				        If Trim(.vspdData.Text) = "1" Then	            
				        strVal = strVal & "Y" & Parent.gRowSep
				        Else		            
				        strVal = strVal & "N" & Parent.gRowSep
				        End If		            
						
				        lGrpCnt = lGrpCnt + 1 
					
					Case ggoSpread.DeleteFlag							'☜: 삭제 
						strDel = strDel & "D" & Parent.gColSep	& lRow & Parent.gColSep'☜: D=Delete
						.vspdData.Col = S_BillType
						strDel = strDel & FilterVar(Trim(.vspdData.Text),"","SNM") & Parent.gRowSep

						lGrpCnt = lGrpCnt + 1
				End Select
				

		Next
	
		.txtMaxRows.value = lGrpCnt
		.txtSpread.value = strDel & strVal
		
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										'☜: 비지니스 ASP 를 가동		
	
	End With
	
    If Err.number = 0 Then	 
       DbSave = True                                                             
    End If

    Set gActiveElement = document.ActiveElement   
    
    
End Function

'===============================================================================================================
Function DbSaveOk()															
    On Error Resume Next
    Err.Clear           
    						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData
    Call InitVariables
    
    Call MainQuery()
    
    Set gActiveElement = document.ActiveElement   

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>매출채권형태</font></td>
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
									<TD CLASS=TD5 NOWRAP>매출채권형태</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtBillType" ALT="매출채권형태" TYPE="Text" MAXLENGTH="4" SIZE=10 tag="15XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSOType" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCondtionPopup">&nbsp;<INPUT NAME="txtBillTypeNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
									<TD CLASS=TD5 NOWRAP>사용여부</TD>
									<TD CLASS=TD6 NOWRAP>
										<input type=radio CLASS="RADIO" name="rdoUsageFlg" id="rdoUsageFlgAll" value=" " tag = "11XXX" checked>
											<label for="rdoUsageFlgAll">전체</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS="RADIO" name="rdoUsageFlg" id="rdoUsageFlgYes" value="Y" tag = "11XXX">
											<label for="rdoUsageFlgYes">사용</label>&nbsp;&nbsp;&nbsp;&nbsp;
										<input type=radio CLASS = "RADIO" name="rdoUsageFlg" id="rdoUsageFlgNo" value="N" tag = "11XXX">
											<label for="rdoUsageFlgNo">미사용</label></TD>
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
										<script language =javascript src='./js/s1913ma1_I407246023_vspdData.js'></script>
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
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=<%=BizSize%>  FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">
<INPUT TYPE=HIDDEN NAME="txtBatch" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHSOType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadio" TAG="14">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>	
</BODY>
</HTML>
