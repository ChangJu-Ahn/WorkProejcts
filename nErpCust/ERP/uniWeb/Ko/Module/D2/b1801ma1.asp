<%@ LANGUAGE="VBSCRIPT" %>
<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : b1801ma1
'*  4. Program Name         : 경비항목설정 
'*  5. Program Desc         : 경비항목설정 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/08/02
'*  8. Modified date(Last)  : 2003/10/07
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : Cho in kuk
'* 11. Comment              : 2002/11/15 : UI성능 적용 
'**********************************************************************************************
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT> 
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                                 
<!-- #Include file="../../inc/lgvariables.inc" -->	

Const BIZ_PGM_ID = "b1801mb1.asp"            
Const C_Cost =     "단순경비"
Const C_Material = "물대포함"

Dim C_DomesticFlag 
Dim C_ChargeItem   
Dim C_ChargeItemBtn
Dim C_ChargeItemNm 
Dim C_ChargeAcct   
Dim C_ChargeAcctBtn
Dim C_ChargeAcctNm 
Dim C_ChargeVatFlg

Dim strChk			'Check Value 전체선택 변수 
Dim IsOpenPop		'Popup

'========================================================================================================
Sub initSpreadPosVariables()  

	C_DomesticFlag		= 1  '국내/외 구분 
	C_ChargeItem		= 2  '경비항목코드 
	C_ChargeItemBtn		= 3  '경비항목 Button
	C_ChargeItemNm		= 4  '경비항목명 
	C_ChargeAcct		= 5  '회계계정 
	C_ChargeAcctBtn		= 6  '회계계정 Button
	C_ChargeAcctNm		= 7
	C_ChargeVatFlg		= 8	 'VAT여부 

End Sub

'========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE            
    lgBlnFlgChgValue = False                    
    lgIntGrpCount = 0                               
    lgStrPrevKey = ""
    lgLngCurRows = 0  
    strChk = "0"
End Sub


'=========================================================================================================
Sub SetDefaultVal()
	frm1.txtRadio.value = frm1.rdoConModuleFlag_S.value
	frm1.rdoConModuleFlag_S.checked = True
	frm1.rdoModuleFlag_S.checked = True 
	frm1.txtChargeItem.focus
End Sub

'==========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'==========================================================================================================
Sub InitSpreadSheet()
 
	Call initSpreadPosVariables()    

	With frm1.vspdData

	 .MaxCols = 0
	 .MaxCols = C_ChargeVatFlg + 1 '☜: 최대 Columns의 항상 1개 증가시킴 
	 .Col = .MaxCols               '☜: 공통콘트롤 사용 Hidden Column
	 .ColHidden = True
	 
	 .MaxRows = 0
	 ggoSpread.Source = frm1.vspdData
	 
	 .ReDraw = false

	 ggoSpread.Spreadinit "V20051201",,parent.gAllowDragDropSpread    

	 Call GetSpreadColumnPos("A")
	 
	 ggoSpread.SSSetCombo	C_DomesticFlag, "단순경비여부", 15,0,True
	 ggoSpread.SSSetEdit 	C_ChargeItem, "경비항목", 20,,,20,2
	 ggoSpread.SSSetButton 	C_ChargeItemBtn
	 ggoSpread.SSSetEdit 	C_ChargeItemNm, "경비항목명", 30,,,50
	 ggoSpread.SSSetEdit 	C_ChargeAcct, "회계거래유형", 20,,,,2  
	 ggoSpread.SSSetButton 	C_ChargeAcctBtn
	 ggoSpread.SSSetEdit 	C_ChargeAcctNm, "회계거래유형명", 30,,,50
	 ggoSpread.SSSetCombo	C_ChargeVatFlg, "VAT여부", 10,0,False
	 
	 ggoSpread.SetCombo C_Cost & vbTab & C_Material ,C_DomesticFlag  
	 ggoSpread.SetCombo "Y" & vbTab & "N" ,C_ChargeVatFlg  
	 Call ggoSpread.MakePairsColumn(C_ChargeItem,C_ChargeItemBtn)
	 Call ggoSpread.MakePairsColumn(C_ChargeAcct,C_ChargeAcctBtn)

	 .ReDraw = true
	   
	End With
    
End Sub

'===========================================================================================================
Sub SetSpreadLock()
End Sub


'===========================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

	Dim lngIndex
	With frm1	    
		.vspdData.ReDraw = False
		 
		For lngIndex = pvStartRow to pvEndRow 
			.vspdData.Col = C_DomesticFlag : .vspdData.Row = lngIndex
			.vspdData.Text = C_Cost
		Next
		For lngIndex = pvStartRow to pvEndRow 
			.vspdData.Col = C_ChargeVatFlg : .vspdData.Row = lngIndex
			.vspdData.Text = "N"
		Next

		 ggoSpread.SSSetProtected C_DomesticFlag, pvStartRow, pvEndRow
		 If .rdoModuleFlag_M.checked = True Then
			ggoSpread.SpreadUnLock C_DomesticFlag, pvStartRow,C_DomesticFlag,pvEndRow
			ggoSpread.SSSetRequired C_DomesticFlag, pvStartRow, pvEndRow
		 End If
		 ggoSpread.SSSetRequired C_ChargeItem, pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected C_ChargeItemNm, pvStartRow, pvEndRow
		 ggoSpread.SSSetRequired C_ChargeAcct, pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected C_ChargeAcctNm, pvStartRow, pvEndRow  

		 ggoSpread.SSSetProtected C_ChargeVatFlg, pvStartRow, pvEndRow
		 If .rdoModuleFlag_M.checked = True Then
			ggoSpread.SpreadUnLock C_ChargeVatFlg, pvStartRow,C_ChargeVatFlg,pvEndRow
			ggoSpread.SSSetRequired C_ChargeVatFlg, pvStartRow, pvEndRow
		 End If
		 .vspdData.ReDraw = True	    
	End With

End Sub


'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)            
			C_DomesticFlag		= iCurColumnPos(1) 
			C_ChargeItem		= iCurColumnPos(2)  
			C_ChargeItemBtn		= iCurColumnPos(3)
			C_ChargeItemNm		= iCurColumnPos(4) 
			C_ChargeAcct		= iCurColumnPos(5)  
			C_ChargeAcctBtn		= iCurColumnPos(6)
			C_ChargeAcctNm		= iCurColumnPos(7)
			C_ChargeVatFlg		= iCurColumnPos(8)
    End Select    
End Sub


'==========================================================================================================
Function OpenCondtionPopup()

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "경비항목"      
	arrParam(1) = "A_JNL_ITEM"        
	arrParam(2) = frm1.txtChargeItem.value     
	arrParam(3) = ""         
	arrParam(4) = "JNL_TYPE = " & FilterVar("EC", "''", "S") & " "      
	arrParam(5) = "경비항목"      
	 
	arrField(0) = "JNL_CD"        
	arrField(1) = "JNL_NM"        
	    
	arrHeader(0) = "경비항목"      
	arrHeader(1) = "경비항목명"      

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCondtionPopup(arrRet)
	End If 
 
End Function


'===========================================================================
Function OpenSpreadPop(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True	  

	Select Case iWhere
	Case 0 '경비항목 
		arrParam(0) = "경비항목"   
		arrParam(1) = "A_JNL_ITEM"       
		arrParam(2) = strCode        
		arrParam(3) = ""         
		arrParam(4) = "JNL_TYPE = " & FilterVar("EC", "''", "S") & " "      
		arrParam(5) = "경비항목"      
 
		arrField(0) = "JNL_CD"        
		arrField(1) = "JNL_NM"        
		  
		arrHeader(0) = "경비항목"      
		arrHeader(1) = "경비항목명"      
	 
	Case 1 '회계계정(회계거래유형)

		If frm1.rdoModuleFlag_S.checked = True then
			arrParam(4) = "mo_cd=" & FilterVar("S", "''", "S") & " " 
		Else
			arrParam(4) = "mo_cd=" & FilterVar("M", "''", "S") & " " 
		End If  
	
		arrParam(0) = "회계거래유형"
		arrParam(1) = "A_ACCT_TRANS_TYPE"     
		arrParam(2) = strCode        
		arrParam(3) = ""         
		arrParam(5) = "회계거래유형"     
 
		arrField(0) = "TRANS_TYPE"       
		arrField(1) = "TRANS_NM"      
		  
		arrHeader(0) = "회계거래유형"      
		arrHeader(1) = "회계거래유형명"     
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	 "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
	 Exit Function
	Else
	 Call SetSpreadPop(arrRet, iWhere)
	End If 
 
End Function


'=================================================================================================================
Function SetCondtionPopup(Byval arrRet)
	frm1.txtChargeItem.value = arrRet(0)
	frm1.txtChargeItemNm.value = arrRet(1)    
	frm1.txtChargeItem.focus
End Function

'=================================================================================================================
Function SetSpreadPop(Byval arrRet,ByVal iWhere)

	With frm1

	 Select Case iWhere
	 Case 0 '경비항목 
	  .vspdData.Col = C_ChargeItem
	  .vspdData.Text = arrRet(0)
	  .vspdData.Col = C_ChargeItemNm
	  .vspdData.Text = arrRet(1)
	 Case 1 '회계계정 
	  .vspdData.Col = C_ChargeAcct
	  .vspdData.Text = arrRet(0)
	  .vspdData.Col = C_ChargeAcctNm
	  .vspdData.Text = arrRet(1)
	 End Select
	  
	 Call vspdData_Change(.vspdData.Col, .vspdData.Row)  

	End With

	lgBlnFlgChgValue = True
 
End Function


'=================================================================================================================
Sub SetQuerySpreadColor(ByVal lRow)
	Dim IRow
	With frm1
	   
	 .vspdData.ReDraw = False

	 ggoSpread.SSSetProtected C_DomesticFlag, lRow, .vspdData.MaxRows
	 ggoSpread.SSSetProtected C_ChargeItem, lRow, .vspdData.MaxRows
	 ggoSpread.SSSetProtected C_ChargeItemNm, lRow, .vspdData.MaxRows
	 ggoSpread.SSSetProtected C_ChargeItemBtn, lRow, .vspdData.MaxRows
	 ggoSpread.SSSetRequired C_ChargeAcct, lRow, .vspdData.MaxRows
	 ggoSpread.SSSetProtected C_ChargeAcctNm, lRow, .vspdData.MaxRows
	 If .rdoModuleFlag_M.checked = False Then
		 ggoSpread.SSSetProtected C_ChargeVatFlg, lRow, .vspdData.MaxRows
	 Else
	 	ggoSpread.SSSetRequired C_ChargeVatFlg, lRow, .vspdData.MaxRows
	 End If
	  
	 .vspdData.ReDraw = True
	   
	End With

End Sub

'=================================================================================================================
Sub Form_Load()

	Call LoadInfTB19029       
	Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")                                   	    
	    
	Call InitSpreadSheet
	
	Call SetDefaultVal 
	Call InitVariables       
	Call SetToolbar("11101101001011")      

End Sub

'=============================================================================================================
Sub rdoConModuleFlag_S_OnClick()
	frm1.txtRadio.value = frm1.rdoConModuleFlag_S.value 
End Sub

Sub rdoConModuleFlag_M_OnClick()
	frm1.txtRadio.value = frm1.rdoConModuleFlag_M.value
End Sub

Sub rdoModuleFlag_S_OnClick()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SSSetProtected C_DomesticFlag, 1, frm1.vspdData.MaxRows
	ggoSpread.SSSetProtected C_ChargeVatFlg, 1, frm1.vspdData.MaxRows
End Sub

Sub rdoModuleFlag_M_OnClick()
	ggoSpread.Source = frm1.vspdData
	ggoSpread.SpreadUnLock C_DomesticFlag, 1,C_DomesticFlag,frm1.vspdData.MaxRows
	ggoSpread.SSSetRequired C_DomesticFlag, 1, frm1.vspdData.MaxRows
	ggoSpread.SpreadUnLock C_ChargeVatFlg, 1,C_ChargeVatFlg,frm1.vspdData.MaxRows
	ggoSpread.SSSetRequired C_ChargeVatFlg, 1, frm1.vspdData.MaxRows
End Sub

'=============================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
	 
		ggoSpread.Source = frm1.vspdData
		   
		If Row > 0 And Col = C_ChargeItemBtn Then
			.Col = C_ChargeItem
			.Row = Row
			Call OpenSpreadPop(.Text, 0)
		ElseIf Row > 0 And Col = C_ChargeAcctBtn Then
			.Col = C_ChargeAcct
			.Row = Row
			Call OpenSpreadPop(.Text, 1)
		End If
	    
	    Call SetActiveCell(frm1.vspdData,Col - 1,Row,"M","X","X")
	    
	End With

End Sub

'=============================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	Dim strsstemp1, strsstemp2
	
	ggoSpread.Source = frm1.vspdData

    With frm1.vspdData
		.Row = Row
		Select Case Col
			Case C_DomesticFlag, C_ChargeVatFlg
				.Col = C_DomesticFlag
				strsstemp1 = .text
				.Col = C_ChargeVatFlg
				strsstemp2 = Trim(UCase(.text))
				If strsstemp2 = "Y" Or strsstemp2 = "N" Then
					If Trim(strsstemp1) = "물대포함" And strsstemp2 = "Y" Then
						.text = "N"
						Call SheetFocus(Row, C_ChargeVatFlg)
				        Call DisplayMsgBox("176135","x","x","x")
				        Exit Sub
					End If
				Else
					.text = "N"
					Call SheetFocus(Row, C_ChargeVatFlg)
			        Call DisplayMsgBox("17A003","x","VAT여부","x")
			        Exit Sub
				End If
		End Select
	End With	 

	ggoSpread.UpdateRow Row

	lgBlnFlgChgValue = True

End Sub


'=============================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row )
    Call SetPopupMenuItemInf("1111111111")	
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData

    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
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

'=============================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)
End Sub

'=============================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
		Exit Sub
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
    	Exit Sub
    End If


End Sub

'=============================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData

End Sub


'=============================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'=============================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub


'=============================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )    
    If OldLeft <> NewLeft Then Exit Sub
    
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If CheckRunningBizProcess = True Then Exit Sub		    
    	If lgStrPrevKey <> "" Then								
           Call DisableToolBar(parent.TBC_QUERY)
           Call DBQuery
    	End If
    End If    
End Sub

'=============================================================================================================
Function FncQuery() 
    Dim IntRetCD 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
        
    FncQuery = False       

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")  
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
         						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData
    Call InitVariables               

	Call SetToolbar("11101101001011")
    Call ggoOper.LockField(Document, "N")  


    Call DbQuery                
    
    If Err.number = 0 Then	
       FncQuery = True                                                            
    End If

    Set gActiveElement = document.ActiveElement       

    FncQuery = True                
        
End Function

'==============================================================================================================
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

    Call SetToolbar("11101101001011")          
    Call SetDefaultVal

    If Err.number = 0 Then	
       FncNew = True                                                              
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'==========================================================================================================
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

'==============================================================================================================
Function FncSave() 
    Dim IntRetCD 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
        
    FncSave = False         
    
    ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")        
        Exit Function
    End If       

    If Not chkField(Document, "2") Then     
       Exit Function
    End If    

    If ggoSpread.SSDefaultCheck = False Then     
       Exit Function
    End If

    If DbSave = False Then                                                        '☜: Query db data

       Exit Function
    End If
    

    If Err.number = 0 Then	
       FncSave = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
    
End Function

'==========================================================================================================
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
			Call SetSpreadColor (.ActiveRow, .ActiveRow)
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	    
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function

'==========================================================================================================
Function FncCancel() 
 	Dim iDx

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCancel = False                                                             '☜: Processing is NG    
    ggoSpread.Source = Frm1.vspdData	
    ggoSpread.EditUndo  
        
    If Err.number = 0 Then	
       FncCancel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function

'==========================================================================================================
Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status
        
    FncInsertRow = False                                                         '☜: Processing is NG
	
	If IsNumeric(Trim(pvRowCnt)) then
		imRow = Cint(pvRowCnt)
	Else
		imRow = AskSpdSheetAddRowCount()
		If imRow = "" Then
		    Exit Function
		End If
    End If
	
	
    frm1.vspdData.ReDraw = False
    frm1.vspdData.focus
    ggoSpread.Source = frm1.vspdData
    ggoSpread.InsertRow ,imRow    
    
    Call SetSpreadColor (frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow + imRow - 1)
    frm1.vspdData.ReDraw = True    
    
    lgBlnFlgChgValue = True
    
    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
	
    Set gActiveElement = document.ActiveElement    

End Function

'==========================================================================================================
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
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then	
       FncDeleteRow = True                                                            
    End If

    Set gActiveElement = document.ActiveElement   
    
End Function

'==========================================================================================================
Function FncPrint() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncPrint = False                                                              '☜: Processing is NG
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    
    '--------- Developer Coding Part (End) ------------------------------------------------------------
	Call Parent.FncPrint()                                                        

    If Err.number = 0 Then	 
       FncPrint = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'==========================================================================================================
Function FncExcel() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExcel = False                                                              '☜: Processing is NG

	Call Parent.FncExport(Parent.C_SINGLEMULTI)

    If Err.number = 0 Then	 
       FncExcel = True                                                            '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function

'==========================================================================================================
Function FncFind() 

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncFind = False                                                               '☜: Processing is NG

	Call Parent.FncFind(Parent.C_SINGLEMULTI, False)

    If Err.number = 0 Then	 
       FncFind = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   

End Function


'==========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'==========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'==========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call SetQuerySpreadColor(1)
End Sub


'==========================================================================================================
Function FncExit()
	Dim IntRetCD

    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncExit = False                                                               '☜: Processing is NG
    

    ggoSpread.Source = frm1.vspdData
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If

    If Err.number = 0 Then	 
       FncExit = True                                                             '☜: Processing is OK
    End If

    Set gActiveElement = document.ActiveElement   


End Function

'==========================================================================================================
Function DbDelete() 
    On Error Resume Next                                                    
End Function


'==========================================================================================================
Function DbDeleteOk()              
    On Error Resume Next                                                    
End Function


'==========================================================================================================
Function DbQuery() 

	Err.Clear                                                               
	    
	DbQuery = False                                                         

	If   LayerShowHide(1) = False Then
	  Exit Function 
	End If

	     
	Dim strVal
	    
	If lgIntFlgMode = Parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         
		strVal = strVal & "&txtChargeItem=" & Trim(frm1.txtHChargeItem.value)     
		If frm1.rdoModuleFlag_S.checked = True Then
			strVal = strVal & "&txtModuleType=" & "S" 
		Else
			strVal = strVal & "&txtModuleType=" & "M"
		End if  	
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	Else
		strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001         
		strVal = strVal & "&txtChargeItem=" & Trim(frm1.txtChargeItem.value)   
		strVal = strVal & "&txtModuleType=" & Trim(frm1.txtRadio.value) 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
	End If

	Call RunMyBizASP(MyBizASP, strVal)            
	 
	DbQuery = True                 

End Function

'==========================================================================================================
Function DbQueryOk()              
 
    lgIntFlgMode = Parent.OPMD_UMODE            
  
    Call ggoOper.LockField(Document, "Q")         
    Call SetToolbar("11101111001111")        
	Call SetQuerySpreadColor(1)
	Call ggoOper.SetReqAttr(frm1.rdoModuleFlag_S, "Q")
	Call ggoOper.SetReqAttr(frm1.rdoModuleFlag_M, "Q")
 
	ggoSpread.Source = frm1.vspdData

End Function


'========================================================================================
Function DbSave() 

	Err.Clear                
 
	Dim lRow        
	Dim lGrpCnt     
	Dim strVal, strDel
 
	DbSave = False  
	
	If   LayerShowHide(1) = False Then
       Exit Function 
	End If
 
	With frm1
		.txtMode.value = Parent.UID_M0002
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
    
		lGrpCnt = 1    
		strVal = ""
		strDel = ""

		For lRow = 1 To .vspdData.MaxRows
    
			.vspdData.Row = lRow
			.vspdData.Col = 0

			Select Case .vspdData.Text
			Case ggoSpread.InsertFlag       '☜: 신규 
				strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep'☜: C=Create
			    
			Case ggoSpread.UpdateFlag       '☜: 수정 
				strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep'☜: U=Update
			    
			Case ggoSpread.DeleteFlag       '☜: 삭제 
				strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep'☜: D=Delete
     
				'--- 경비항목코드 
				.vspdData.Col = C_ChargeItem              
				strDel = strDel & FilterVar(Trim(.vspdData.Text), "", "SNM") & Parent.gColSep              
				'-- 모듈구분 
				If frm1.rdoModuleFlag_S.checked = True Then
					strDel = strDel & "S" & Parent.gColSep
				Else
					strDel = strDel & "M" & Parent.gColSep
				End If
				'--- 단순경비여부 
				.vspdData.Col = C_DomesticFlag              
				If Trim(.vspdData.TypeComboBoxCurSel) = 0 Then   'C_Cost
					strDel = strDel & "C" & Parent.gRowSep
				ElseIf Trim(.vspdData.TypeComboBoxCurSel) = 1 Then  'C_Material
					strDel = strDel & "M" & Parent.gRowSep
				End IF
				lGrpCnt = lGrpCnt + 1 
			End Select

			Select Case .vspdData.Text
			Case ggoSpread.InsertFlag,ggoSpread.UpdateFlag
			       
				 '--- 경비항목코드 
				.vspdData.Col = C_ChargeItem              
				strVal = strVal & FilterVar(Trim(.vspdData.Text), "", "SNM") & Parent.gColSep		           
				 '--- 모듈구분 
				If frm1.rdoModuleFlag_S.checked = True Then
					strVal = strVal & "S" & Parent.gColSep
					strVal = strVal & "C" & Parent.gColSep
				Else
					strVal = strVal & "M" & Parent.gColSep
					'--- 단순경비여부 
					.vspdData.Col = C_DomesticFlag              
					If Trim(.vspdData.TypeComboBoxCurSel) = 0 Then   'C_Cost
						strVal = strVal & "C" & Parent.gColSep
					ElseIf Trim(.vspdData.TypeComboBoxCurSel) = 1 Then  'C_Material
						strVal = strVal & "M" & Parent.gColSep
					Else 
						strVal = strVal & "E" & Parent.gColSep      '단순경비여부에 string을 입력했을때 error raise를 위해 
					End IF
				End If       
				'--- 회계계정 
				.vspdData.Col = C_ChargeAcct              
				strVal = strVal & FilterVar(Trim(.vspdData.Text), "", "SNM") & Parent.gColSep
				    
				'--- 부대비배부여부 
				If frm1.rdoModuleFlag_S.checked = True Then
					strVal = strVal & "N" & Parent.gColSep
				Else
					.vspdData.Col = C_DomesticFlag              
					If Trim(.vspdData.TypeComboBoxCurSel) = 0 Then   'C_Cost
						strVal = strVal & "N" & Parent.gColSep
					ElseIf Trim(.vspdData.TypeComboBoxCurSel) = 1 Then  'C_Material
						strVal = strVal & "Y" & Parent.gColSep
					Else 
			            strVal = strVal & "E" & Parent.gColSep      '단순경비여부에 string을 입력했을때 error raise를 위해 
					End IF
				End If

				'--- VAT여부 
				If frm1.rdoModuleFlag_S.checked = True Then
					strVal = strVal & "N" & Parent.gRowSep
				Else
					.vspdData.Col = C_DomesticFlag
					If Trim(.vspdData.TypeComboBoxCurSel) = 1 Then
						strVal = strVal & "N" & Parent.gRowSep
					Else
						.vspdData.Col = C_ChargeVatFlg
						If Trim(.vspdData.TypeComboBoxCurSel) = 0 Or Trim(.vspdData.TypeComboBoxCurSel) = 1 Then
							strVal = strVal & FilterVar(Trim(.vspdData.Text), "", "SNM") & Parent.gRowSep
						Else
							Call DisplayMsgBox("17A003", parent.VB_INFORMATION, lRow & "행" & " " & "VAT여부","x")
							LayerShowHide(0)
							Exit Function
						End If
					End If
				End If
				
				lGrpCnt = lGrpCnt + 1 

			End Select
        
		Next
 
		.txtMaxRows.value = lGrpCnt-1
		.txtSpread.value = strDel & strVal
		         
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)          '☜: 비지니스 ASP 를 가동 
 
	End With
 
	DbSave = True                                                           
    
End Function

'=============================================================================================================
Function DbSaveOk()               

	Call ggoOper.ClearField(Document, "2")	         						
	ggoSpread.Source = frm1.vspdData 
	ggoSpread.ClearSpreadData
	Call InitVariables
	    
	If frm1.rdoModuleFlag_S.checked = True Then
		frm1.rdoConModuleFlag_S.checked = True
		frm1.txtRadio.value = "S"
	Else
		frm1.rdoConModuleFlag_M.checked = True
		frm1.txtRadio.value = "M"
	End If
	    
	Call MainQuery()

End Function

'==============================================================================
' Function : SheetFocus
' Description : 에러발생시 Spread Sheet에 포커스줌 
'==============================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>경비항목설정</font></td>
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
         <TD CLASS=TD5 NOWRAP>모듈구분</TD>
         <TD CLASS=TD6 NOWRAP>
          <input type=radio CLASS="RADIO" name="rdoConModuleFlag" id="rdoConModuleFlag_S" value="S" tag = "11XXX">
           <label for="rdoConModuleFlag_S">영업</label>&nbsp;&nbsp;&nbsp;&nbsp;
          <input type=radio CLASS = "RADIO" name="rdoConModuleFlag" id="rdoConModuleFlag_M" value="M" tag = "11XXX">
           <label for="rdoConModuleFlag_M">구매</label>
         </TD>
         <TD CLASS=TD5 NOWRAP>경비항목</TD>
         <TD CLASS=TD6 NOWRAP><INPUT NAME="txtChargeItem" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnChargeItem" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenCondtionPopup">&nbsp;<INPUT NAME="txtChargeItemNm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="24"></TD>
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
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
        <TD CLASS=TD5 NOWRAP>모듈구분</TD>
        <TD CLASS=TD6 NOWRAP>
         <input type=radio CLASS="RADIO" name="rdoModuleFlag" id="rdoModuleFlag_S" value="S" tag = "11XXX">
          <label for="rdoModuleFlag_S">영업</label>&nbsp;&nbsp;&nbsp;&nbsp;
         <input type=radio CLASS = "RADIO" name="rdoModuleFlag" id="rdoModuleFlag_M" value="M" tag = "11XXX">
          <label for="rdoModuleFlag_M">구매</label>
        </TD>
        <TD CLASS=TDT NOWRAP></TD>
        <TD CLASS=TD6 NOWRAP></TD>
       </TR>
       <TR>
        <TD HEIGHT="100%" WIDTH="100%" COLSPAN=4>
         <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
        </TD>
       </TR>
      </TABLE>
     </TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD <%=HEIGHT_TYPE_01%>></TD>
 </TR>
 <TR >
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
  </TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtSpread" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadio" tag="14">

<INPUT TYPE=HIDDEN NAME="txtHChargeItem" tag="24">
<INPUT TYPE=HIDDEN NAME="txtHRadio" tag="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
 <IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV> 
</BODY>
</HTML>
