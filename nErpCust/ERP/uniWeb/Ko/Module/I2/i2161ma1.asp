<%@ LANGUAGE="VBSCRIPT" %>
<!--'**********************************************************************************************
'*  1. Module Name          : Inventory
'*  2. Function Name        : 실사선별등록(Manual)
'*  3. Program ID           : i2161ma1 
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/19
'*  8. Modified date(Last)  : 2000/04/19
'*  9. Modifier (First)     : Kim Nam Hoon
'* 10. Modifier (Last)      : kim Nam Hoon
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/19 : ..........
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>

<SCRIPT LANGUAGE = "VBScript">
Option Explicit										

Const BIZ_LOOKUP_PGM_ID = "i2161mb1.asp"							
const BIZ_PGM_ID		= "i2161mb2.asp"

<!-- #Include file="../../inc/lgvariables.inc" -->

Dim lgStrKey1
Dim lgStrKey2
Dim lgStrKey3
Dim lgStrKey4
Dim lgCheckall     
Dim strClosedDt   

Dim IsOpenPop          
Dim SaveCheck

Dim C_Check
Dim C_ItemCd 									
Dim C_ItemNm
Dim C_Spec
Dim C_TrackingNo
Dim C_LotNo
Dim C_LotSubNo

'==========================================  2.1.1 InitVariables()  ======================================
Sub InitVariables()

	lgIntFlgMode = Parent.OPMD_CMODE         	         	
	lgBlnFlgChgValue = False     	               	
	lgIntGrpCount = 0                           	
	
	lgStrKey1 = ""
    lgStrKey2 = ""
    lgStrKey3 = ""
    lgStrKey4 = ""
	lgLngCurRows = 0                            	
	lgCheckall = 0
	SaveCheck = False
    Call SetToolbar("11000000000011")								

End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
Sub SetDefaultVal()

	frm1.btnRun.Disabled = True

	If Trim(frm1.txtPlantCd.value) = "" Then
		frm1.txtPlantNm.value = ""
	End if
	
	If Trim(frm1.txtSLCd.value) = "" Then
		frm1.txtSLNm.value = ""
	End if
	
	If Parent.gPlant <> "" Then
		frm1.txtPlantCd.value = Parent.gPlant
		frm1.txtPlantNm.value = Parent.gPlantNm
		frm1.txtSLCd.focus 
	Else
		frm1.txtPlantCd.focus 
	End If
	
End Sub


'========================================================================================
' Function Name : LoadInfTB19029
'========================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call LoadInfTB19029A("I", "*", "NOCOOKIE", "MA") %>
End Sub

'=============================================== 2.2.3 InitSpreadSheet() ================
' Function Name : InitSpreadSheet
'========================================================================================
Sub InitSpreadSheet()

	Call InitSpreadPosVariables()
	
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData
 		ggoSpread.Spreadinit "V20021106", , Parent.gAllowDragDropSpread
		
		.ReDraw = false
		
		.MaxCols = C_LotSubNo + 1					
		.MaxRows = 0
		
		
 		Call GetSpreadColumnPos("A")
 		Call AppendNumberPlace("6", "3", "0")
		
		ggoSpread.SSSetCheck C_Check, "", 4,,,1
		ggoSpread.SSSetEdit C_ItemCd, "품목", 20
		ggoSpread.SSSetEdit C_ItemNm, "품목명", 38
		ggoSpread.SSSetEdit C_Spec, "규격", 20
		ggoSpread.SSSetEdit C_TrackingNo, "Tracking No", 20
		ggoSpread.SSSetEdit C_LotNo, "LOT NO", 12
		ggoSpread.SSSetFloat	C_LotSubNo, "Lot No.순번", 8, "6", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec
		
 		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		ggoSpread.SpreadLockWithOddEvenRowColor()
		ggoSpread.SpreadUnLock C_Check, -1, C_Check
   		ggoSpread.SSSetSplit2(3)  

		.ReDraw = true
 
    End With
End Sub

'==========================================  2.2.7 InitSpreadPosVariables()  =============================
Sub InitSpreadPosVariables()
	C_Check			= 1
	C_ItemCd		= 2										
	C_ItemNm		= 3
	C_Spec			= 4
	C_TrackingNo	= 5
	C_LotNo			= 6
	C_LotSubNo		= 7
End Sub

'==========================================  2.2.8 GetSpreadColumnPos()  ==================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
 	Dim iCurColumnPos
 	
 	Select Case UCase(pvSpdNo)
 	Case "A"
 		ggoSpread.Source = frm1.vspdData 
 		
 		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

		C_Check			= iCurColumnPos(1)
		C_ItemCd		= iCurColumnPos(2)		  	
		C_ItemNm		= iCurColumnPos(3)
		C_Spec			= iCurColumnPos(4)
		C_TrackingNo	= iCurColumnPos(5)
		C_LotNo			= iCurColumnPos(6)
		C_LotSubNo		= iCurColumnPos(7)
 	End Select
End Sub

'------------------------------------------  OpenPlant()  -------------------------------------------------
Function OpenPlant()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "공장팝업"	
	arrParam(1) = "B_PLANT"				
	arrParam(2) = Trim(frm1.txtPlantCd.Value)
	arrParam(3) = ""
	arrParam(4) = ""			
	arrParam(5) = "공장"			
	
	arrField(0) = "PLANT_CD"	
	arrField(1) = "PLANT_NM"	
	
	arrHeader(0) = "공장"		
	arrHeader(1) = "공장명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtPlantCd.focus
		Exit Function
	Else
		frm1.txtPlantCd.Value    = arrRet(0)		
		frm1.txtPlantNm.Value    = arrRet(1)	
		frm1.txtPlantCd.focus	
	End If	
End Function

'------------------------------------------  OpenSL()  -------------------------------------------------
Function OpenSL()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X") 
		frm1.txtPlantCd.focus
		Exit Function
	End If

    If Plant_SLCd_Check(0) = False Then 
		Exit Function
	End If

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "창고팝업"	
	arrParam(1) = "B_STORAGE_LOCATION"				
	arrParam(2) = Trim(frm1.txtSLCd.Value)
	arrParam(3) = ""
	arrParam(4) = "PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S")	
	arrParam(5) = "창고"			
	
	arrField(0) = "SL_CD"	
	arrField(1) = "SL_NM"	
	
	arrHeader(0) = "창고"		
	arrHeader(1) = "창고명"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtSLCd.focus
		Exit Function
	Else
		frm1.txtSLCd.Value    = arrRet(0)		
		frm1.txtSLNm.Value    = arrRet(1)	
		frm1.txtSLCd.focus	
	End If	
End Function

'-----------------------  OpenItem()  -------------------------------------------------
Function OpenItem()
 
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
 
	If Trim(frm1.txtPlantCd.Value) = "" then 
		Call DisplayMsgBox("169901","X", "X", "X")
		frm1.txtPlantCd.focus   
		Exit Function
	End If
 
	If Plant_SLCd_Check(0) = False Then Exit Function

	If IsOpenPop = True Then Exit Function
 
	IsOpenPop = True

	arrParam(0) = "품목"     
	arrParam(1) = "B_Item_By_Plant,B_Item"
	arrParam(2) = Trim(frm1.txtItemCd.Value)
	arrParam(3) = ""
	arrParam(4) = "B_Item_By_Plant.Item_Cd = B_Item.Item_Cd And "
	arrParam(4) = arrParam(4) & "B_Item_By_Plant.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S")
	arrParam(5) = "품목"      

	arrField(0) = "B_Item_By_Plant.Item_Cd"  
	arrField(1) = "B_Item.Item_NM"   
	    
	arrHeader(0) = "품목"      
	arrHeader(1) = "품목명"    
    
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=445px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtItemCd.focus
		Exit Function
	Else
		frm1.txtItemCd.Value = arrRet(0)
		frm1.txtItemNm.Value = arrRet(1)
		frm1.txtItemCd.focus
	End If 
	Set gActiveElement = document.activeElement
End Function

'------------------------------------------  OpenTrackingNo()  -------------------------------------------------
Function OpenTrackingNo()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	If IsOpenPop = True Or UCase(frm1.txtPlantCd.ClassName)= UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True
	
	arrParam(0) = "Tracking No"			
	arrParam(1) = "s_so_tracking"			
	arrParam(2) = frm1.txtTrackingNo.value 	
	
	arrParam(3) = ""						
	arrParam(4) = ""						
	arrParam(5) = "Tracking No"			
	
    arrField(0) = "Tracking_No"	
    arrField(1) = "Item_Cd"	
    
    arrHeader(0) = "Tracking No"		
    arrHeader(1) = "품목"		
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtTrackingNo.focus
		Exit Function
	Else
		frm1.txtTrackingNo.Value	=arrRet(0)
		frm1.txtTrackingNo.focus
	End If	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
Sub Form_Load()
    Call LoadInfTB19029                                              
    Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)	
    Call ggoOper.LockField(Document, "N")                                   
    
    Call InitSpreadSheet                                                    
    Call InitVariables                                                      
    Call SetdefaultVal
End Sub

'=======================================================================================================
'   Event Name : txtInspDt_DblClick(Button)
'=======================================================================================================
Sub txtInspDt_DblClick(Button)
    If Button = 1 Then
        frm1.txtInspDt.Action = 7
        Call SetFocusToDocument("M")        
        frm1.txtInspDt.Focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtInspDt_Change()
'=======================================================================================================
Sub txtInspDt_Change()
	If lgIntFlgMode = Parent.OPMD_CMODE Then	
		lgBlnFlgChgValue = False
	Else
		lgBlnFlgChgValue = True	
	End if
End Sub

'==========================================================================================
'   Event Name : vspdData_Change
'==========================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

	ggoSpread.Source = frm1.vspdData
	
	Frm1.vspdData.Row = Row
	Frm1.vspdData.Col = Col
	
	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
		If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
			Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
		End If
	End If
End Sub

'==========================================================================================
'   Event Name : vspdData_ButtonClicked
'==========================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	
	With frm1.vspdData 
	
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_Check Then
			.Col = Col
			.Row = Row									
			IF .Text = 1 Then
				.Col = 0
				.Text = ggoSpread.UpdateFlag
				lgBlnFlgChgValue = True
			Elseif .Text = 0 Then
				.Col = 0
				.Text = ""
				lgBlnFlgChgValue = False
			End if  
							
		End If	
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_LeaveCell
'==========================================================================================
Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
    
	With frm1.vspdData 
	
		If Row = NewRow Then Exit Sub
	
		ggoSpread.Source = frm1.vspdData
		
		If Row > 0 And Col = C_Check Then
			.Col = Col
			.Row = Row									
			IF .Text = 1 Then
				.Col = 0
				.Text = ggoSpread.UpdateFlag
			Elseif .Text = 0 Then
				.Col = 0
				.Text = ""
			End if  
		End If	
	End With
End Sub

'==========================================================================================
'   Event Name : vspdData_TopLeftChange
'==========================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

	If OldLeft <> NewLeft Then Exit Sub
	If CheckRunningBizProcess = True Then Exit Sub
	
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop) Then	
		If (lgStrKey1 <> "" and lgStrKey2 <> "" and lgStrKey3 <> "" and lgStrKey4 <> "") Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End if
		End If
	End if  
End Sub

'========================================================================================
' Function Name : vspdData_Click
'========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
 	
    If lgIntFlgMode = Parent.OPMD_CMODE Then
 		Call SetPopupMenuItemInf("0000111111") 
 	Else
 	 	Call SetPopupMenuItemInf("0001111111")
 	End If
 	gMouseClickStatus = "SPC"   
 	Set gActiveSpdSheet = frm1.vspdData
    
 	If frm1.vspdData.MaxRows = 0 Then Exit Sub
 	
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

'========================================================================================
' Function Name : vspdData_MouseDown
'========================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 
 
'========================================================================================
' Function Name : vspdData_ColWidthChange
'========================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub 
 
'========================================================================================
' Function Name : vspdData_ScriptDragDropBlock
'========================================================================================
Sub vspdData_ScriptDragDropBlock( Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

   	If NewCol = C_Check or Col = C_Check Then
		Cancel = True
		Exit Sub
	End If

    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    Call GetSpreadColumnPos("A")
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
    Call InitSpreadSheet
    Call ggoSpread.ReOrderingSpreadData
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
' Function Name : FncQuery
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim TempInspDt 
    FncQuery = False                                           
    
    Err.Clear                                                  
     '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then	Exit Function				

    '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"X", "X")	
		If IntRetCD = vbNo Then	Exit Function
    End If

     '-----------------------
    'Erase contents area
    '-----------------------
    TempInspDt = frm1.txtInspDt.text
    Call ggoOper.ClearField(Document, "2")	
    
    If Plant_SLCd_Check(1) = False Then 
	    Call InitVariables
		Exit Function
	End If

    If SaveCheck = False Then
		frm1.txtCondPhyInvNo.Value = ""
	  	frm1.txtInspDt.text = UniConvDateAToB("<%=GetSvrDate%>",Parent.gServerDateFormat,Parent.gDateFormat)
    Else
		frm1.txtInspDt.text = TempInspDt
    End If

    Call InitVariables

    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then Exit Function
       
    FncQuery = True								
    
End Function

'========================================================================================
' Function Name : FncNew
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                            
    
    Err.Clear                                                          

    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData	
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    	IntRetCD = DisplayMsgBox("900015",Parent.VB_YES_NO,"X", "X")    	
		If IntRetCD = vbNo Then	Exit Function
	End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")
    Call ggoOper.LockField(Document, "N")                     
    Call InitVariables                                        
    Call SetDefaultVal    
    
    FncNew = True                                                       

End Function

'========================================================================================
' Function Name : FncSave
'========================================================================================
Function FncSave() 
    Dim IntRetCD 
	Dim strInspDt
	Dim strClosedDt1
	    
    FncSave = False                                                       
    
    Err.Clear                                                         
    On Error Resume Next                                             
    
    '-----------------------
    'Check content area
    '-----------------------
    If Not chkField(Document, "2")  Then Exit Function        
    
    '-----------------------
    'Precheck area
    '-----------------------
    ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","X", "X", "X")                       
        Exit Function
    End If
  
  	If Trim(frm1.txtPlantCd.value) <> Trim(frm1.txthPlantCd.value) OR _
  		Trim(frm1.txtSLCd.value) <> Trim(frm1.txthSLCd.value) Then
		Call DisplayMsgBox("900002","X","X","X")
		frm1.txtPlantCd.focus
		Exit Function
    End If

    If frm1.vspdData.MaxRows < 1 then
       Call DisplayMsgBox("900002","X", "X", "X")
       frm1.txtPlantCd.focus  
	   exit function
	End if 

	
	Set gActiveElement = document.activeElement

	If DBSave() = False Then Exit Function
    
    FncSave = True                                                 
    
End Function

'========================================================================================
' Function Name : FncCancel
'========================================================================================
Function FncCancel() 
    If frm1.vspdData.maxrows < 1 then exit function
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                             
End Function

'========================================================================================
' Function Name : FncPrint
'========================================================================================
Function FncPrint() 
    Call parent.FncPrint()
End Function

'========================================================================================
' Function Name : FncExcel
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)								
End Function

'========================================================================================
' Function Name : FncFind
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False)                                
End Function

'========================================================================================
' Function Name : FncExit
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X", "X")		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'========================================================================================
' Function Name : RemovedivTextArea
'========================================================================================
Function RemovedivTextArea()
	Dim i
	For i = 1 To divTextArea.children.length
		divTextArea.removeChild(divTextArea.children(0))
	Next
End Function

'========================================================================================
' Function Name : DbQuery
'========================================================================================
Function DbQuery() 

    Call LayerShowHide(1)       
    
    DbQuery = False
    
    Err.Clear                                                         
    Dim strVal
    Dim strFlag
    Dim strValid
    
	If frm1.RadioOutputType.rdoCase1.Checked Then
		strFlag = "Y"
	Else
		strFlag = "N"
	End if
	
	If frm1.RadioOutputType2.rdoCase3.Checked Then
		strValid = "Y"
	Else
		strValid = "N"
	End If
    
    
    With frm1    
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_LOOKUP_PGM_ID &	"?txtMode="				& Parent.UID_M0001				& _					
											"&txtPlantCd="			& Trim(.txthPlantCd.value)		& _			
											"&txtSLCd="				& Trim(.txthSLCd.value)			& _
											"&txtTrackingNo="		& Trim(.txtTrackingNo.value)	& _
											"&txtFlag="				& strFlag						& _
											"&txtValid="			& strValid						& _
											"&txtItemCd="			& Trim(.txtItemCd.value)		& _
											"&txthItemCd="			& Trim(lgstrKey1)				& _
											"&txtTrackingNo2="		& Trim(lgStrKey2)				& _
											"&txtLotNo="			& Trim(lgStrKey3)				& _
											"&txtLotSubNo="			& Trim(lgStrKey4)				& _
											"&txtMaxRows="			& .vspdData.MaxRows
		else                                          
			strVal = BIZ_LOOKUP_PGM_ID &	"?txtMode="				& Parent.UID_M0001				& _	
											"&txtPlantCd="			& Trim(.txtPlantCd.value)		& _			
											"&txtSLCd="				& Trim(.txtSLCd.value)			& _
											"&txtTrackingNo="		& Trim(.txtTrackingNo.value)	& _
											"&txtFlag="				& strFlag						& _
											"&txtValid="			& strValid						& _
											"&txtItemCd="			& Trim(.txtItemCd.value)		& _
											"&txthItemCd="			& Trim(lgstrKey1)				& _
											"&txtTrackingNo2="		& Trim(lgStrKey2)				& _
											"&txtLotNo="			& Trim(lgStrKey3)				& _
											"&txtLotSubNo="			& Trim(lgStrKey4)				& _
											"&txtMaxRows="			& .vspdData.MaxRows
		end if

		Call RunMyBizASP(MyBizASP, strVal)				
    End With
    DbQuery = True
    
End Function

'========================================================================================
' Function Name : DbQueryOk
'========================================================================================
Function DbQueryOk()						
    lgIntFlgMode = Parent.OPMD_UMODE					
    Call SetToolbar("11101001000111")
    frm1.btnRun.Disabled = True
    SaveCheck = False
End Function

'========================================================================================
' Function Name : DbSave
'========================================================================================
Function DbSave() 

    Dim lRow        
    Dim strVal
    Dim ColSep, RowSep     

	Dim strCUTotalvalLen
	Dim objTEXTAREA
	Dim iTmpCUBuffer
	Dim iTmpCUBufferCount
	Dim iTmpCUBufferMaxCount
	
    Call LayerShowHide(1)
        
    Err.Clear		
	
    DbSave = False                                                   
 
	frm1.txtMode.value = Parent.UID_M0002

	iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT
	ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
	iTmpCUBufferCount = -1
	strCUTotalvalLen = 0

 	With frm1.vspdData
   
		'-----------------------
		'Data manipulate area
		'-----------------------
		ColSep = Parent.gColSep
		RowSep = Parent.gRowSep
	
	    For lRow = 1 To .MaxRows
			.Row = lRow
			.Col = 0
        
			Select Case .Text

				Case ggoSpread.UpdateFlag

					.Col = C_Check
					If .Text = 1 Then		
	
						strVal = "U" & ColSep & lRow & ColSep 
					    .Col = C_ItemCd	
					    strVal = strVal & Trim(.Text) & ColSep
					    .Col = C_TrackingNo	
					    strVal = strVal & Trim(.Text) & ColSep
					    .Col = C_LotNo		
					    strVal = strVal & Trim(.Text) & ColSep
					    .Col = C_LotSubNo	
					    strVal = strVal & Trim(.Text) & RowSep


						If strCUTotalvalLen + Len(strVal) >  parent.C_FORM_LIMIT_BYTE Then
						                            
							Set objTEXTAREA = document.createElement("TEXTAREA")
							objTEXTAREA.name = "txtCUSpread"
							objTEXTAREA.value = Join(iTmpCUBuffer,"")
							divTextArea.appendChild(objTEXTAREA)     
							 
							iTmpCUBufferMaxCount = parent.C_CHUNK_ARRAY_COUNT   
							ReDim iTmpCUBuffer(iTmpCUBufferMaxCount)
							iTmpCUBufferCount = -1
							strCUTotalvalLen  = 0
									
						End If
						       
						iTmpCUBufferCount = iTmpCUBufferCount + 1
						      
						If iTmpCUBufferCount > iTmpCUBufferMaxCount Then    
							iTmpCUBufferMaxCount = iTmpCUBufferMaxCount + parent.C_CHUNK_ARRAY_COUNT
							ReDim Preserve iTmpCUBuffer(iTmpCUBufferMaxCount)
						End If   
									
						iTmpCUBuffer(iTmpCUBufferCount) =  strVal         
						strCUTotalvalLen = strCUTotalvalLen + Len(strVal)

					End if
			End Select

		Next

	End With

	If iTmpCUBufferCount > -1 Then 
	   Set objTEXTAREA = document.createElement("TEXTAREA")
	   objTEXTAREA.name   = "txtCUSpread"
	   objTEXTAREA.value = Join(iTmpCUBuffer,"")
	   divTextArea.appendChild(objTEXTAREA)     
	Else
		Call DisplayMsgBox("169912","X", "X", "X")   
		Call LayerShowHide(0)
		Exit function
	End If  
	 	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)						
	
    DbSave = True                                                   
    
End Function

'========================================================================================
' Function Name : DbSaveOk
'========================================================================================
Function DbSaveOk()						
    Dim PhyInvNo
    
    PhyInvNo = frm1.txtCondPhyInvNo.Value
    
    Call DisplayMsgBox("169916","X" ,PhyInvNo, "X") 	
	
	Call InitVariables
	SaveCheck = True
	ggoSpread.source = frm1.vspddata
    frm1.vspdData.MaxRows = 0
    Call FncQuery()

End Function

'========================================================================================
' Function Name : Checkall()
'========================================================================================
Function Checkall()
	
 Dim IRowCount 
 Dim IClnCount
 ggoSpread.Source = frm1.vspdData
 With frm1.vspdData    
  IF lgCheckall = 0 Then 
   for IClnCount = 0 to C_Check
   	for IRowCount = 1 to .MaxRows
   	     if IClnCount <> 0 then   	     	 
        	 .Row = IRowCount 
        	 .Col = IClnCount	 
 	         .text = 1     
 	     Else
 	         .Row = IRowCount
 	         .Col = IClnCount
 	         .Text =ggoSpread.UpdateFlag
 	     End if
	next    
   next
   lgCheckall = 1
   Else
   
   for IClnCount = 0 to C_Check
   	for IRowCount = 1 to .MaxRows
   	     if IClnCount <> 0 then   	     	 
        	 .Row = IRowCount 
        	 .Col = IClnCount	 
 	         .text = 0     
 	     Else
 	     End if
	next    
   next
   lgCheckall = 0
  End If

 End With
 
End Function

'========================================================================================
' Function Name : Plant_SLCd_Check
'========================================================================================
Function Plant_SLCd_Check(ByVal ChkIndex)

	Select Case ChkIndex
	
		Case 0
			'-----------------------
			'Check Plant CODE		
			'-----------------------
			If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
				Call DisplayMsgBox("125000","X","X","X")
				frm1.txtPlantNm.Value = ""
				frm1.txtPlantCd.focus
				Plant_SLCd_Check = False
				Exit function
			End If
			lgF0 = Split(lgF0, Chr(11))
			frm1.txtPlantNm.Value = lgF0(0)
			Plant_SLCd_Check = True	

		Case 1
		'-----------------------
		'Check SLCd CODE	
		'-----------------------
			If 	CommonQueryRs(" A.SL_NM, B.PLANT_NM, CONVERT(CHAR(10), B.INV_CLS_DT, 21) "," B_STORAGE_LOCATION A, B_PLANT B ", " A.PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S") & " AND A.PLANT_CD = B.PLANT_CD AND A.SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
				lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
					
				If 	CommonQueryRs(" PLANT_NM "," B_PLANT ", " PLANT_CD = " & FilterVar(frm1.txtPlantCd.Value, "''", "S"), _
					lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
				
					Call DisplayMsgBox("125000","X","X","X")
					frm1.txtPlantNm.Value = ""
					frm1.txtPlantCd.focus 
					Plant_SLCd_Check = False
					Exit function

				Else
					If 	CommonQueryRs(" SL_NM "," B_STORAGE_LOCATION ", " SL_CD = " & FilterVar(frm1.txtSLCd.Value, "''", "S"), _
						lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = False Then
							
						Call DisplayMsgBox("125700","X","X","X")
						frm1.txtSLNm.Value = ""
						frm1.txtSLCd.focus
						Plant_SLCd_Check = False 
						Exit function

					End If
				End If
									
				Call DisplayMsgBox("169922","X","X","X")
				frm1.txtSLCd.focus
				Plant_SLCd_Check = False
				Exit function
			End If
			
			lgF0 = Split(lgF0, Chr(11))
			lgF1 = Split(lgF1, Chr(11))
			lgF2 = Split(lgF2, Chr(11))
			frm1.txtSLNm.Value		= lgF0(0)
			frm1.txtPlantNm.Value	= lgF1(0)
			strClosedDt = UniConvDateAToB(lgF2(0),Parent.gServerDateFormat,Parent.gDateFormat)

			Plant_SLCd_Check = True	

	End Select
	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>실사선별(Manual)</font></td>
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
								<TD CLASS="TD5" NOWRAP>공장</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPlantCd" SIZE=8 MAXLENGTH=4 tag="12XXXU" ALT="공장"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPlant()">&nbsp;<INPUT TYPE=TEXT NAME="txtPlantNm" SIZE=25 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP>창고</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtSLCd" SIZE=8 MAXLENGTH=7 tag="12XXXU" ALT="창고"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSL" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSL()">&nbsp;<INPUT TYPE=TEXT NAME="txtSLNm" SIZE=25 tag="14"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtItemCd" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="품목"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnItemCd" align=top  TYPE="BUTTON" ONCLICK="vbscript:OpenItem()">&nbsp;<input TYPE=TEXT NAME="txtItemNm" SIZE="25" tag="14" ></TD>
								<TD CLASS="TD5" NOWRAP>Tracking No</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtTrackingNo" SIZE=15 MAXLENGTH=25 tag="11XXXU" ALT="Tracking No"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnTrackingNo" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenTrackingNo()"></TD>						                           
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>수량유무여부</TD>
								<TD CLASS="TD6">
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase1" TAG="1X"><LABEL FOR="rdoCase1">수량있음</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType" ID="rdoCase2" TAG="1X" checked><LABEL FOR="rdoCase2">전품목</LABEL>
								</TD>
								<TD CLASS="TD5" NOWRAP>품목유효일체크</TD>
								<TD CLASS="TD6">
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType2" ID="rdoCase3" TAG="1X"><LABEL FOR="rdoCase3">예</LABEL>
								<INPUT TYPE="RADIO" CLASS="RADIO" NAME="RadioOutputType2" ID="rdoCase4" TAG="1X" checked><LABEL FOR="rdoCase4">아니오</LABEL>
								</TD>						                           
							</TR>
						</TABLE>
					</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% VALIGN=TOP>
			<TABLE <%=LR_SPACE_TYPE_60%>>
				<TR>
					<TD CLASS="TD5" NOWRAP>실사번호</TD>
					<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtCondPhyInvNo" SIZE=20 MAXLENGTH=16 tag="11XXXU" ALT="실사번호"></TD>
					<TD CLASS="TD5" NOWRAP>실사일</TD>
					<TD CLASS="TD6" NOWRAP>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> name=txtInspDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="실사일" tag="22x1" id=fpDateTime1> <PARAM Name="AllowNull" Value="-1"><PARAM Name="Text" Value=""> </OBJECT>');</SCRIPT></TD>
				</TR>
				<TR>
						<TD WIDTH=100% HEIGHT=100% COLSPAN=4>
							<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%>NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
	
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnRun" CLASS="CLSMBTN" ONCLICK="vbscript:Checkall()" Flag=1>전체 선택/취소</BUTTON></TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txthPlantCd" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txthSLCd" tag="24" TABINDEX="-1"><INPUT TYPE=HIDDEN NAME="txthItemCd" tag="24" TABINDEX="-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

