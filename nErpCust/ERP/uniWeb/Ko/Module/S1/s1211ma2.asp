<%@ LANGUAGE="VBSCRIPT" %>
<%
'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1211MA2
'*  4. Program Name         : 고객품목등록 
'*  5. Program Desc         : 고객품목등록 
'*  6. Comproxy List        : PS1G105.dll, PS1G106.dll
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              : 2002/11/21 : Grid성능 적용, Kang Jun Gu
'*                            2002/12/10 : INCLUDE 다시 성능 적용, Kang Jun Gu
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit                 '☜: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim iDBSYSDate
Dim EndDate, StartDate

iDBSYSDate = "<%=GetSvrDate%>"
'------ ☆: 초기화면에 뿌려지는 마지막 날짜 ------
EndDate = UNIConvDateAtoB(iDBSYSDate, parent.gServerDateFormat, parent.gDateFormat)

'------ ☆: 초기화면에 뿌려지는 시작 날짜 ------
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

'**********************************************************************************************************%>
Const BIZ_PGM_ID = "s1211mb2.asp"            '☆: 비지니스 로직 ASP명 

Dim C_Cust_cd
Dim C_Cust_Popup
Dim C_Cust_nm
Dim C_Item_cd
Dim C_Item_Popup
Dim C_Item_Nm
Dim C_Spec
Dim C_CustItem_cd
Dim C_CustItem_nm
Dim C_Cust_unit
Dim C_Cust_unit_Popup
Dim C_Cust_Item_spec
Dim C_ChgFlg

Dim lgStrPrevKey1
Dim lgStrPrevKey2
Dim gblnWinEvent

'========================================================================================================
Sub initSpreadPosVariables()  
	C_Cust_cd			= 1
	C_Cust_Popup		= 2
	C_Cust_nm			= 3
	C_Item_cd			= 4
	C_Item_Popup		= 5
	C_Item_Nm			= 6
	C_Spec				= 7
	C_CustItem_cd		= 8
	C_CustItem_nm		= 9
	C_Cust_unit			= 10
	C_Cust_unit_Popup	= 11
	C_Cust_Item_spec	= 12
	C_ChgFlg			= 13
End Sub

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey1 = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                           'initializes Previous Key

    lgLngCurRows = 0                            'initializes Deleted Rows Count
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
 frm1.txtconBp_cd.focus  
 lgBlnFlgChgValue = False
End Sub

'========================================================================================================= 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()    
	ggoSpread.Source = frm1.vspdData

	With frm1.vspdData

       ggoSpread.Spreadinit "V20030401",,parent.gAllowDragDropSpread    

		.MaxRows = 0 : .MaxCols = 0
		.MaxCols = C_ChgFlg + 1            '☜: 최대 Columns의 항상 1개 증가시킴 


       Call GetSpreadColumnPos("A")
	   .ReDraw = false
		        
		ggoSpread.SSSetEdit  C_Cust_cd,           "고객", 10, 0,,10,2
		ggoSpread.SSSetButton C_Cust_Popup        
		ggoSpread.SSSetEdit  C_Cust_nm,           "고객명",20,0
		ggoSpread.SSSetEdit  C_Item_cd,           "품목", 15, 0,,18,2
		ggoSpread.SSSetButton C_Item_Popup
		ggoSpread.SSSetEdit  C_Item_nm,           "품목명",25, 0
		ggoSpread.SSSetEdit  C_Spec,              "품목규격",20, 0
		ggoSpread.SSSetEdit  C_CustItem_cd,       "고객품목코드", 15, 0,,40,2   '20071220::hanc::20-->40
		ggoSpread.SSSetEdit  C_CustItem_nm,       "고객품목명", 20,0,,50,1
		ggoSpread.SSSetEdit  C_Cust_unit,         "고객품목단위", 15, 0,,3,2
		ggoSpread.SSSetButton C_Cust_unit_Popup
		ggoSpread.SSSetEdit  C_Cust_Item_spec,    "고객품목규격", 20, 0,,50,1
		ggoSpread.SSSetEdit  C_ChgFlg, "Chgfg", 1, 2
		  
		SetSpreadLock "", 0, -1, ""
		.ReDraw = true   

		call ggoSpread.MakePairsColumn(C_Cust_cd,C_Cust_Popup)
		call ggoSpread.MakePairsColumn(C_Item_cd,C_Item_Popup)
		call ggoSpread.MakePairsColumn(C_Cust_unit,C_Cust_unit_Popup)

		Call ggoSpread.SSSetColHidden(C_ChgFlg,C_ChgFlg,True)		
		Call ggoSpread.SSSetColHidden(.MaxCols,.MaxCols,True)   '☜: 공통콘트롤 사용 Hidden Column

	End With
	    
End Sub

'===========================================================================================================
Sub SetSpreadLock(Byval stsFg, Byval Index, ByVal lRow , ByVal lRow2)
    With frm1
  .vspdData.ReDraw = False
   
  ggoSpread.spreadlock C_Cust_nm, lRow, -1
  ggoSpread.spreadUnlock C_Item_Cd, lRow, -1
  ggoSpread.spreadlock C_Item_nm, lRow, -1
  ggoSpread.spreadUnlock C_Cust_Cd, lRow, -1
      
  .vspdData.ReDraw = True

 End With
End Sub

'========================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetRequired  C_Cust_cd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Cust_nm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_Item_cd, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Item_nm, pvStartRow, pvEndRow
		ggoSpread.SSSetProtected C_Spec, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_CustItem_cd, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_CustItem_nm, pvStartRow, pvEndRow
		ggoSpread.SSSetRequired  C_Cust_unit, pvStartRow, pvEndRow
		.vspdData.ReDraw = True
	End With

End Sub

'=================SetSpreadColor1(조회후)================================================
Sub SetSpreadColor1(ByVal lRow)
	Dim Index

	With frm1
		.vspdData.ReDraw = False
		ggoSpread.SSSetProtected  C_Cust_cd, lRow, lRow
		ggoSpread.SSSetProtected  C_Cust_Popup, lRow, lRow
		ggoSpread.SSSetProtected  C_Cust_nm, lRow, lRow
		ggoSpread.SSSetProtected  C_Item_cd, lRow, lRow
		ggoSpread.SSSetProtected  C_Item_Popup, lRow, lRow
		ggoSpread.SSSetProtected  C_Item_nm, lRow, lRow
		ggoSpread.SSSetProtected  C_Spec,  lRow, lRow
		ggoSpread.SSSetRequired   C_CustItem_cd, lRow, lRow
		ggoSpread.SSSetRequired   C_CustItem_nm, lRow, lRow
		ggoSpread.SSSetRequired   C_Cust_unit, lRow, lRow
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
			C_Cust_cd			= iCurColumnPos(1)
			C_Cust_Popup		= iCurColumnPos(2)
			C_Cust_nm			= iCurColumnPos(3)    
			C_Item_Cd			= iCurColumnPos(4)
			C_Item_Popup		= iCurColumnPos(5)
			C_Item_Nm			= iCurColumnPos(6)
			C_Spec				= iCurColumnPos(7)
			C_CustItem_cd		= iCurColumnPos(8)
			C_CustItem_nm		= iCurColumnPos(9)
			C_Cust_unit			= iCurColumnPos(10)
			C_Cust_unit_Popup   = iCurColumnPos(11)
			C_Cust_Item_spec    = iCurColumnPos(12)
			C_ChgFlg			= iCurColumnPos(13)
    End Select    
End Sub

'===========================================================================================================
 Function OpenBizPartner(ByVal strCode)
	  Dim arrRet
	  Dim arrParam(5), arrField(6), arrHeader(6)
	        ggoSpread.Source = frm1.vspdData                                   
	 
	  frm1.vspdData.Col = 0
	  frm1.vspdData.Row = frm1.vspdData.ActiveRow
	  
	'  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
	  If gblnWinEvent = True Then Exit Function

	  gblnWinEvent = True

	  arrParam(0) = "고객"       
	  arrParam(1) = "B_BIZ_PARTNER"  
	  arrParam(2) = Trim(strCode)    
	  arrParam(3) = ""         
	  arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"    
	  arrParam(5) = "고객"       

	  arrField(0) = "BP_CD"        
	  arrField(1) = "BP_NM"        

	  arrHeader(0) = "고객"    
	  arrHeader(1) = "고객명"  
	 
	  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	  gblnWinEvent = False

	  If arrRet(0) = "" Then
	   Exit Function
	  Else
	   Call SetBizPartner(arrRet)
	  End If
 End Function

'===========================================================================================================
 Function OpenItem(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        ggoSpread.Source = frm1.vspdData                                   
 
  frm1.vspdData.Col = 0
  frm1.vspdData.Row = frm1.vspdData.ActiveRow
  
'  If frm1.vspdData.Text <> ggoSpread.InsertFlag Then Exit Function 
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "품목"       
  arrParam(1) = "B_ITEM"        
  arrParam(2) = Trim(strCode)   
  arrParam(3) = ""         
  arrParam(4) = ""         
  arrParam(5) = "품목" 

  arrField(0) = "ITEM_CD"  
  arrField(1) = "ITEM_NM"  
	arrField(2) = "Spec"	

  arrHeader(0) = "품목" 
  arrHeader(1) = "품목명"
	arrHeader(2) = "규격"

	
  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
  
  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetItem(arrRet)
  End If
 End Function

'===========================================================================================================
 Function OpenUnit(ByVal strCode)
  Dim arrRet
  Dim arrParam(5), arrField(6), arrHeader(6)
        
  If gblnWinEvent = True Then Exit Function

  gblnWinEvent = True

  arrParam(0) = "고객품목단위"      
  arrParam(1) = "B_unit_of_measure"     
  arrParam(2) = Trim(strCode)     
  arrParam(3) = ""                
  arrParam(4) = ""                
  arrParam(5) = "고객품목단위"

  arrField(0) = "UNIT"            
  arrField(1) = "UNIT_NM"         

  arrHeader(0) = "품목단위"   
  arrHeader(1) = "품목단위명" 

  arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
    "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

  gblnWinEvent = False

  If arrRet(0) = "" Then
   Exit Function
  Else
   Call SetUnit(arrRet)
  End If
 End Function

'===========================================================================================================
Function OpenConSItemDC(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If gblnWinEvent = True Then Exit Function

	gblnWinEvent = True

	Select Case iWhere
	Case 0
		arrParam(1) = "B_BIZ_PARTNER"      
		arrParam(2) = Trim(frm1.txtconBp_cd.value)
		arrParam(3) = ""                          
		arrParam(4) = "BP_TYPE in (" & FilterVar("C", "''", "S") & " ," & FilterVar("CS", "''", "S") & ")"     
		arrParam(5) = "고객"     
 
		arrField(0) = "BP_CD"        
		arrField(1) = "BP_NM"        
		  
		arrHeader(0) = "고객"    
		arrHeader(1) = "고객명"  
		frm1.txtconBp_cd.focus 
	Case 1
		arrParam(1) = "b_item"               
		arrParam(2) = Trim(frm1.txtconItem_cd.Value)  
		arrParam(3) = ""                              
		arrParam(4) = ""                              
		arrParam(5) = "품목"         
 
		arrField(0) = "item_cd"            
		arrField(1) = "item_nm"            
		arrField(2) = "spec"               
			
	   
		arrHeader(0) = "품목"          
		arrHeader(1) = "품목명"        
		arrHeader(2) = "규격"          
		frm1.txtConItem_cd.focus 
	Case 2
		arrParam(1) = "S_BP_ITEM"               
		arrParam(2) = Trim(frm1.txtConCustItem_cd.value)  
		arrParam(3) = ""                              
		arrParam(4) = ""                              
		arrParam(5) = "고객품목"         
 
		arrField(0) = "bp_item_cd"            
		arrField(1) = "bp_item_nm"            
		arrField(2) = "bp_item_spec"               
			
	   
		arrHeader(0) = "고객품목"          
		arrHeader(1) = "고객품목명"        
		arrHeader(2) = "고객품목규격"          
		frm1.txtConCustItem_cd.focus  
	End Select
	   
	arrParam(3) = "" 
	arrParam(0) = arrParam(5)        <%' 팝업 명칭 %>

	Select Case iWhere
	Case 1, 2
	 arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	   "dialogWidth=700px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	Case Else
	 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
	  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	End Select

	gblnWinEvent = False

	If arrRet(0) = "" Then
	 Exit Function
	Else
	 Call SetConSItemDC(arrRet, iWhere)
	End If 
 
End Function

'===========================================================================================================
Function SetBizPartner(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Cust_cd
  .vspdData.Text = arrRet(0)
  .vspdData.Col = C_Cust_Nm
  .vspdData.Text = arrRet(1)
 End With
End Function

'===========================================================================================================
Function SetUnit(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Cust_Unit
  .vspdData.Text = arrRet(0)
 Call vspdData_Change(C_Cust_Unit, .vspdData.ActiveRow)
 End With
End Function

'===========================================================================================================
Function SetItem(Byval arrRet)  
 With frm1
  .vspdData.Col = C_Item_cd
  .vspdData.Text = arrRet(0)
  .vspdData.Col = C_Item_Nm
  .vspdData.Text = arrRet(1)
  .vspdData.Col = C_Spec
  .vspdData.Text = arrRet(2)
 End With
End Function

'===========================================================================================================
Function SetConSItemDC(Byval arrRet, Byval iWhere)
	 With frm1
		Select Case iWhere
		Case 0
			.txtconBp_cd.value = arrRet(0) 
			.txtconBp_nm.value = arrRet(1)   
		Case 1
			.txtconItem_nm.value = arrRet(1)   
			.txtConItem_Spec.value = arrRet(2)
			.txtconItem_cd.value = arrRet(0) 
		Case 2
			.txtconCustItem_nm.value = arrRet(1)   
			.txtConCustItem_Spec.value = arrRet(2) 
			.txtconCustItem_cd.value = arrRet(0) 
		End Select
	 End With
End Function

'===========================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Call SetActiveCell(frm1.vspdData, iDx, iRow,"M","X","X")			
              Exit For
           End If
                      
       Next
          
    End If   
End Sub
'===========================================================================================================
Sub Form_Load()

	Call LoadInfTB19029()

	Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field

	'----------  Coding part  -------------------------------------------------------------
	Call InitSpreadSheet
	Call SetDefaultVal
	Call InitVariables                                                      '⊙: Initializes local global variables

    Call SetToolBar("1110110100101111")          '⊙: 버튼 툴바 제어 
	
End Sub

'===========================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )
   
End Sub

'===========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
 With frm1.vspdData 
 
  ggoSpread.Source = frm1.vspdData
  
  If Row > 0 Then
	Select Case Col
	Case C_Cust_Popup
      .Col = Col - 1
      .Row = Row
      Call OpenBizPartner(.text)
	Case C_Item_Popup
      .Col = Col - 1
      .Row = Row
      Call OpenItem(.Text)
	Case C_Cust_Unit_Popup
      .Col = Col - 1
      .Row = Row
      Call OpenUnit(.Text)
	End Select 
  End If
  
  Call SetActiveCell(frm1.vspdData,Col-1,Row,"M","X","X")
    
 End With

End Sub

'===========================================================================================================
Sub vspdData_Click(ByVal Col , ByVal Row)
	gMouseClickStatus = "SPC"   

	Set gActiveSpdSheet = frm1.vspdData
	' Context 메뉴의 입력, 삭제, 데이터 입력, 취소 
	Call SetPopupMenuItemInf(Mid(gToolBarBit, 6, 2) + "0" + Mid(gToolBarBit, 8, 1) & "111111")
	    
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
	'---frm1.vspdData.Col = C_MajorCd
		
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'===========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'===========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    If Row <= 0 Then
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    End If
	
End Sub

'===========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)

    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If

End Sub

'===========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
End Sub

'===========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

   Frm1.vspdData.Row = Row
   Frm1.vspdData.Col = Col

   If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
   End If
  
   ggoSpread.UpdateRow Row
End Sub

'===========================================================================================================
 Sub vspdData_LeaveCell(ByVal Col, ByVal Row, ByVal NewCol, ByVal NewRow, Cancel)
  With frm1.vspdData
   If Row >= NewRow Then
    Exit Sub
   End If

   If NewRow = .MaxRows Then
    If lgStrPrevKey1 <> "" And lgStrPrevKey2 <> "" Then       <% '⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
     Call DbQuery()
    End If
   End If
  End With
 End Sub

'===========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then Exit Sub
    
    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop)  Then 
		If lgStrPrevKey1 <> "" And lgStrPrevKey2 <> "" Then
			If CheckRunningBizProcess Then Exit Sub
			Call DisableToolBar(Parent.TBC_QUERY)
			Call DbQuery()
		End If
	End if        
   
End Sub

'===========================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing

    '-----------------------
    'Check previous data area
    '-----------------------
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO,"x","x")
			If IntRetCD = vbNo Then
				Exit Function
			End If
    End If
     
    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")          '⊙: Clear Contents  Field
    Call InitVariables
                   '⊙: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then         '⊙: This function check indispensable field
       Exit Function
    End If
    
    '-----------------------
    'Query function call area
    '-----------------------
    Call DbQuery
       
    FncQuery = True   
                 '⊙: Processing is OK
	Call ggoOper.LockField(Document, "3")
     
End Function

'===========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    'On Error Resume Next                                                    '☜: Protect system from crashing
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Check previous data area
    '-----------------------
	ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If       
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                         '⊙: Clear Condition Field
    Call ggoOper.LockField(Document, "N")                                          '⊙: Lock  Suitable  Field
    Call InitVariables                                                      '⊙: Initializes local global variables
    Call SetDefaultVal
    Call SetToolBar("1110110100101111")          '⊙: 버튼 툴바 제어 
    FncNew = True                                                           '⊙: Processing is OK

End Function

'===========================================================================================================
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         '⊙: Processing is NG
    
    Err.Clear                                                               '☜: Protect system from crashing
    On Error Resume Next                                                    '☜: Protect system from crashing
    
    ggoSpread.Source = frm1.vspdData
    
    '-----------------------
    'Precheck area
    '-----------------------
    If ggoSpread.SSCheckChange = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        Exit Function
    End If
    
    '-----------------------
    'Check content area
    '-----------------------    
 If Not chkField(Document, "2") Then  <% '⊙: Check contents area %>
  Exit Function
 End If

 If Not ggoSpread.SSDefaultCheck Then  <% '⊙: Check contents area %>
  Exit Function
 End If
 
    '-----------------------
    'Save function call area
    '-----------------------
    Call DbSave                                                      '☜: Save db data
    
    FncSave = True                                                          '⊙: Processing is OK
    
End Function

'===========================================================================================================
Function FncCopy() 
 
 frm1.vspdData.ReDraw = False
 
 if frm1.vspdData.maxrows < 1 then exit function
 
    ggoSpread.Source = frm1.vspdData 
    ggoSpread.CopyRow
    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow
    
 frm1.vspdData.ReDraw = True
 
End Function

'===========================================================================================================
Function FncCancel() 

    if frm1.vspdData.maxrows < 1 then exit function

 ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

'===========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim IntRetCD
    Dim imRow
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

	With frm1
	 
		FncInsertRow = False                                                         '☜: Processing is NG

		If Not chkField(Document, "2") Then
		Exit Function
		End If

		If IsNumeric(Trim(pvRowCnt)) Then
		    imRow = CInt(pvRowCnt)
		Else
		    imRow = AskSpdSheetAddRowCount()
		    If imRow = "" Then
		        Exit Function
		    End If
		End If
		.vspdData.ReDraw = False
		.vspdData.focus
		ggoSpread.Source = .vspdData
		ggoSpread.InsertRow ,imRow
		SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.ReDraw = True
	End With
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
    If Err.number = 0 Then
		FncInsertRow = True                                                          '☜: Processing is OK
    End If           
    Set gActiveElement = document.ActiveElement   
    
End Function

'===========================================================================================================
Function FncDeleteRow()
 Dim lDelRows
 Dim iDelRowCnt, i
 
 if frm1.vspdData.maxrows < 1 then exit function
 
 With frm1.vspdData 
  If .MaxRows = 0 Then
   Exit Function
  End If

  .focus
  ggoSpread.Source = frm1.vspdData

  lDelRows = ggoSpread.DeleteRow

  lgBlnFlgChgValue = True
 End With
End Function

'===========================================================================================================
 Function FncPrint()
     ggoSpread.Source = frm1.vspdData
  Call parent.FncPrint()             <%'☜: Protect system from crashing%>
 End Function

'===========================================================================================================
 Function FncExcel() 
  Call parent.FncExport(Parent.C_MULTI)
 End Function

'===========================================================================================================
 Function FncFind() 
  Call parent.FncFind(Parent.C_MULTI, False)
 End Function

'===========================================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'===========================================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'===========================================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call SetSpreadColor1(-1)
End Sub

'===========================================================================================================
 Function FncExit()
  Dim IntRetCD

  FncExit = False

  ggoSpread.Source = frm1.vspdData
  If ggoSpread.SSCheckChange = True Then
   IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")   <%'⊙: "Will you destory previous data"%>
   If IntRetCD = vbNo Then
    Exit Function
   End If
  End If

  FncExit = True
 End Function

'===========================================================================================================
Function DbQuery() 
    Dim LngLastRow      
    Dim LngMaxRow       
    Dim StrNextKey      

    DbQuery = False
    
    Err.Clear                                                               '☜: Protect system from crashing

	If   LayerShowHide(1) = False Then Exit Function 

	Dim strVal
	   
    With frm1
		If lgIntFlgMode = Parent.OPMD_UMODE Then
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001   
			strVal = strVal & "&txtconBp_cd=" & Trim(.txtHconBp_cd.value)
			strVal = strVal & "&txtconItem_cd=" & Trim(.txtHconItem_cd.value)
			strVal = strVal & "&txtconItem_Nm=" & Trim(.txtHconItem_nm.value)
			strVal = strVal & "&txtconItem_spec=" & Trim(.txtHconItem_spec.value)
			strVal = strVal & "&txtConCustItem_cd=" & Trim(.txtHconCustItem_cd.value)
			strVal = strVal & "&txtConCustItem_nm=" & Trim(.txtHconcustItem_nm.value)
			strVal = strVal & "&txtConCustItem_spec=" & Trim(.txtHConCustItem_Spec.value)

			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		Else
			strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001   
			strVal = strVal & "&txtconBp_cd=" & Trim(.txtconBp_cd.value)
			strVal = strVal & "&txtconItem_cd=" & Trim(.txtconItem_cd.value)
			strVal = strVal & "&txtconItem_nm=" & Trim(.txtconItem_nm.value)
			strVal = strVal & "&txtconItem_spec=" & Trim(.txtConItem_Spec.value)
			strVal = strVal & "&txtConCustItem_cd=" & Trim(.txtconCustItem_cd.value)
			strVal = strVal & "&txtConCustItem_nm=" & Trim(.txtconCustItem_nm.value)
			strVal = strVal & "&txtConCustItem_spec=" & Trim(.txtConCustItem_Spec.value)

			strVal = strVal & "&lgStrPrevKey1=" & lgStrPrevKey1
			strVal = strVal & "&lgStrPrevKey2=" & lgStrPrevKey2
		End if

		Call RunMyBizASP(MyBizASP, strVal)          '☜: 비지니스 ASP 를 가동 
   End With

   DbQuery = True

End Function

'===========================================================================================================
 Function DbSave() 
  Dim lRow
  Dim lGrpCnt
  Dim strVal, strDel
  Dim intInsrtCnt
  Dim TotDocAmt, dblQty, dblPrice, dblOldQty

  DbSave = False              <% '⊙: Processing is OK %>

  Call LayerShowHide(1)

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
     Case ggoSpread.InsertFlag        <% '☜: 신규 %>
      strVal = strVal & "C" & Parent.gColSep & lRow & Parent.gColSep <% '☜: C=Create, Row위치 정보 %>

      .vspdData.Col = C_Cust_cd        <% '2 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_Item_cd        <% '3 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_CustItem_cd       <% '4 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_CustItem_Nm       <% '4 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Cust_unit        <% '5 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Cust_Item_spec      <% '6 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep

      lGrpCnt = lGrpCnt + 1
  
     Case ggoSpread.UpdateFlag        <% '☜: Update %>
      strVal = strVal & "U" & Parent.gColSep & lRow & Parent.gColSep <% '☜: U=Update, Row위치 정보 %>
      
      .vspdData.Col = C_Cust_cd        <% '2 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_Item_cd        <% '3 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_CustItem_cd       <% '4 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_CustItem_Nm       <% '4 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_Cust_unit        <% '5 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Cust_Item_spec      <% '6 %>
      strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep
      
      lGrpCnt = lGrpCnt + 1
 
     Case ggoSpread.DeleteFlag        <% '☜: 삭제 %>
      strDel = strDel & "D" & Parent.gColSep & lRow & Parent.gColSep <% '☜: D=Update, Row위치 정보 %>

      .vspdData.Col = C_Cust_cd        <% '2 %>
      strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_Item_cd        <% '3 %>
      strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_CustItem_cd       <% '4 %>
      strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_CustItem_Nm       <% '4 %>
      strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep

      .vspdData.Col = C_Cust_unit        <% '5 %>
      strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
      
      .vspdData.Col = C_Cust_Item_spec      <% '6 %>
      strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep
      
      lGrpCnt = lGrpCnt + 1

    End Select
   Next

   .txtMaxRows.value = lGrpCnt-1
   .txtSpread.value = strDel & strVal
   
   Call ExecMyBizASP(frm1, BIZ_PGM_ID)      <% '☜: 비지니스 ASP 를 가동 %>

  End With

  DbSave = True              <% '⊙: Processing is NG %>
 End Function
 
'===========================================================================================================
Function DbQueryOk()             <% '☆: 조회 성공후 실행로직 %>
	<% '------ Reset variables area ------ %>
	lgIntFlgMode = Parent.OPMD_UMODE           <% '⊙: Indicates that current mode is Update mode %>
	lgBlnFlgChgValue = False
  
	Call ggoOper.LockField(Document, "Q")        <% '⊙: This function lock the suitable field %>
	Call SetToolBar("1110111100111111")         <% '⊙: 버튼 툴바 제어 %>
  
	frm1.txtconBp_cd.focus
End Function
 
'===========================================================================================================
 Function DbSaveOk()              <%'☆: 저장 성공후 실행 로직 %>
  Call ggoOper.ClearField(Document, "2")
  Call InitVariables
  Call MainQuery()
 End Function
 

</SCRIPT>
<!-- #Include file="../../inc/UNI2kcm.inc" --> 
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
							<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>고객품목등록</font></td>
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
								<TD CLASS="TD5" NOWRAP>고객</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtconBp_cd" ALT="고객" TYPE="Text" MAXLENGTH=10 SiZE=10  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconBp_cd.value,0">&nbsp;<INPUT NAME="txtconBp_nm" TYPE="Text" MAXLENGTH="50" SIZE=25 tag="14"></TD>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP></TD>
							</TR>
							<TR>	
								<TD CLASS="TD5" NOWRAP>고객품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConCustItem_cd" ALT="고객품목" TYPE="Text" MAXLENGTH=20 SiZE=20  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconcustItem_cd.value, 2"></TD>
								<TD CLASS="TD5" NOWRAP>품목</TD>
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConItem_cd" ALT="품목" TYPE="Text" MAXLENGTH=18 SiZE=20  tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConSItemDC frm1.txtconItem_cd.value,1"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>고객품목명
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConCustItem_nm" TYPE="Text" ALT="고객품목명" MAXLENGTH="50" SIZE=40 tag="11"></TD>
								<TD CLASS="TD5" NOWRAP>품목명
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConItem_nm" TYPE="Text" ALT="품목명" MAXLENGTH="40" SIZE=40 tag="11"></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>고객품목규격
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConCustItem_Spec" TYPE="Text" ALT="고객품목규격" MAXLENGTH="50" SIZE=40 tag="11"></TD>
								<TD CLASS="TD5" NOWRAP>규격
								<TD CLASS="TD6" NOWRAP><INPUT NAME="txtConItem_Spec" TYPE="Text" ALT="규격" MAXLENGTH="50" SIZE=40 tag="11"></TD>
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
								<script language =javascript src='./js/s1211ma2_I487075269_vspdData.js'></script>
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
	<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX="-1"></IFRAME>
</TD>
</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX="-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHconBp_cd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHconItem_cd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHconItem_nm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHconItem_spec" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHconCustItem_cd" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHconcustItem_nm" tag="24" TABINDEX="-1">
<INPUT TYPE=HIDDEN NAME="txtHConCustItem_Spec" tag="24" TABINDEX="-1">
</FORM>

<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML>

