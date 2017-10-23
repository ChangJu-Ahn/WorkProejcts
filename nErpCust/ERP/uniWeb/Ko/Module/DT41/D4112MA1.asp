
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 전자세금계산서(SmartBill)
'*  2. Function Name        : 
'*  3. Program ID           : D4112ma1.asp
'*  4. Program Name         : 인증서관리
'*  5. Program Desc         :
'*  6. Modified date(First) : 2011/05/13
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>


<SCRIPT LANGUAGE=VBSCRIPT>

Option Explicit

Const BIZ_PGM_ID = "D4112mb1.asp"												<%'비지니스 로직 ASP명 %>

Dim C_Check  '선택
Dim C_Reg_No  
Dim C_Reg_No_Pop
Dim C_Com_Nm
Dim C_User_Dn
Dim C_Expr_Date


Dim IsOpenPop          
Dim lgSortKey1
Dim lgSortKey2

<!-- #Include file="../../inc/lgvariables.inc" -->

Sub InitSpreadPosVariables()
     C_Check        = 1
     C_Reg_No       = 2
     C_Reg_No_Pop   = 3
     C_Com_Nm       = 4
     C_User_Dn      = 5
     C_Expr_Date    = 6        
    
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    lgStrPrevKey = ""                           'initializes Previous Key

    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
End Sub

Sub LoadInfTB19029()
    <!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
    <% Call loadInfTB19029A("I", "B","NOCOOKIE","BA") %>
End Sub

Sub InitSpreadSheet()
    Call initSpreadPosVariables()  

	With frm1.vspdData
	
	 ggoSpread.Source = frm1.vspdData
	 ggoSpread.Spreadinit "V20021125",,parent.gAllowDragDropSpread    
    .ReDraw = false
    .MaxCols = C_Expr_Date + 1												'☜: 최대 Columns의 항상 1개 증가시킴 
    .Col = .MaxCols														'공통콘트롤 사용 Hidden Column
    .ColHidden = True
    .MaxRows = 0			
					
    Call GetSpreadColumnPos("A")  

    ggoSpread.SSSetCheck    C_Check,            "선택", 4,,"", TRUE
	ggoSpread.SSSetEdit		C_Reg_No,		    "사업자등록번호", 20,,,10,1
	ggoSpread.SSSetButton	C_Reg_No_Pop	
	ggoSpread.SSSetEdit		C_Com_Nm,   	    "회사명", 20,,,70,1	
	ggoSpread.SSSetEdit		C_User_Dn,	        "인증서정보", 25,,,150,1			
    ggoSpread.SSSetDate		C_Expr_Date,        "만기일",15,2, gDateFormat   		    	

								
	call ggoSpread.MakePairsColumn(C_Reg_No,C_Reg_No_Pop)

    'Call ggoSpread.SSSetColHidden(C_smartbill_pw,	C_smartbill_pw,True) 

	.ReDraw = true
	
    Call SetSpreadLock 
    
    End With
    
End Sub

Sub SetSpreadLock()
    With frm1

    .vspdData.ReDraw = False

    ggoSpread.SpreadLock	    C_Reg_No	    , -1, 	C_Reg_No		
    ggoSpread.SpreadLock	    C_Reg_No_Pop	, -1, 	C_Reg_No_Pop		
    ggoSpread.SpreadLock	    C_Com_Nm	    , -1, 	C_Com_Nm		
    ggoSpread.SpreadLock	    C_User_Dn	    , -1, 	C_User_Dn		
    ggoSpread.SpreadLock	    C_Expr_Date	    , -1, 	C_Expr_Date		

    ggoSpread.SSSetRequired		C_Reg_No,	-1, -1
            
    'ggoSpread.SSSetRequired		C_user_name,		-1, -1
        
    .vspdData.ReDraw = True

    End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired	C_Reg_No,	pvStartRow, pvEndRow
    
    'ggoSpread.SSSetProtected	C_user_id_pop,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_Com_Nm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_User_Dn,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_Expr_Date,		pvStartRow, pvEndRow
    
    .vspdData.ReDraw = True
    End With
End Sub


Sub SetSpreadColor1(ByVal pvStartRow, ByVal pvEndRow)
    With frm1
    
    .vspdData.ReDraw = False
    ggoSpread.SSSetRequired	C_Reg_No,	pvStartRow, pvEndRow
    
    ggoSpread.SSSetProtected	C_Reg_No,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_Com_Nm,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_User_Dn,		pvStartRow, pvEndRow
    ggoSpread.SSSetProtected	C_Expr_Date,		pvStartRow, pvEndRow
    
    .vspdData.ReDraw = True
    End With
End Sub

Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)            
            C_Check        = iCurColumnPos(1)
            C_Reg_No       = iCurColumnPos(2)
            C_Reg_No_Pop   = iCurColumnPos(3)
            C_Com_Nm       = iCurColumnPos(4)
            C_User_Dn      = iCurColumnPos(5)
            C_Expr_Date    = iCurColumnPos(6)                                                                 

    End Select    
End Sub


Function OpenPopup(Byval strcode, Byval iWhere)
   Dim arrRet
   Dim arrParam(5), arrField(6), arrHeader(6)
	Dim iCalledAspName
	Dim IntRetCD
	Dim strUserid

   If IsOpenPop = True Then Exit Function

   IsOpenPop = True


	Select Case iWhere
	    Case 1 
        
         'frm1.vspddata.Col = C_user_id
         'strUserid = Trim(frm1.vspddata.value) 
         
         'if strUserid = "" then
         '   Call DisplayMsgBox("17A002","X","사용자ID(를)","X")	
         '   '17A002: %1을 입력하세요.               		            
         '   IsOpenPop = False
         '   Exit Function
         'end if
                                                  
         arrParam(0) = "회사"
         arrParam(1) = "B_TAX_BIZ_AREA a (NOLOCK) INNER JOIN B_BIZ_PARTNER b (NOLOCK) on (b.BP_CD = a.TAX_BIZ_AREA_CD) " ' TABLE 명칭 
         arrParam(2) = strcode      ' Code Condition
         arrParam(3) = ""       ' Name Cindition
         arrParam(4) = "" ' Where Condition
         arrParam(5) = "회사"    ' 조건필드의 라벨 명칭 
         

         arrField(0) = "replace(b.BP_RGST_NO,'-','')"     ' Field명(0)
         arrField(1) = "b.BP_NM"     ' Field명(1)

         arrHeader(0) = "사업자등록번호"    ' Header명(0)
         arrHeader(1) = "회사명"     ' Header명(1)
	    		            
		Case Else
		
	     Exit Function
   End Select
	 
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
            "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

   IsOpenPop = False

   If arrRet(0) = "" Then
      Exit Function
   Else
		Call SetPopup(arrRet, iWhere)
   End If 
 
End Function



Function Open_User(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol, TempCd
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "usr_id, usr_nm"				<%' 팝업 명칭 %>
	arrParam(1) = "z_usr_mast_rec"				<%' TABLE 명칭 %>
	arrParam(2) = strCode						<%' Code Condition%>
	arrParam(4) = ""							<%' Name Cindition%>
	arrParam(5) = "사용자"						<%' 조건필드의 라벨 명칭 %>

	arrField(0) = "usr_id"						<%' Field명(0)%>
	arrField(1) = "usr_nm"						<%' Field명(1)%>

	arrHeader(0) = "사용자"						<%' Header명(0)%>
	arrHeader(1) = "사용자명"					<%' Header명(1)%>

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
												Array(arrParam, arrField, arrHeader), _
												"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetUser(arrRet, iWhere)
	End If	

End Function

Function Open_User1()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim OriginCol, TempCd
	Dim IntRetCD

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "회사"	
	arrParam(1) = "B_TAX_BIZ_AREA a (nolock) inner join  B_BIZ_PARTNER b (nolock) on a.TAX_BIZ_AREA_CD = b.BP_CD "    
	arrParam(2) = Trim(frm1.txtRegNo.value)			' Code Condition
	arrParam(4) = ""							' Name Cindition
	arrParam(5) = "회사"					' 조건필드의 라벨 명칭 

	arrField(0) = "replace(b.BP_RGST_NO,'-','')"  'Field명(0)
	arrField(1) = "b.BP_NM"				' Field명(1)

	arrHeader(0) = "사업자등록번호"			' Header명(0)
	arrHeader(1) = "회사명"					' Header명(1)

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", _
												Array(arrParam, arrField, arrHeader), _
												"dialogWidth=520px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	 
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		frm1.txtRegNo.value = arrRet(0)
		frm1.txtRegNm.value = arrRet(1)
	End If	
	
End Function


'========================================================================================================
' Function Name : SetPopUp(Byval arrRet, Byval iWhere)
'========================================================================================================
Function SetPopup(Byval arrRet, Byval iWhere)
   With frm1
      Select Case iWhere
	     Case 1   ' 
           .vspdData.Col = C_Reg_NO
           .vspdData.Text = arrRet(0)
           .vspdData.Col = C_Com_Nm
           .vspdData.Text = arrRet(1)
           
'           Call vspdData_Change(C_PuNo, .vspdData.Row)	     	     
	     Case 2   

      End Select
   End With
End Function


Function SetUser(Byval arrRet, Byval iWhere)
	With frm1 
		.vspdData.Col = C_user_id
		.vspdData.Text = arrRet(0)

		.vspdData.Col = C_user_name
		.vspdData.Text = arrRet(1)

		lgBlnFlgChgValue = True
	End With
End Function


Sub Form_Load()
    Call LoadInfTB19029                                                     <%'Load table , B_numeric_format%>
    
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
                                                                           <%'Format Numeric Contents Field%>                                                                            
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
          
    Call InitSpreadSheet                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>   
 
    Call SetToolbar("1100110100001111")										<%'버튼 툴바 제어 %>

End Sub


Sub vspdData_Change(ByVal Col , ByVal Row )

   Dim strUserid
   Dim strUserRegNo
   Dim strPW1
   Dim strPW


    Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
                        
        Case  C_Reg_No
            strUserRegNo = Trim(Frm1.vspdData.value)
    
            If strUserRegNo = "" Then
  	            Frm1.vspdData.Col = C_Com_Nm
                Frm1.vspdData.value = ""
            Else					
				If CommonQueryRs(" replace(b.BP_RGST_NO,'-',''), b.BP_NM "," B_TAX_BIZ_AREA a (nolock), B_BIZ_PARTNER b (nolock) "," a.TAX_BIZ_AREA_CD = b.BP_CD and replace(b.BP_RGST_NO,'-','') = '" & strUserRegNo & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
					Frm1.vspdData.Col = C_Reg_No
				    Frm1.vspdData.text = Trim(Replace(lgF0,Chr(11),""))
				    Frm1.vspdData.Col = C_Com_Nm
				    Frm1.vspdData.text = Trim(Replace(lgF1,Chr(11),""))
				else
				    Call DisplayMsgBox("970000","X",strUserid,"X")	               		
				    '970000:%1 이(가) 존재하지 않습니다.
				    Frm1.vspdData.Col = C_Reg_No
				    Frm1.vspdData.text = ""
				    Frm1.vspdData.Col = C_Com_Nm
				    Frm1.vspdData.text = ""
				END IF					
            End if    
                                                                      
    End Select

    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row
End Sub

Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("1101011111") 

	gMouseClickStatus = "SPC"   

	Set gActiveSpdSheet = frm1.vspdData

	If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
		Exit Sub
	End If

	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData

		If lgSortKey = 1 Then
			ggoSpread.SSSort Col               'Sort in ascending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
			lgSortKey = 1
		End If
		Exit Sub
	End If

	frm1.vspdData.Row = Row
End Sub

Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData_DblClick(ByVal Col, ByVal Row)				
	Dim iColumnName

	If Row <= 0 Then
		Exit Sub
	End If

	If frm1.vspdData.MaxRows = 0 Then
		Exit Sub
	End If
End Sub

Sub vspdData_GotFocus()
	ggoSpread.Source = Frm1.vspdData
End Sub

Sub vspdData_MouseDown(Button , Shift , x , y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub    

Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub


Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
    'Call InitSpreadComboBox
    Call SetSpreadColor1(-1,-1)
	Call ggoSpread.ReOrderingSpreadData()
	'Call InitData()
End Sub

Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
		    Select Case Col
			    'Case C_user_id_pop
				'  .Row = Row
		        '   .Col = C_user_id
		        '    Call Open_User(.Text, 1)
			    Case C_Reg_No_Pop
			        frm1.vspddata.Col = C_Reg_No			    
				    Call OpenPopUp(.text, 1)        
		    End Select    
	    End If								
    End With                     
End Sub

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If CheckRunningBizProcess = True Then					'☜: 조회중이면 다음 조회 안하도록 체크 
        Exit Sub
	End If
	
	If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
'    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	    <%'☜: 재쿼리 체크 %>
'    	If (lgStrPrevKey <> "" And lgStrPrevKey2 <> "" And lgStrPrevKey3 <> "") Then <%'다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 %>
'      		Call DisableToolBar(parent.TBC_QUERY)					'☜ : Query 버튼을 disable 시킴.
'			If DBQuery = False Then 
'			   Call RestoreToolBar()
'			   Exit Sub 
'			End If 
'    	End If
'    End if

    If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
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
	End If  

End Sub

Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        
    
    Err.Clear                                                               <%'Protect system from crashing%>

  '-----------------------
    'Check previous data area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
  '-----------------------
    'Erase contents area
    '----------------------- 
    If frm1.txtRegNo.value = "" then
        frm1.txtRegNm.value = ""
    End If
    
    Call ggoOper.ClearField(Document, "2")									<%'Clear Contents  Field%>
    ggoSpread.ClearSpreadData
    Call InitVariables                                                      <%'Initializes local global variables%>
    															
  '-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
       Exit Function
    End If
    
  '-----------------------
    'Query function call area
    '----------------------- 
    If DbQuery = False Then Exit Function		  					<%'Query db data%>
       
    FncQuery = True
            
End Function

Function FncSave() 
        
    FncSave = False                                                         
    
    'Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>
    
  '-----------------------
    'Precheck area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = False Then
        Call DisplayMsgBox("900001", "X", "X", "X")                          <%'No data changed!!%>
        Exit Function
    End If
    
  '-----------------------
    'Check content area
    '----------------------- 
    ggoSpread.Source = frm1.vspdData
    If Not ggoSpread.SSDefaultCheck Then     'Not chkField(Document, "2") OR    '⊙: Check contents area
       Exit Function
    End If
    

    
  '-----------------------
    'Save function call area
    '----------------------- 
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 

    frm1.vspdData.ReDraw = False

    if frm1.vspdData.maxrows < 1 then exit function

    ggoSpread.Source = frm1.vspdData 
    ggoSpread.CopyRow

    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow

    'Key field clear
    frm1.vspdData.Col=C_user_id
    frm1.vspdData.Text=""

    frm1.vspdData.Col = C_user_name
    frm1.vspdData.Text=""

    frm1.vspdData.ReDraw = True

End Function

Function FncCancel() 
    ggoSpread.Source = frm1.vspdData	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim IntRetCD
    Dim imRow
    Dim iRow 
    Dim iIntCnt

    'On Error Resume Next                                                          '☜: If process fails
    'Err.Clear                                                                     '☜: Clear error status
   
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
	    .vspdData.ReDraw = False
	    .vspdData.focus
	    ggoSpread.Source = .vspdData
	    ggoSpread.InsertRow  .vspdData.ActiveRow, imRow	    	    
	    SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1

		'For iIntCnt = .vspdData.ActiveRow To .vspdData.ActiveRow + imRow - 1
		'	.vspdData.Row = iIntCnt 			
		'	.vspddata.col = C_smartbill_pw1
		'	.vspddata.text = "**************"
												
		'	.vspdData.ReDraw = True
			
	    'Next
	End With
     
End Function

Function FncDeleteRow() 
    Dim lDelRows

    With frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    
End Function

Function FncPrint()
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
End Function

Function FncExcel() 
    Call parent.FncExport(parent.C_MULTI)											 <%'☜: 화면 유형 %>
End Function

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI, False)                                      <%'☜:화면 유형, Tab 유무 %>
End Function

Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

Function DbQuery() 

    DbQuery = False
    
    'Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
    
    With frm1
    
    If lgIntFlgMode = parent.OPMD_UMODE Then
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001							'☜: 
		strVal = strVal & "&txtuserId=" & .hUserId.value 			'☆: 조회 조건 데이타  			 
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

    Else
		strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001						
		strVal = strVal & "&txtUserId=" & Trim(.txtRegNo.value)			'☆: 조회 조건 데이타
		strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey
		strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows

    End If        
	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 
       
    End With
    
    DbQuery = True
    
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE
   
	Call SetToolbar("1100111100111111")										<%'버튼 툴바 제어 %>

	call SetSpreadColor1(-1, -1)
End Function

Function DbSave() 
    Dim lRow        
    Dim lGrpCnt  
	Dim strVal, strDel
	
    DbSave = False                                                          
    
      If LayerShowHide(1) = False then
    	Exit Function 
    End if
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>

	With frm1
		.txtMode.value = parent.UID_M0002
    
  '-----------------------
    'Data manipulate area
    '----------------------- %>
    lGrpCnt = 1
    
    strVal = ""
    strDel = ""
    
  '-----------------------
    'Data manipulate area
    '----------------------- %>
    ' Data 연결 규칙 
    ' 0: Flag , 1: Row위치, 2~N: 각 데이타 

    For lRow = 1 To .vspdData.MaxRows
    
		    .vspdData.Row = lRow
		    .vspdData.Col = 0
		    
		    Select Case .vspdData.Text

		        Case ggoSpread.InsertFlag								'☜: 신규 
					
					strVal = strVal & "C" & parent.gColSep	& lRow & parent.gColSep	'☜: C=Create
										
		            .vspdData.Col = C_Reg_No	    	'2
		            strVal = strVal & Trim(.vspdData.Text) & parent.gColSep
		            
		            .vspdData.Col = C_Com_Nm		'3
		            strVal = strVal & Trim(.vspdData.Text) & parent.gRowSep
		            		            
		            
		            lGrpCnt = lGrpCnt + 1
		            
		        Case ggoSpread.UpdateFlag								'☜: 수정 
		
					'strVal = strVal & "U" & parent.gColSep & lRow & parent.gColSep	'☜: U=Update
		            		         
		            lGrpCnt = lGrpCnt + 1
		            
		            
		        Case ggoSpread.DeleteFlag								'☜: 삭제 

					strDel = strDel & "D" & parent.gColSep & lRow & parent.gColSep	'☜: U=Update

		            .vspdData.Col = C_Reg_No		'2
		            strDel = strDel & Trim(.vspdData.Text) & parent.gRowSep
		            
		            
		            lGrpCnt = lGrpCnt + 1
		    End Select
		            
		Next
	
	.txtMaxRows.value = lGrpCnt - 1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)										<%'☜: 비지니스 ASP 를 가동 %>
	
	End With
	
    DbSave = True                                                           
    
End Function

Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
	
    Call MainQuery()
End Function

Sub txtRegNo_OnChange()
    lgBlnFlgChgValue = True

     'If frm1.fpdtWk_yymm.Text = "" Then
	'	Call DisplayMsgBox("800489","x",frm1.fpdtWk_yymm.alt,"x")
	'	frm1.fpdtWk_yymm.focus
	'	Exit Sub
	'End If  

   if Trim(frm1.txtuserId.value) <> "" then
    If CommonQueryRs(" replace(b.BP_RGST_NO,'-',''), b.BP_NM "," B_TAX_BIZ_AREA a (nolock), B_BIZ_PARTNER b (nolock) "," a.TAX_BIZ_AREA_CD = b.BP_CD and replace(b.BP_RGST_NO,'-','') = '" & Trim(frm1.txtRegNo.value) & "'" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then
       frm1.txtRegNo.value = Trim(Replace(lgF0,Chr(11),""))
       frm1.txtRegNm.value =  Trim(Replace(lgF1,Chr(11),""))
    else
	   frm1.txtRegNo.value = ""
       frm1.txtRegNm.value =  ""
        Call DisplayMsgBox("970000","X",frm1.txtRegNo.alt,"X")	               		
        '970000:%1 이(가) 존재하지 않습니다.              
	   Exit Sub 	  
    End if 
   end if 
	
	frm1.txtuserId.focus
End Sub


Function fnRegister()
	DIm IntRetCD
	DIm strURL
    Dim StrRegNo
    Dim StrComName
    
	DIm lRow
	

    'If lgIntFlgMode <> parent.OPMD_UMODE Then												'Check if there is retrived data
    '    IntRetCD = DisplayMsgBox("900002","X","X","X")                                       
    '    Exit Function
    'End If
   With Frm1	
	For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = C_Check

            If .vspdData.text = "1" Then
                
                  frm1.vspdData.Col = C_Reg_No
                  StrRegNo = Trim(.vspdData.value)
                  
                  frm1.vspdData.Col = C_Com_Nm
                  StrComName = Trim(.vspdData.value)
                
                If CommonQueryRs(" MINOR_NM "," B_MINOR (nolock) "," MAJOR_CD = 'DT400' AND MINOR_CD = '01' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
	                    strURL =  Trim(Replace(lgF0,Chr(11),"")) & "XXSB_DTI_CERT.asp?CERT_REGNO=" + StrRegNo + "&CERT_COMNAME=" + StrComName + ""

                    window.open "", "legacy", ""   	
	                .target = "legacy"	
                    .action =  strURL
                    .submit()
                    'arrRet =  window.showModalDialog(strUrl ,, "dialogWidth=700px; dialogHeight=350px; center: Yes; help: No; resizable: No; status: no; scroll:Yes;")       

                else
                    Call DisplayMsgBox("970000","X", "스마트빌서버URL","X")	               		
                    '970000:%1 이(가) 존재하지 않습니다. 
                    strURL = ""             
                   Exit Function 	  
                end if               
            
	        End if
	Next
   End With	
						 				
	
End Function


Function fnDown()
	DIm IntRetCD
	DIm strURL

	DIm lRow
	

    'If lgIntFlgMode <> parent.OPMD_UMODE Then												'Check if there is retrived data
    '    IntRetCD = DisplayMsgBox("900002","X","X","X")                                       
    '    Exit Function
    'End If
	
	If CommonQueryRs(" MINOR_NM "," B_MINOR (nolock) "," MAJOR_CD = 'DT400' AND MINOR_CD = '02' " ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = True Then   
	    strURL =  Trim(Replace(lgF0,Chr(11),""))
	    
	    if Right(Trim(strURL),4) <> ".exe"  then	    	    	    
	        strURL =  Trim(Replace(lgF0,Chr(11),"")) & "/AgentFile/SBCAgentForClientDemo.exe"
	    else
	        strURL =  Trim(Replace(lgF0,Chr(11),""))
	    end if    
	    'strURL =  Trim(Replace(lgF0,Chr(11),"")) & "/AgentFile/SBCAgent_Demo(v1.2.1.4).exe"
    else
        Call DisplayMsgBox("W70001","X", "Agent 다운로드 URL이 유효하지 않습니다.","X")	               		
        '970000:%1 이(가) 존재하지 않습니다.              
	   Exit Function 	  
    end if   
			 				
		'.result.value= strXML
		'.loginid.value=lgSSOID
		'.formversion.value=strVersion
		'.D1.value=lgSSOD1		'싱글아이디 정보
		'.businessid.value= "budget"

		
	'Call BtnDisabled(True)
	 open strURL
	'Call elementEnabled(True)
	'frm1.submit


End Function




</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD5">회사</TD>
									<TD CLASS="TD656" colspan =3>
										<INPUT TYPE=TEXT NAME="txtRegNo" SIZE=13  MAXLENGTH=10 tag="11XXXU" ALT="사업자등록번호"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnUser" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Open_User1()">
										<INPUT TYPE=TEXT NAME="txtRegNm" tag="14X">
									</TD>
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
									<SCRIPT LANGUAGE=JavaScript>ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1><PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
                <TD WIDTH="100%" >
                    <TABLE <%=LR_SPACE_TYPE_30%>>
                        <TR>
                            <TD WIDTH=10>&nbsp;</TD>
                            <TD><BUTTON NAME="btnRegister" CLASS="CLSSBTN" OnClick="VBScript:Call fnRegister()">인증서등록</BUTTON>&nbsp;
                                <BUTTON NAME="btnDown" CLASS="CLSSBTN" OnClick="VBScript:Call fnDown()">Agent 다운로드</BUTTON>&nbsp;                                                                
								<font color="red">* 인증서 등록을 위해서는 Agent프로그램을 먼저 설치하시기 바랍니다.</font>
                                </TD>                                
                            <TD WIDTH=10>&nbsp;</TD>
                        </TR>
                    </TABLE>
                </TD>
            </TR>			
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=No noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

