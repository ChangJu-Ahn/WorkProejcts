
<%@ LANGUAGE="VBSCRIPT" %>
<%'======================================================================================================
'*  1. Module Name          : 법인세 
'*  2. Function Name        : 접대비 입력프로그램 
'*  3. Program ID           :
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 2006-01-25
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : HJO
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%
'=========================  로긴중인 유저의 법인코드를 출력하기 위해  ======================
    Call LoadBasisGlobalInf()
%>
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">



<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliGrid.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"	SRC="../WCM/inc_CliOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'============================================  상수/변수 선언  ====================================

<!-- #Include file="../../inc/lgvariables.inc" -->
Const BIZ_MNU_ID     =  "W3114MA1"	 
Const BIZ_PGM_ID	 =	"W3114MB1.asp"											 '☆: 비지니스 로직 ASP명 

Dim C_SEQ_NO
Dim C_ACCT_GP_NM
Dim C_ACCT_CD
Dim C_ACCT_POP
Dim C_ACCT_NM
Dim C_ACCT_DT
Dim C_ACCT_AMT
Dim C_DOC_TYPE2
Dim C_DOC_TYPE
Dim C_ACCT_DEC
dim strMode



Const C_SHEETMAXROWS = 100

Dim IsOpenPop          
Dim lgStrPrevKey2
Dim lgRefMode	' 레퍼런스 참조상태 : ERP 추출사용유무 

'============================================  초기화 함수  ====================================
Sub InitSpreadPosVariables()
	C_SEQ_NO		= 1

	C_ACCT_GP_NM	= 2
	C_ACCT_CD		= 3
	C_ACCT_POP		= 4
	C_ACCT_NM		= 5
	C_ACCT_DT		= 6
	C_ACCT_AMT		= 7
	C_DOC_TYPE2		= 8
	C_DOC_TYPE		= 9
	C_ACCT_DEC		= 10


	 
End Sub

Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    
    '---- Coding part--------------------------------------------------------------------
    
    lgStrPrevKey = ""                           'initializes Previous Key
    lgStrPrevKey2 = ""                          'initializes Previous Key
    lgLngCurRows = 0                            'initializes Deleted Rows Count
    lgSortKey = 1
    lgRefMode = False
End Sub

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
End Sub

'========================================================================================================
' Function Name : InitComboBox
' Function Desc : This method set cobobox
'========================================================================================================
Sub InitComboBox()
	Dim IntRetCD
	'============================================  버전구분 콤보 박스 채우기  ====================================
	call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W1018' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboREP_TYPE ,lgF0  ,lgF1  ,Chr(11))
    
    '============================================ 계정코드  ====================================
    call CommonQueryRs("ACCT_CD,ACCT_NM"," TB_ACCT_MATCH "," MATCH_CD = '10' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboAcct ,lgF0  ,lgF1  ,Chr(11))
    
        '============================================ 신용카드금액구분  ====================================
    call CommonQueryRs("MINOR_CD,MINOR_NM"," B_MINOR "," MAJOR_CD = 'W2017' ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call SetCombo2(frm1.cboAmtFG ,lgF0  ,lgF1  ,Chr(11)) 
    
End Sub
Sub SetDefaultVal()


End Sub



'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
	Dim strYear
	Dim strMonth
	Dim strInsurDt
	Dim stReturnrInsurDt

   lgKeyStream = UCASE(Frm1.txtCo_Cd.Value) &  parent.gColSep  
   lgKeyStream = lgKeyStream & (Frm1.txtFISC_YEAR.text) &  parent.gColSep ' 
   lgKeyStream = lgKeyStream & UCASE(Frm1.cboREP_TYPE.Value ) &  parent.gColSep '
   lgKeyStream = lgKeyStream &  (frm1.cboACCT.value )   &  parent.gColSep    '계정코드 
   
   if frm1.txtAcct_fg(0).checked Then                                         '사용여부 
	  lgKeyStream = lgKeyStream &  ""  &  parent.gColSep '
   Elseif frm1.txtAcct_fg(1).checked Then 
      lgKeyStream = lgKeyStream &  "Y"  &  parent.gColSep '
   
   Elseif frm1.txtAcct_fg(2).checked Then 
       lgKeyStream = lgKeyStream &  "N"  &  parent.gColSep '
   end if 
	
   lgKeyStream = lgKeyStream &  (Frm1.cboAmtFG.value)   &  parent.gColSep   '금액 
   
   
    if frm1.txtCard_fg(0).checked Then                                         '신용카드사용여부 
	  lgKeyStream = lgKeyStream &  ""  &  parent.gColSep '
   Elseif frm1.txtCard_fg(1).checked Then 
      lgKeyStream = lgKeyStream &  "Y"  &  parent.gColSep '
   
   Elseif frm1.txtCard_fg(2).checked Then 
       lgKeyStream = lgKeyStream &  "N"  &  parent.gColSep '
   end if 

    
 

End Sub 



Sub InitSpreadSheet()

    Call initSpreadPosVariables()       	
	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
		'patch version
		ggoSpread.Spreadinit "V20041222",,parent.gForbidDragDropSpread    
							 
		.ReDraw = false

		.MaxCols = C_ACCT_DEC + 1							'☜: 최대 Columns의 항상 1개 증가시킴 
		.Col = .MaxCols														'☆: 사용자 별 Hidden Column
		.ColHidden = True    
								       
		.MaxRows = 0
		ggoSpread.ClearSpreadData

		Call AppendNumberPlace("6","3","2")
						
		Call GetSpreadColumnPos("A")    

		ggoSpread.SSSetEdit     C_SEQ_NO, "순번", 10,,,10,1
		ggoSpread.SSSetEdit     C_ACCT_GP_NM, "대표계정명",   20,,,100,1
		ggoSpread.SSSetEdit     C_ACCT_CD, "계정코드", 10,,,10,1
		ggoSpread.SSSetButton 	 C_ACCT_POP
		ggoSpread.SSSetEdit     C_ACCT_NM, "계정명",   20,,,100,1
		ggoSpread.SSSetDate     C_ACCT_DT, "일자",     13,2, gDateFormat
		ggoSpread.SSSetFloat    C_ACCT_AMT,           "금액",10, parent.ggAmtOfMoneyNo, ggStrIntegeralPart,  ggStrDeciPointPart, parent.gComNum1000, parent.gComNumDec
		ggoSpread.SSSetCheck    C_DOC_TYPE2, "접대비해당여부", 20,,,True
		ggoSpread.SSSetCheck    C_DOC_TYPE, "신용카드사용여부", 20,,,True
		ggoSpread.SSSetEdit    C_ACCT_DEC, "적요", 50,,,50,1
		Call ggoSpread.SSSetColHidden(.MaxCols, .MaxCols, True)
		Call ggoSpread.SSSetColHidden(C_SEQ_NO, C_SEQ_NO, True)
						
		.ReDraw = true
									 
		Call SetSpreadLock 				
								 
	End With
End Sub


'============================================  그리드 함수  ====================================

Sub InitSpreadComboBox()
    Dim iCodeArr , IntRetCD1
    Dim iNameArr
    Dim iDx



End Sub


Sub SetSpreadLock()


        With frm1
		 	.vspdData.ReDraw = False
	
		 	 ggoSpread.SpreadLock	     C_ACCT_GP_NM,		-1,		C_ACCT_GP_NM 
             ggoSpread.SpreadLock	     C_ACCT_CD,			-1,		C_ACCT_CD 
             ggoSpread.SSSetRequired     C_ACCT_DT,			-1,		C_ACCT_DT 
		 	 ggoSpread.SpreadLock		 C_ACCT_NM,			-1,		C_ACCT_NM
		 	' ggoSpread.SpreadLock		 C_ACCT_DEC,	    -1,		C_ACCT_DEC
		 	 ggoSpread.SSSetRequired	 C_ACCT_AMT,	    -1,		C_ACCT_AMT
				
		 	.vspdData.ReDraw = True

		 End With
End Sub

Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)
dim sumRow
    ggoSpread.Source = frm1.vspdData
    With frm1

    .vspdData.ReDraw = False
       
        ggoSpread.SSSetProtected  C_ACCT_GP_NM , pvStartRow, pvEndRow
        ggoSpread.SSSetRequired   C_ACCT_CD ,	 pvStartRow, pvEndRow	 
	    ggoSpread.SSSetRequired   C_ACCT_DT ,	 pvStartRow, pvEndRow
	    ggoSpread.SSSetProtected  C_ACCT_NM ,	 pvStartRow, pvEndRow
        ggoSpread.SSSetRequired  C_ACCT_AMT ,	C_ACCT_DEC, pvEndRow 
    .vspdData.ReDraw = True
    
    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr, parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <>  parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to 
              Exit For
           End If           
       Next          
    End If   
End Sub


Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos

    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_SEQ_NO		= iCurColumnPos(1)

				C_ACCT_GP_NM	= iCurColumnPos(2)
				C_ACCT_CD		= iCurColumnPos(3)
				C_ACCT_POP		= iCurColumnPos(4)
				C_ACCT_NM		= iCurColumnPos(5)
				C_ACCT_DT		= iCurColumnPos(6)
				C_ACCT_AMT		= iCurColumnPos(7)
				C_DOC_TYPE2		= iCurColumnPos(8)
				C_DOC_TYPE		= iCurColumnPos(9)
				C_ACCT_DEC		= iCurColumnPos(10)

    End Select    
End Sub

'============================================  조회조건 함수  ====================================
'======================================================================================================
'   Function Name : OpenPopUp()
'   Function Desc : 
'=======================================================================================================
Function  OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strCd
	
	If IsOpenPop = True Then Exit Function
	
	Select Case iWhere
		Case 0

		Case 2
	
		Case Else
			Exit Function
	End Select

	IsOpenPop = True
			
		arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	IsOpenPop = False
	
	If arrRet(0) = "" Then	    
		Exit Function
	Else
		Call SetPopup(arrRet, iWhere)
	End If
End Function
'======================================================================================================
'   Function Name : SetPopup(Byval arrRet)
'   Function Desc : 
'=======================================================================================================



Function SetPopup(Byval arrRet,Byval iWhere)
	With frm1
		Select Case iWhere
			Case 0
			Case 2
				.vspdData.Col = C_ACCT_CD
				.vspdData.Text = arrRet(0)
				.vspdData.Col = C_ACCT_NM
				.vspdData.Text = arrRet(1)
				
				Call vspdData_Change(C_ACCT_CD, frm1.vspdData.activerow )	 ' 변경이 읽어났다고 알려줌 
		
		End Select
	End With
	If iwhere  <> 0 Then
		lgBlnFlgChgValue = True
	End If
End Function



'============================================  폼 함수  ====================================
Sub Form_Load()

    Call LoadInfTB19029     
                                                    <%'Load table , B_numeric_format%>
    Call ggoOper.LockField(Document, "N")                                   <%'Lock  Suitable  Field%>                         
   
    Call InitSpreadSheet()                                                    <%'Setup the Spread sheet%>
    Call InitVariables                                                      <%'Initializes local global variables%>
    
    Call SetToolbar("1100110000000111")										<%'버튼 툴바 제어 %>

	' 변경한곳 
    Call InitComboBox
   
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)

    Call ggoOper.FormatDate(frm1.txtFISC_YEAR, parent.gDateFormat,3)

   frm1.txtFISC_YEAR.text 	= "<%=wgFISC_YEAR%>"

	frm1.txtCO_CD.value 	= "<%=wgCO_CD%>"
	frm1.txtCO_NM.value 	= "<%=wgCO_NM%>"
	frm1.cboREP_TYPE.value 	= "<%=wgREP_TYPE%>"
     

    Call FncQuery
End Sub


'============================================  이벤트 함수  ====================================
Function GetRef()	
    Dim IntRetCD , i
    Dim sMesg
   
   
    'Call ggoOper.ClearField(Document, "2")
   
	if wgConfirmFlg = "Y" then    '확정시 
	   Exit function
	end if   
	

	sMesg = wgRefDoc & vbCrLf & vbCrLf
    
    IntRetCD = DisplayMsgBox("WC0003", parent.VB_YES_NO, sMesg, "X")           '⊙: "Will you destory previous data"
	If IntRetCD = vbNo Then
		Exit Function
	End If
	
	
     
     CALL getdata()
    	
End Function





Function GetData()	

	Dim IntRetCD1
	dim strWhere
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	    
    strWhere = FilterVar(Trim(frm1.txtCO_CD.value ),"","S")  
    strWhere = strWhere & " ," & FilterVar(Trim(frm1.txtFISC_YEAR.text ),"","S")
    strWhere = strWhere & " ," & FilterVar(Trim(frm1.cboREP_TYPE.value ),"","S") 
	
	
	call CommonQueryRs("w1,w2,w3,w10,w13,w14"," dbo.ufn_TB_33_GetRef("& strWhere &")" ,,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
          frm1.txtW1.value= unicdbl(replace(lgF0 ,Chr(11),""))
          frm1.txtW2.value= unicdbl(replace(lgF1 ,Chr(11),""))
          frm1.txtW3.value= unicdbl(replace(lgF2 ,Chr(11),""))
          frm1.txtW10.value=unicdbl(replace(lgF3 ,Chr(11),""))
          frm1.txtW13.value=unicdbl(replace(lgF4 ,Chr(11),""))
          frm1.txtW14.value=unicdbl(replace(lgF5 ,Chr(11),""))
   

    call CommonQueryRs("w18,w19,w20"," dbo.ufn_TB_33_GetRef  ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
         
          frm1.txtW18.value= unicdbl(replace(lgF0 ,Chr(11),""))
          frm1.txtW19.value= unicdbl(replace(lgF1 ,Chr(11),""))
          frm1.txtW20.value= unicdbl(replace(lgF2 ,Chr(11),""))
          

End Function


' 전체선택 
'========================================
Sub chkSelectAll1_onClick()
	Dim iStrOldValue
	Dim i
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub

	ggoSpread.Source = frm1.vspdData	
	With frm1.vspdData
		.Row = 1			:	.Row2 = .MaxRows
		
		' 전체선택 
		If frm1.chkSelectAll1.checked Then
			
			For i=1 to .MaxRows
				' 선택버튼의 선택여부 설정 
				.Col = C_DOC_TYPE2		:	.Col2 = C_DOC_TYPE2
				.Row=i :.Row2 = i
				If .text="0" Then				
					.Clip = Replace(.Clip, "0", "1")
					' Row Header 설정(수정)
					.Col=0 : .Col2=0
					If trim(.Text)<>"" Then 
					.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")								
					Else
					.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")		
					.Clip = Replace(.Clip, vbCrLf, ggoSpread.UpdateFlag & vbCrLf)	
					End IF	
				End If
			Next					
		' 전체선택 취소 
		Else		
			For i=1 to .MaxRows
				' 선택버튼의 선택여부 설정 
				.Col = C_DOC_TYPE2		:	.Col2 = C_DOC_TYPE2
				.Row=i :.Row2 = i
				If .text="1" Then				
					.Clip = Replace(.Clip, "1", "0")					
					' Row Header 설정(수정)					
					.Col=0 : .Col2=0
					If trim(.Text)<>"" Then 
					.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")
								
					Else
					.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")		
					.Clip = Replace(.Clip, vbCrLf, ggoSpread.UpdateFlag & vbCrLf)	
					End IF
				End If
			Next				
		End if		
	End With
End Sub


' 전체선택 
'========================================
Sub chkSelectAll2_onClick()
	Dim iStrOldValue,i
	
	If frm1.vspdData.MaxRows = 0 Then Exit Sub

	ggoSpread.Source = frm1.vspdData	
	With frm1.vspdData
		.Row = 1			:	.Row2 = .MaxRows
		
		' 전체선택 
		If frm1.chkSelectAll2.checked Then
			For i=1 to .MaxRows
				' 선택버튼의 선택여부 설정 
				.Col = C_DOC_TYPE		:	.Col2 = C_DOC_TYPE
				.Row=i :.Row2 = i
				If .text="0" Then				
					.Clip = Replace(.Clip, "0", "1")
					' Row Header 설정(수정)
					.Col=0 : .Col2=0
					If trim(.Text)<>"" Then 
					.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")								
					Else
					.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")		
					.Clip = Replace(.Clip, vbCrLf, ggoSpread.UpdateFlag & vbCrLf)	
					End IF	
				End If
			Next					
			
		' 전체선택 취소 
		Else
			For i=1 to .MaxRows
				' 선택버튼의 선택여부 설정 
				.Col = C_DOC_TYPE		:	.Col2 = C_DOC_TYPE
				.Row=i :.Row2 = i
				If .text="1" Then				
					.Clip = Replace(.Clip, "1", "0")					
					' Row Header 설정(수정)					
					.Col=0 : .Col2=0
					If trim(.Text)<>"" Then 
					.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")
								
					Else
					.Clip = Replace(.Clip, ggoSpread.UpdateFlag, "")		
					.Clip = Replace(.Clip, vbCrLf, ggoSpread.UpdateFlag & vbCrLf)	
					End IF
				End If
			Next				  
		End if
	End With

End Sub




'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim strTemp
	Dim intPos1


	
	With frm1.vspdData
      frm1.vspdData.Row = Row
      frm1.vspdData.Col = Col
     If Row <= 0 then Exit sub 
	Select Case Col
	    Case C_ACCT_POP 
           
		    frm1.vspdData.Col = Col - 1
		    frm1.vspdData.Row = Row
		
           Call OpenGPACCT(frm1.vspdData.Text,Col ,  Row)
	 
    End Select
        
    End With
End Sub

Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex

	
End Sub


Function OpenGPACCT(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strWhere

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "TB_ACCT_MATCH"					<%' TABLE 명칭 %>
	arrParam(2) = strCode		<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	strWhere = " MATCH_CD = '10'"
	strWhere = strWhere & " AND CO_CD = '" & frm1.txtCO_CD.value & "' "
	strWhere = strWhere & " AND FISC_YEAR = '" & frm1.txtFISC_YEAR.text & "' "
	strWhere = strWhere & " AND REP_TYPE = '" & frm1.cboREP_TYPE.value & "' "

	arrParam(4) = strWhere							<%' Where Condition%>
	arrParam(5) = "계정"						<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "ED7" & Chr(11) & "ACCT_CD" & Chr(11)					<%' Field명(2)%>
    arrField(1) = "ED20" & Chr(11) & "ACCT_NM" & Chr(11)					<%' Field명(3)%>
    arrField(2) = "ED20" & Chr(11) & "dbo.ufn_GetCodeName('W1056', ACCT_GP_CD)" & Chr(11)					<%' Field명(1)%>
   
    
    arrHeader(0) = "계정코드"					<%' Header명(0)%>
    arrHeader(1) = "계정명"						<%' Header명(1)%>
    arrHeader(2) = "대표계정명"					<%' Header명(2)%>
    arrHeader(3) = ""						<%' Header명(3)%>
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=520px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetGPACCT(arrRet,iWhere)
	End If	
	
End Function



'======================================================================================================
'	Name : SetCode()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Function SetGPACCT(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		    Case C_ACCT_POP   
		    	.vspdData.Col = C_ACCT_GP_NM
		    	.vspdData.text = arrRet(2)   
		        .vspdData.Col = C_ACCT_CD
		    	.vspdData.text = arrRet(0)
		    	.vspdData.Col = C_ACCT_NM
		    	.vspdData.text = arrRet(1)  
		    	.vspdData.action =0
		
        End Select

	End With

End Function

Function OpenAccount(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
	Dim strWhere

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = "계정 팝업"					<%' 팝업 명칭 %>
	arrParam(1) = "TB_ACCT_MATCH"					<%' TABLE 명칭 %>
	arrParam(2) = strCode		<%' Code Condition%>
	arrParam(3) = ""							<%' Name Cindition%>
	strWhere = " MATCH_CD = '10'"
	strWhere = strWhere & " AND CO_CD = '" & frm1.txtCO_CD.value & "' "
	strWhere = strWhere & " AND FISC_YEAR = '" & frm1.txtFISC_YEAR.text & "' "
	strWhere = strWhere & " AND REP_TYPE = '" & frm1.cboREP_TYPE.value & "' "

	arrParam(4) = strWhere							<%' Where Condition%>
	arrParam(5) = "계정"						<%' 조건필드의 라벨 명칭 %>
	
    arrField(0) = "ED5" & Chr(11) & "ACCT_GP_CD" & Chr(11)					<%' Field명(0)%>
    arrField(1) = "ED10" & Chr(11) & "dbo.ufn_GetCodeName('W1085', ACCT_GP_CD)" & Chr(11)					<%' Field명(1)%>
    arrField(2) = "ED7" & Chr(11) & "ACCT_CD" & Chr(11)					<%' Field명(2)%>
    arrField(3) = "ED15" & Chr(11) & "ACCT_NM" & Chr(11)					<%' Field명(3)%>
    
    arrHeader(0) = "대표계정코드"					<%' Header명(0)%>
    arrHeader(1) = "대표계정명"						<%' Header명(1)%>
    arrHeader(2) = "계정코드"					<%' Header명(2)%>
    arrHeader(3) = "계정명"						<%' Header명(3)%>
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetAccount(arrRet,iWhere)
	End If	
	
End Function

Function SetAccount(byval arrRet,Byval iWhere)
    With frm1
		If iWhere = 1 Then 'Spread1(Condition)
			.vspdData.Col = C_W16
			.vspdData.Text = arrRet(0)
			.vspdData.Col = C_W16_NM
			.vspdData.Text = arrRet(1)
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
			lgBlnFlgChgValue = True
		ElseIf iWhere = 2 Then 'Spread2(Condition)
			.vspdData2.Col = C_W23
			.vspdData2.Text = arrRet(0)
			.vspdData2.Col = C_W23_NM
			.vspdData2.Text = arrRet(1)
		    ggoSpread.Source = frm1.vspdData
		    ggoSpread.UpdateRow frm1.vspdData.ActiveRow
			lgBlnFlgChgValue = True
		End If
	End With
End Function


'==========================================================================================
Sub InitData()
	Dim intRow
	Dim intIndex 


		
        frm1.txtTotalAmt.value = FncSumSheet(frm1.vspdData,C_ACCT_AMT, 1, frm1.vspdData.MaxRows,FALSE , -1, -1, "V")

End Sub

Sub vspdData_Change(ByVal Col , ByVal Row )
     Dim iDx
    Dim IntRetCD,strWhere
    Dim i
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
  '------ Developer Coding part (Start ) -------------------------------------------------------------- 

 
  '--------------------'그리드에 입력된 내역이 기존데이터에 있을때 체크----------------------------------
    Select Case Col
      case C_ACCT_CD
				strWhere = " AND MATCH_CD = '10'"
				strWhere = strWhere & " AND CO_CD = '" & frm1.txtCO_CD.value & "' "
				strWhere = strWhere & " AND FISC_YEAR = '" & frm1.txtFISC_YEAR.text & "' "
				strWhere = strWhere & " AND REP_TYPE = '" & frm1.cboREP_TYPE.value & "' "
				    
				   Frm1.vspdData.Row = Row
				  Frm1.vspdData.Col = Col
'				If CommonQueryRs("ACCT_NM", " TB_WORK_6 (NOLOCK)" , "ACCT_CD = '" & Frm1.vspdData.Text &"' AND ACCT_CD IN (SELECT ACCT_CD FROM TB_ACCT_MATCH (NOLOCK) WHERE MATCH_CD = '07')", lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				If CommonQueryRs("dbo.ufn_GetCodeName('W1085', ACCT_GP_CD) ,ACCT_NM", " TB_ACCT_MATCH " , "ACCT_CD= '" & Frm1.vspdData.Text &"'" & strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			    	frm1.vspdData.Col	= C_ACCT_GP_NM
			    	frm1.vspdData.Text	= replace(lgF0, Chr(11),"")
			    	frm1.vspdData.Col	= C_ACCT_NM
					frm1.vspdData.Text	= Replace(lgF1, Chr(11),"")
				Else
		
					frm1.vspdData.Col	= C_ACCT_GP_NM
					frm1.vspdData.Text	= ""
					frm1.vspdData.Col	= C_ACCT_CD
					frm1.vspdData.Text	= ""
					frm1.vspdData.Col	= C_ACCT_NM
					frm1.vspdData.Text	= ""
				End If
		
'		End If

	
        
    End Select
    
 '------ Developer Coding part (End   ) -------------------------------------------------------------- 
    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
      'If CDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
      '   Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      'End If
    End If
	
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
   	    
   ' If Row <= 0 Then
   '    ggoSpread.Source = frm1.vspdData
       
   '    If lgSortKey = 1 Then
   '        ggoSpread.SSSort Col               'Sort in ascending
   '        lgSortKey = 2
   '    Else
   '        ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
   '        lgSortKey = 1
   '    End If
       
   '    Exit Sub
   ' End If

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

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If

	If CheckRunningBizProcess = True Then
	   Exit Sub
	End If
    
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	           
    	If lgStrPrevKeyIndex <> "" Then                         
      	   Call DisableToolBar(Parent.TBC_QUERY)
			If DbQuery = False Then
				Call RestoreTooBar()
			    Exit Sub
			End If  				
    	End If
    End if
End Sub

Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
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







'============================================  툴바지원 함수  ====================================
'=====================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          
    
	ggoSpread.Source = frm1.vspdData 
	If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO, "X", "X") 
		If IntRetCD = vbNo Then	Exit Function
    End If

     Call ggoOper.ClearField(Document, "2")
    Call SetDefaultVal
    Call InitVariables               

    Call SetToolbar("1100100100000111")          '⊙: 버튼 툴바 제어 
    FncNew = True                

End Function

'=====================================================
Function FncQuery() 

    Dim IntRetCD 
    
    On Error Resume next
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status

     ggoSpread.Source = Frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    ggoSpread.ClearSpreadData
    															
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If


    Call InitVariables                                                           '⊙: Initializes local global variables
    
    Call MakeKeyStream("X")
    frm1.chkSelectAll1.checked = False
    frm1.chkSelectAll2.checked = False
	Call  DisableToolBar( parent.TBC_QUERY)
	If DbQuery = False Then
		Call  RestoreToolBar()
		Exit Function
	End If			
              
    FncQuery = True 
    Set gActiveElement = document.ActiveElement

End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim IntRetCd

    FncDelete = False                                                             '☜: Processing is NG
    
    
    <%  '-----------------------
    'Check previous data area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO, "X", "X")			    <%'데이타가 변경되었습니다. 조회하시겠습니까?%>
    	If IntRetCD = vbNo Then
      	Exit Function
    	End If
    End If
    
    
    
    If lgIntFlgMode <>  parent.OPMD_UMODE  Then                                            'Check if there is retrived data
        Call  DisplayMsgBox("900002","X","X","X")                                  '☜: Please do Display first.
        Exit Function
    End If

    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"X","X")		                  '☜: Do you want to delete?
	If IntRetCD = vbNo Then
		Exit Function
	End If

    If DbDelete= False Then
       Exit Function
    End If												                  '☜: Delete db data

    FncDelete=  True                                                              '☜: Processing is OK
End Function





Function FncSave() 
   dim IntRetCD
    FncSave = False                                                         
    
    
    
    

    Err.Clear                                                               <%'☜: Protect system from crashing%>
    'On Error Resume Next                                                   <%'☜: Protect system from crashing%>    
    
<%  '-----------------------
    'Precheck area
    '----------------------- %>
    ggoSpread.Source = frm1.vspdData
    if frm1.vspdData.maxrows = 0 then 
       IntRetCD =  DisplayMsgBox("WC0002","x","x","x")                           '☜:There is no changed data. 
	    Exit Function
    end if 
    
    If ggoSpread.SSCheckChange = False   Then
                 ggoSpread.Source = frm1.vspdData
			If  ggoSpread.SSCheckChange = False Then
			    IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
			    Exit Function
			End If
       
    End If
    ggoSpread.Source = frm1.vspdData
	If Not  ggoSpread.SSDefaultCheck Then                                         '☜: Check contents area
	      Exit Function
	End If  
    
   Call MakeKeyStream("X")
<%  '-----------------------
    'Save function call area
    '----------------------- %>
    If DbSave = False Then Exit Function                                        '☜: Save db data
    
    FncSave = True                                                          
    
End Function

Function FncCopy() 
    On Error Resume Next                                                          '☜: If process fails
    Err.Clear                                                                     '☜: Clear error status

    FncCopy = False                                                               '☜: Processing is NG

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
 
    ggoSpread.Source = Frm1.vspdData

	With frm1
		If .vspdData.ActiveRow > 0 and .vspdData.ActiveRow <> .vspdData.maxrows Then
			.vspdData.focus
			.vspdData.ReDraw = False
		
			ggoSpread.CopyRow
			SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow

			
			.vspdData.ReDraw = True
		End If
	End With
	
    If Err.number = 0 Then	
       FncCopy = True                                                            '☜: Processing is OK
    End If
	
    Set gActiveElement = document.ActiveElement   
	
End Function

Function FncCancel() 
     ggoSpread.Source = Frm1.vspdData	
     ggoSpread.EditUndo  
       
    
End Function

Function FncInsertRow(ByVal pvRowCnt) 
    Dim imRow    

    On Error Resume Next 
    Err.Clear 
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
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1            
        
       .vspdData.ReDraw = True
    End With

    If Err.number = 0 Then
       FncInsertRow = True                                                          '☜: Processing is OK
    End If   
    
    Set gActiveElement = document.ActiveElement   
    
End Function

Function FncDeleteRow() 
   If frm1.vspdData.MaxRows < 1 Then Exit Function

    Dim lDelRows
    Dim iDelRowCnt, i
    
    With frm1  

    .vspdData.focus
    ggoSpread.Source = .vspdData 
 
       lDelRows = ggoSpread.DeleteRow                                              '☜: Protect system from crashing

	
	
 
    lgBlnFlgChgValue = True
    
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
    If ggoSpread.SSCheckChange = True or lgBlnFlgChgValue = true Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO, "X", "X")                <%'데이타가 변경되었습니다. 종료 하시겠습니까?%>
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function


'============================================  DB 억세스 함수  ====================================
Function DbQuery() 

    DbQuery = False
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
	
	Call LayerShowHide(1)
	
	Dim strVal
 
       
        With Frm1
			strVal = BIZ_PGM_ID & "?txtMode="            &  parent.UID_M0001						         
			strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
			strVal = strVal     & "&txtMaxRows="         & .vspdData.MaxRows
			strVal = strVal     & "&lgStrPrevKey="  & lgStrPrevKey                 '☜: Next key tag
		End With



		Call RunMyBizASP(MyBizASP, strVal)   

    DbQuery = True  
  
End Function

Function DbQueryOk()													<%'조회 성공후 실행로직 %>
	
    Dim iNameArr
    Dim iDx    
    Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE

   									<%'버튼 툴바 제어 %>
    Call CheckTaxDoc(frm1.txtCO_CD.value, frm1.txtFISC_YEAR.text, frm1.cboREP_TYPE.value, BIZ_MNU_ID)
    Call InitData
	'1 컨펌체크 
	If wgConfirmFlg = "Y" Then

	    Call SetToolbar("1100000000000111")	
		
	Else
	   '2 디비환경값 , 로드시환경값 비교 
		 Call SetToolbar("1100111100000111")										<%'버튼 툴바 제어 %>
	
	End If
  
  

	frm1.vspdData.focus			
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status

	DbDelete = False			                                                 '☜: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If

	strVal = BIZ_PGM_ID & "?txtMode=" &  parent.UID_M0003                                '☜: Delete
	strVal = strVal & "&txtCo_Cd=" & Trim(frm1.txtCo_Cd.value)				'☆: 조회 조건 데이타 
    strVal = strVal & "&txtFISC_YEAR=" & Trim(frm1.txtFISC_YEAR.text)				'☆: 조회 조건 데이타 
	strVal = strVal & "&cboREP_TYPE=" & Trim(frm1.cboREP_TYPE.value)				'☆: 조회 조건 데이타 
	strVal = strVal		& "&lgStrPrevKey=" & lgStrPrevKey

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	DbDelete = True                                                              '⊙: Processing is NG

End Function
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave() 
   Dim pP21011
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    DbSave = False                                                          
    
	if LayerShowHide(1) = false then
	exit Function
	end if

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
    
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text


               Case  ggoSpread.InsertFlag                                      '☜: Update추가 
												
                     lGrpCnt = lGrpCnt + 1
                     								  strVal = strVal & "C"  &  parent.gColSep
													  strVal = strVal & lRow &  parent.gColSep
					.vspdData.Col = C_ACCT_CD		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
					.vspdData.Col = C_ACCT_NM		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_ACCT_DT		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep       
                    .vspdData.Col = C_ACCT_AMT		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep       
                    .vspdData.Col = C_ACCT_DEC		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep          
                    .vspdData.Col = C_DOC_TYPE2		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep       
                    .vspdData.Col = C_DOC_TYPE		: strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep
                    
                     
               Case  ggoSpread.UpdateFlag     
                                             '☜: Update
													  strVal = strVal & "U"  &  parent.gColSep
													  strVal = strVal & lRow &  parent.gColSep
					.vspdData.Col = C_SEQ_NO		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep													  
                    .vspdData.Col = C_ACCT_CD		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_ACCT_DT		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep       
                    .vspdData.Col = C_ACCT_AMT		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep       
                    .vspdData.Col = C_ACCT_DEC		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep          
                    .vspdData.Col = C_DOC_TYPE2		: strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep       
                    .vspdData.Col = C_DOC_TYPE		: strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep  
                    lGrpCnt = lGrpCnt + 1
                                        
               Case  ggoSpread.DeleteFlag                                      '☜: Delete
														strDel = strDel & "D"  &  parent.gColSep
													  strDel = strDel & lRow &  parent.gColSep
					.vspdData.Col = C_SEQ_NO		: strDel = strDel & Trim(.vspdData.Text) &  parent.gRowSep		
					  lGrpCnt = lGrpCnt + 1

           End Select
       Next

       .txtMode.value        =  parent.UID_M0002
       .txtKeyStream.value   =  lgKeyStream
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal

	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)	
    DbSave = True             

                                                   
End Function


Function DbSaveOk()													        <%' 저장 성공후 실행 로직 %>
	Call InitVariables
	frm1.vspdData.MaxRows = 0
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData

    Call MainQuery()
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
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP_BAK">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" NOWRAP><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB_BAK"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS="TD5">사업연도</TD>
									<TD CLASS="TD6"><script language =javascript src='./js/w3114ma1_txtFISC_YEAR_txtFISC_YEAR.js'></script>
									<TD CLASS="TD5">법인명</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtCO_CD" Size=10 tag="14">
										<INPUT TYPE=TEXT NAME="txtCO_NM" Size=20 tag="14">
									</TD>
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">버전구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboREP_TYPE" ALT="버전구분" STYLE="WIDTH: 50%" tag="14"></SELECT>
									</TD>
									<TD CLASS="TD5"></TD>
									<TD CLASS="TD6">
									</TD>
								</TR>
								<TR>
									<TD CLASS="TD5">접대비여부</TD>
								
				                	<TD CLASS="TD6"> <INPUT TYPE="RADIO" CLASS="RADIO" ID="txtAcct_fg0" NAME="txtAcct_fg" TAG="21X" VALUE="%" CHECKED><LABEL FOR="txtAcct_fg0">전체</LABEL>&nbsp;&nbsp;&nbsp;
				                	  				 <INPUT TYPE="RADIO" CLASS="RADIO" ID="txtAcct_fg1" NAME="txtAcct_fg" TAG="21X" VALUE="Y" ><LABEL FOR="txtAcct_fg1">여</LABEL>&nbsp;&nbsp;&nbsp;
				                			         <INPUT TYPE="RADIO" CLASS="RADIO" ID="txtAcct_fg2" NAME="txtAcct_fg" TAG="21X" VALUE="N"><LABEL FOR="txtAcct_fg2">부</LABEL></TD>
                                        
									<TD CLASS="TD5">신용카드여부</TD>
								
				                	<TD CLASS="TD6"> <INPUT TYPE="RADIO" CLASS="RADIO" ID="txtCard_fg0" NAME="txtCard_fg" TAG="21X" VALUE="%" CHECKED><LABEL FOR="txtCard_fg0">전체</LABEL>&nbsp;&nbsp;&nbsp;
				                	  				 <INPUT TYPE="RADIO" CLASS="RADIO" ID="txtCard_fg1" NAME="txtCard_fg" TAG="21X" VALUE="Y" ><LABEL FOR="txtCard_fg1">여</LABEL>&nbsp;&nbsp;&nbsp;
				                			         <INPUT TYPE="RADIO" CLASS="RADIO" ID="txtCard_fg2" NAME="txtCard_fg" TAG="21X" VALUE="N"><LABEL FOR="txtCard_fg2">부</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">계정</TD>
									<TD CLASS="TD6"><SELECT NAME="cboACCT" ALT="계정" STYLE="WIDTH: 50%" tag="1X"><OPTION VALUE="">전체</OPTION></SELECT>
									</TD>
									<TD CLASS="TD5">신용카드 금액구분</TD>
									<TD CLASS="TD6"><SELECT NAME="cboAmtFG" ALT="신용카드 금액구분" STYLE="WIDTH: 50%" tag="1X"><OPTION VALUE="%">전체</OPTION></SELECT>
									</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>일괄선택</TD>
									<TD CLASS=TD6 NOWRAP>
										접대비 사용여부<INPUT TYPE=CHECKBOX NAME="chkSelectAll1" ID="chkSelectAll1" tag="21" Class="Check">
										신용카드 사용여부<INPUT TYPE=CHECKBOX NAME="chkSelectAll2" ID="chkSelectAll2" tag="21" Class="Check">
									</TD>
									<TD CLASS=TD5 NOWRAP></TD>
									<TD CLASS=TD6 NOWRAP>
									
									</TD>
								
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%> </TD>
				</TR>
				
	
				<TR>
				    
						<TD WIDTH=100%  HEIGHT=*  valign=top >
						   
										<TABLE <%=LR_SPACE_TYPE_20%>>
										     
											       
													<TR>
														<TD HEIGHT="100%">
															<script language =javascript src='./js/w3114ma1_vaSpread1_vspdData.js'></script>
														</TD>
														
													</TR>
													<TR>
													<TD WIDTH=100% HEIGHT=*   VALIGN=TOP>
													    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=LEFT></LEGEND>
													        <TABLE WIDTH=100% HEIGHT="100%" CELLSPACING=0>
	 														   <TR>  
															    <TD CLASS=TDT NOWRAP>
															       합계<script language =javascript src='./js/w3114ma1_txtTotalAmt_txtTotalAmt.js'></script></TD>
														      </TR>
														    </TABLE>
													    </FIELDSET>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS=hidden NAME=txtSpread tag="24" tabindex="-1"></TEXTAREA>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

