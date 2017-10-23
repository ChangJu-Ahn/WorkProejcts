<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        :                   
'*  3. Program ID           : B1263MA1                    
'*  4. Program Name         : 사업자이력등록                  
'*  5. Program Desc         : 사업자이력등록               
'*  6. Comproxy List        : PB5CS41.dll, PB5CS44.dll, PB5CS45.dll
'*  7. Modified date(First) : 2001/01/05                
'*  8. Modified date(Last)  : 2001/01/05                
'*  9. Modifier (First)     : Kim Hyungsuk                
'* 10. Modifier (Last)      : Sonbumyeol  
'* 11. Comment              :               
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									
'*                            this mark(⊙) Means that "may  change"									
'*                            this mark(☆) Means that "must change"									
'* 13. History              : 2002/12/02 : INCLUDE 추가 성능 적용, Kang Jun Gu
'*                            2002/12/09 : INCLUDE 다시 성능 적용, Kang Jun Gu
'********************************************************************************************************
%>

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="javascript"   SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit                                                             '☜: indicates that All variables must be declared in advance

<!-- #Include file="../../inc/lgvariables.inc" -->	

Dim  EndDate,  StartDate 
EndDate = UNIConvDateAtoB("<%=GetSvrDate%>", parent.gServerDateFormat, parent.gDateFormat)
StartDate = UnIDateAdd("m", -1, EndDate, parent.gDateFormat)

Const BIZ_PGM_ID = "b1263mb1.asp" 
Const BIZ_PGM_JUMP_ID1 = "b1261ma1" 
Const BIZ_PGM_JUMP_ID2 = "b1263ma8"

Dim IsOpenPop      ' Popup

'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed
    lgIntGrpCount = 0                           'initializes Group View Size
    '---- Coding part--------------------------------------------------------------------
    
End Sub

'========================================================================================================= 
Sub SetDefaultVal()
 frm1.txtConBp_cd.focus 
 frm1.txtConValidFromDt.Text = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat) 
 frm1.txtValidFromDt.Text    = frm1.txtConValidFromDt.Text
End Sub

'========================================================================================================= 
<% '== 등록 == %>
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A( "I", "*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================= 
Function OpenConBp_cd()

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 arrParam(0) = "거래처"			<%' 팝업 명칭 %>
 arrParam(1) = "B_BIZ_PARTNER"      <%' TABLE 명칭 %>
 arrParam(2) = Trim(frm1.txtConBp_cd.Value)		<%' Code Condition%>
 arrParam(3) = ""								<%' Name Cindition%>
 arrParam(4) = ""								<%' Where Condition%>
 arrParam(5) = "거래처"						<%' TextBox 명칭 %>
 
 arrField(0) = "BP_CD"        <%' Field명(0)%>
 arrField(1) = "BP_NM"        <%' Field명(1)%>
    
 arrHeader(0) = "거래처"				<%' Header명(0)%>
 arrHeader(1) = "거래처약칭"           <%' Header명(1)%>
    
 frm1.txtConBp_cd.focus 
 
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function
 Else
  Call SetConBp_cd(arrRet)
 End If 
 
End Function

'========================================================================================================= 
Function OpenBp_cd(Byval iWhere)

 Dim arrRet
 Dim arrParam(5), arrField(6), arrHeader(6)

 If IsOpenPop = True Then Exit Function

 IsOpenPop = True

 Select Case iWhere
 Case 0
	If frm1.txtBp_cd.readOnly = True Then
	 IsOpenPop = False
	 Exit Function
	End If
	arrParam(0) = "거래처"			<%' 팝업 명칭 %>
	arrParam(1) = "B_BIZ_PARTNER"					<%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtBp_cd.Value)			<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "거래처"					<%' TextBox 명칭 %>
	 
	arrField(0) = "BP_CD"        <%' Field명(0)%>
	arrField(1) = "BP_NM"        <%' Field명(1)%>
	    
	arrHeader(0) = "거래처"				<%' Header명(0)%>
	arrHeader(1) = "거래처약칭"           <%' Header명(1)%>
	
	frm1.txtBp_cd.focus
	
 Case 1            <%' 업태 %>
	arrParam(0) = "업태"      <%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"       <%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtInd_Class.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("B9003", "''", "S") & " "    <%' Where Condition%>
	arrParam(5) = "업태"            <%' TextBox 명칭 %>
  
	arrField(0) = "MINOR_CD"      <%' Field명(0)%>
	arrField(1) = "MINOR_NM"      <%' Field명(1)%>
		     
	arrHeader(0) = "업태"			<%' Header명(0)%>
	arrHeader(1) = "업태명"           <%' Header명(1)%>

	frm1.txtInd_Class.focus
 Case 2            <%' 업종 %>
	arrParam(0) = "업종"      <%' 팝업 명칭 %>
	arrParam(1) = "B_MINOR"       <%' TABLE 명칭 %>
	arrParam(2) = Trim(frm1.txtInd_Type.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("B9002", "''", "S") & " "    <%' Where Condition%>
	arrParam(5) = "업종"            <%' TextBox 명칭 %>
		  
	arrField(0) = "MINOR_CD"      <%' Field명(0)%>
	arrField(1) = "MINOR_NM"      <%' Field명(1)%>
		     
	arrHeader(0) = "업종"			<%' Header명(0)%>
	arrHeader(1) = "업종명"           <%' Header명(1)%>

	frm1.txtInd_Type.focus
 End Select
    
 arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
  "dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

 IsOpenPop = False

 If arrRet(0) = "" Then
  Exit Function  
 Else  
  Call SetBp_cd(arrRet, iWhere)
 End If 
 
End Function

'========================================================================================================= 
Function OpenZip()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function
	
	IsOpenPop = True
	
	arrParam(0) = Trim(frm1.txtZIP_cd.value)
	arrParam(1) = ""
	arrParam(2) = Parent.gCountry

	frm1.txtZIP_cd.focus 
	
	arrRet = window.showModalDialog("../../comasp/ZipPopup.asp", Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetZip(arrRet)
	End If	
			
End Function

'========================================================================================================= 
Function SetConBp_cd(Byval arrRet)

 With frm1
  .txtConBp_cd.value = arrRet(0) 
  .txtConBp_nm.value = arrRet(1)   
 End With
 
End Function

'========================================================================================================= 
Function SetBp_cd(Byval arrRet, Byval iWhere)

 With frm1

  Select Case iWhere
  Case 0
   .txtBp_cd.value = arrRet(0) 
   .txtBp_nm.value = arrRet(1)   
   Call BpLookUp()
   .txtBp_cd.focus
  Case 1
   .txtInd_Class.value = arrRet(0) 
   .txtInd_ClassNm.value = arrRet(1)   
   .txtInd_Class.focus
  Case 2
   .txtInd_Type.value = arrRet(0) 
   .txtInd_TypeNm.value = arrRet(1)   
   .txtInd_Type.focus
  End Select

   lgBlnFlgChgValue = True

 End With
 
End Function

'========================================================================================================= 
Sub SetZip(arrRet)
	With frm1
		.txtZIP_cd.value = arrRet(0)
		.txtADDR1.value = arrRet(1)
		.txtADDR2.value = ""
		lgBlnFlgChgValue = True
	
		.txtADDR2.focus 
	End With
End Sub

'========================================================================================
' Function Desc : 한글을 구분한다.
'========================================================================================
Public Function CodeSect(ByVal strIndata) 
    
    Dim codehex , i
    Dim tmp1, tmp2

    CodeSect = "-1"
    
    If strIndata = "" Then
        Exit Function
    End If
    
    for i = 1 to len(strIndata)
  codehex = Right("0000" & Hex(Asc(Mid(strIndata,i,1))), 4)
    
  tmp1 = UCase(Left(codehex, 2))
  tmp2 = UCase(Right(codehex, 2))
    
  If (tmp2 >= "A1") And (tmp2 <= "F8") Then
   CodeSect = "0"
   Exit Function
  End If
    Next

End Function

'========================================================================================================= 
Function BpLookUp() 
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
     
 If   LayerShowHide(1) = False Then
             Exit Function 
    End If

     
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "BizPartLookUp"        <%'☜: 비지니스 처리 ASP의 상태 %>
    strVal = strVal & "&txtBp_cd=" & Trim(frm1.txtBp_cd.value)    <%'☜: 조회 조건 데이타 %>

 Call RunMyBizASP(MyBizASP, strVal)          <%'☜: 비지니스 ASP 를 가동 %>
 
End Function


Function CookiePage(Byval Kubun)

 On Error Resume Next
 
 Const CookieSplit = 4877
 
 Dim strTemp, arrVal

 If Kubun = 1 Then

  WriteCookie CookieSplit , frm1.txtConBp_cd.value & parent.gRowSep & frm1.txtConBp_nm.value

 ElseIf Kubun = 0 Then

  strTemp = ReadCookie(CookieSplit)

  If strTemp = "" then Exit Function 

  arrVal = Split(strTemp, parent.gRowSep)

  If arrVal(0) = "" then Exit Function

  frm1.txtConBp_cd.value =  arrVal(0)
  frm1.txtConBp_nm.value =  arrVal(1)

  If Err.number <> 0 Then 
   Err.Clear
   WriteCookie CookieSplit , ""
   Exit Function
  End If

  Call MainQuery()
  
  WriteCookie CookieSplit , ""

 End IF
 
End Function

'========================================================================================================= 
Function JumpChgCheck(strVal)

 Dim IntRetCD

 '************ 싱글인 경우 **************
 If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
    'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
  If IntRetCD = vbNo Then Exit Function
 End If

 Call CookiePage(1)
 Call PgmJump(strVal)

End Function


'========================================================================================
' Function Desc : This function is related to ID Check
'========================================================================================
Function IDCheck(intIDFirst, intIDSecond)
<%
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'
'    주민등록 체크 방법 
'
'    Ex) 680312-1532520
'
'        6,  8,  0,  3,  1,  2,  1,  5,  3,  2,  5,  2
'    x)  2,  3,  4,  5,  6,  7,  8,  9,  2,  3,  4,  5
'    --------------------------------------------------
'    +) 12  24   0  15   6  14   8  45   6   6  20  10  = 166
'
'    11 - ( 166 / 11 ) = 11 - 1 = 10
'    따라서 680312-153252(0)
'    If [11-2=9] Then 680312-153252(9)


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
%>

    Dim arrID(1, 5)      ' 각 주민등록번호 받는 배열 
    Dim seqNum
    Dim logNum1, logNum2
    Dim TotalSum

    logNum1 = 1: logNum2 = 7   ' 주민등록 로직에 필요한 일정순번 초기화....
    
    For seqNum = 0 To 5
        
        logNum1 = logNum1 + 1   ' 생년월일 각 자리를 배열로 선언 
        arrID(0, seqNum) = CInt(Mid(intIDFirst, seqNum + 1, 1)) * logNum1
        
        logNum2 = logNum2 + 1   ' 뒷 7자리중 각 6자리를 배열로 선언 
        arrID(1, seqNum) = CInt(Mid(intIDSecond, seqNum + 1, 1)) * logNum2
        If logNum2 = 9 Then logNum2 = 1  '지우지 말것.... 주민등록 로직에서 필요.... 
    
    Next

    For seqNum = 0 To 5     ' 각 배열로 받은 자리수를 더한다....

        TotalSum = TotalSum + arrID(0, seqNum) + arrID(1, seqNum)

    Next

    IDCheck = 11 - (TotalSum Mod 11) ' 주민등록 맨뒷자리 생성....(가장 중요 로직)

End Function


'========================================================================================================= 
Function Check_ENTP_RGST(ByVal sNumber)


 Check_ENTP_RGST = False
<% 
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'
'    사업자등록번호체크로직 (2002-06-14 sonbumyeol - 새로운 사업자등록번호체크로직)
'
'    거래처별로 사업자 등록번호체크 
'    (해당거래처의 국가 코드가 한국(KR)이 아닐경우 체크하지않음 
'
'    Ex) 603-81-13055
'
'    1. 확인변수 0,3,7,0,3,7,0,3,0.5,0
'    2. 확인변수가 '0'일경우는  더하고, '0'이외일 경우의 숫자는 곱함 
'    3. 확인변수 0.5의 경우는 곱하여 나온수의 정수부와 소수부 를 더함 
'    4. 상기계산으로 합계숫자의 끝자리가 '0'이 되면 정확한 사업자 번호임 
' 
'
'    <사업자 번호 검증예>
'    Ex) 603-81-13055
' 
'        확인변수      
'
'    6  +  0        =  6 
'    0  *  3        =  3
'    3  *  7        =  21 
'    _________________ 
'    8  +  0        =  8
'    1  *  3        =  3
'    _________________
'    1  *  7        =  7
'    3  +  0        =  3
'    0  *  3        =  0
'    5  *  0.5      =  2.5 ( 2+5 =7)
'    5  +  0        =  5
'   _________________________________
'    합계              60     
'
'    --> 합계의 끝자리수가 '0'이므로 정확한 사업자 번호임 
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
%> 
 
 Dim sum,i  
 Dim li_chkvalue(9)  
 Dim NumCnt, Number,NumCnt0,NumCnt1,NumCnt2,NumCnt3,NumCnt4,NumCnt5,NumCnt6,NumCnt7,NumCnt8,NumCnt9


 Number = Replace(sNumber, "-", "")
 
 If isNumeric(Number) = False Then 
    Exit Function
 End If
 
 NumCnt = Len(Number)
 
 Select Case NumCnt
 Case 13
  Exit Function
 
 Case 10
  
  sum = 0
    
  For i = 1 To 10
   li_chkvalue(i-1) = Mid(Number,i,1) 
  Next
          
    
  NumCnt0 = li_chkvalue(0) + 0
  NumCnt1 = li_chkvalue(1) * 3
  NumCnt2 = li_chkvalue(2) * 7
  NumCnt3 = li_chkvalue(3) + 0
  NumCnt4 = li_chkvalue(4) * 3
  NumCnt5 = li_chkvalue(5) * 7
  NumCnt6 = li_chkvalue(6) + 0
  NumCnt7 = li_chkvalue(7) * 3
  NumCnt8 = Int(li_chkvalue(8) * 0.5) + Int(((li_chkvalue(8) * 0.5) * 10) Mod 10)    
  NumCnt9 = li_chkvalue(9) + 0


  sum = (NumCnt0 + NumCnt1 + NumCnt2 + NumCnt3 + NumCnt4 + NumCnt5 + NumCnt6 + NumCnt7 + NumCnt8 + NumCnt9)
    
  if int(sum) MOD 10 <> 0 then Exit Function
    
 Case Else 
 
  Exit Function
 End Select 

 Check_ENTP_RGST = True

End Function

'========================================================================================================= 
Function Check_INDI_RGST(ByVal sID) 

Check_INDI_RGST = False

Dim Weight 
Dim Total 
Dim Chk 
Dim Rmn 
Dim i 
Dim dt 
Dim wt 
Dim Number, Numcnt


Number = Replace(sID, "-", "")
Numcnt = Len(Number)

Select Case Numcnt
 Case 13
  Chk = CDbl(Right(Number, 1))

  Weight = "234567892345"
  Total = 0

  For i = 1 To 12
  dt = CDbl(Mid(Number, i, 1))
  wt = CDbl(Mid(Weight, i, 1))
  Total = Total + (dt * wt)
  Next 

  Rmn = 11 - (Total Mod 11)

  If Rmn > 9 Then Rmn = Rmn Mod 10

  If Rmn <> Chk Then Exit Function

  Case 0
   Check_INDI_RGST = True 
  Exit Function

  Case Else
    
  Exit Function
 End Select 

 Check_INDI_RGST = True
End Function

'========================================================================================
' Function Desc : 숫자만 입력받는 형식 체크 
'========================================================================================
Function NumericCheck()

 Dim objEl, KeyCode
 
 Set objEl = window.event.srcElement
 KeyCode = window.event.keycode

 Select Case KeyCode
    Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
 Case Else
  window.event.keycode = 0
 End Select

End Function


'========================================================================================================= 
Sub Form_Load()

 Call LoadInfTB19029              '⊙: Load table , B_numeric_format
 Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec) 
 Call ggoOper.LockField(Document, "N")                                   '⊙: Lock  Suitable  Field

 Call SetDefaultVal
 Call InitVariables              '⊙: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------

 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄/찾기 
    Call SetToolBar("11101000000011")         '⊙: 버튼 툴바 제어 
 Call CookiePage(0)

End Sub


'========================================================================================================= 
Sub Form_QueryUnload(Cancel , UnloadMode )
 
End Sub

'========================================================================================================= 
Sub txtValidFromDt_Change()
 lgBlnFlgChgValue = True
End Sub

'========================================================================================================= 
Sub txtConValidFromDt_DblClick(Button)
 If Button = 1 Then
  frm1.txtConValidFromDt.Action = 7
  Call SetFocusToDocument("M")   
  Frm1.txtConValidFromDt.Focus
 End If
End Sub
Sub txtValidFromDt_DblClick(Button)
 If Button = 1 Then
  frm1.txtValidFromDt.Action = 7
  Call SetFocusToDocument("M")   
  Frm1.txtValidFromDt.Focus
 End If
End Sub

'========================================================================================================= 
Sub txtConValidFromDt_KeyDown(KeyCode, Shift)
 If KeyCode = 13 Then Call MainQuery()
End Sub
Sub txtValidFromDt_KeyDown(KeyCode, Shift)
 If KeyCode = 13 Then Call MainQuery()
End Sub


'========================================================================================================= 
Sub txtRepre_Rgst1_OnKeyPress()
 Call NumericCheck()
End Sub

Sub txtRepre_Rgst2_OnKeyPress()
 Call NumericCheck()
End Sub

'========================================================================================================= 
Function txtZIP_cd_OnChange()

	If gLookUpEnable = False Then Exit Function

	frm1.txtADDR1.value = ""
	frm1.txtADDR2.value = ""

	If Trim(frm1.txtZIP_cd.value) = "" Then Exit Function
        
'--
    Call CommonQueryRs(" ADDRESS "," B_ZIP_CODE "," COUNTRY_CD =  " & FilterVar(parent.gCountry, "''", "S") & " AND ZIP_CD =  " & FilterVar(frm1.txtZIP_cd.value, "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    if lgf0 = "" then 
		frm1.txtADDR1.value = ""
	else 
	    frm1.txtADDR1.value = Trim(Replace(lgF0,Chr(11),""))
    end if 
'--

End Function


'========================================================================================================= 
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                        <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>

    '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 조회하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")          '⊙: Clear Contents  Field
    Call InitVariables               '⊙: Initializes local global variables
    
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then         <%'⊙: This function check indispensable field%>
       Exit Function
    End If

 Call ggoOper.LockField(Document, "N")                                        <%'⊙: Lock  Suitable  Field%>
 Call SetToolBar("11101000000011")

    '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery                '☜: Query db data
       
    FncQuery = True                '⊙: Processing is OK
        
End Function

'========================================================================================================= 
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '⊙: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x") 
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 신규작업을 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                      '⊙: Clear Condition,Contents Field
    Call ggoOper.LockField(Document, "N")                                       '⊙: Lock  Suitable  Field

 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄/찾기 
    Call SetToolBar("11101000000011")         '⊙: 버튼 툴바 제어 
    Call SetDefaultVal
    Call InitVariables                '⊙: Initializes local global variables
    
    FncNew = True                 '⊙: Processing is OK

End Function


'========================================================================================================= 
Function FncDelete() 
    
    Dim IntRetCD
    
    FncDelete = False              '⊙: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002","x","x","x")
        'Call MsgBox("조회한후에 삭제할 수 있습니다.", vbInformation)
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO,"x","x")
    If IntRetCD = vbNo then exit function
    
	'-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete               '☜: Delete db data
    
    FncDelete = True                                                        <%'⊙: Processing is OK%>
    
End Function


'========================================================================================================= 
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         <%'⊙: Processing is NG%>
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        'Call MsgBox("No data changed!!", vbInformation)
        Exit Function
    End If


	'-----------------------
	    'Check CodeSect 거래처영문명 체크 
    '-----------------------
	 if CodeSect(frm1.txtCust_eng_nm.value ) = "0" then
	'  if msgbox  ("거래처영문명에 한글이 입력됐습니다. 계속하시겠습니까" , vbYesNo) = vbNo then
	  Dim Check_CodeSect
	  Check_CodeSect = DisplayMsgBox("126144", parent.VB_YES_NO, "X", "X")
	   If Check_CodeSect = vbNo Then                                                        <%'⊙: Processing is OK%>
	    Exit Function
	    FncSave = True
	   End If   
	'  end if 
	 end if


	'-----------------------
    'Check Check_ENTP_RGST(사업자 등록번호 체크)
    '-----------------------

	If UCase(parent.gCountry) <> "KR"  Then
	 
	Elseif UCase(parent.gCountry) = "KR"  Then
	  if Check_ENTP_RGST(Trim(frm1.txtBp_Rgst_No.value)) = False then 
	  Dim Check_ENTP
	  Check_ENTP = DisplayMsgBox("126140", parent.VB_YES_NO, "X", "X")
	   If Check_ENTP = vbNo Then                                                        <%'⊙: Processing is OK%>
	    Exit Function
	    FncSave = True
	   End If 
	  End If
	   
	End If
 
	'-----------------------
    'Check Check_INDI_RGST (주민등록번호 체크)
    '-----------------------

	If UCase(parent.gCountry) <> "KR"  Then
	 
	Elseif UCase(parent.gCountry) = "KR"  Then
	  IF Check_INDI_RGST(Trim(frm1.txtRepre_Rgst.value)) = False then
	  Dim Check_INDI
	  Check_INDI = DisplayMsgBox("126139", parent.VB_YES_NO, "X", "X")
	   If Check_INDI = vbNo Then                                                       <%'⊙: Processing is OK%>
	    Exit Function
	    FncSave = True   
	   End If
	  End if 
	End If
    
    '-----------------------
    'Check CodeSect 영문주소 체크 
    '-----------------------
    
	if CodeSect(frm1.txtADDR1_Eng.value ) = "0" Or CodeSect(frm1.txtADDR2_Eng.value ) = "0" Or CodeSect(frm1.txtADDR3_Eng.value ) = "0" then
		
		IntRetCD = DisplayMsgBox("126314", parent.VB_YES_NO,"x","x")
		'IntRetCD = msgbox  ("영문주소에 한글이 입력됐습니다. 저장하시겠습니까?" , vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If	
		
	End if


<%  '-----------------------
    'Check content area
    '-----------------------%>
    If Not chkField(Document, "2") Then                             <%'⊙: Check contents area%>
       Exit Function
    End If

<%  '-----------------------
    'Save function call area
    '-----------------------%>
    CAll DbSave                                                    <%'☜: Save db data%>
    
    FncSave = True                                                          <%'⊙: Processing is OK%>
    
End Function


'========================================================================================================= 
Function FncCopy() 
 Dim IntRetCD

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 계속 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE            <%'⊙: Indicates that current mode is Crate mode%>
    
    <% ' 조건부 필드를 삭제한다. %>
    Call ggoOper.ClearField(Document, "1")                                  <%'⊙: Clear Condition Field%>
    Call ggoOper.LockField(Document, "N")         <%'⊙: This function lock the suitable field%>
    Call InitVariables               <%'⊙: Initializes local global variables%>
 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄/찾기 
    Call SetToolBar("11101000000011")         '⊙: 버튼 툴바 제어 
    
    frm1.txtValidFromDt.Text = ""
	frm1.txtValidFromDt.focus    
    
End Function

'========================================================================================================= 
Function FncPrint() 
 Call Parent.FncPrint()
End Function

'========================================================================================================= 
Function FncExcel() 
 Call Parent.FncExport(parent.C_SINGLE)
End Function

'========================================================================================================= 
Function FncFind() 
    Call Parent.FncFind(parent.C_SINGLE , False)                                     <%'☜:화면 유형, Tab 유무 %>
End Function

'========================================================================================================= 
Function FncExit()
 Dim IntRetCD

 FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("데이타가 변경되었습니다. 종료 하시겠습니까?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True

End Function

'========================================================================================================= 
Function DbDelete() 
    Err.Clear                                                               <%'☜: Protect system from crashing%>

    DbDelete = False              <%'⊙: Processing is NG%>

     
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If
    

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003       <%'☜: 비지니스 처리 ASP의 상태 %>
    strVal = strVal & "&txtBp_cd=" & Trim(frm1.txtBp_cd.value)  <%'☜: 삭제 조건 데이타 %>
    strVal = strVal & "&txtValidFromDt=" & Trim(frm1.txtValidFromDt.Text)
    strVal = strVal & "&txtZIP_cd=" & Trim(frm1.txtZIP_cd.value)
        
	Call RunMyBizASP(MyBizASP, strVal)          <%'☜: 비지니스 ASP 를 가동 %>
 
    DbDelete = True                                                         <%'⊙: Processing is NG%>

End Function

'========================================================================================================= 
Function DbDeleteOk()              <%'☆: 삭제 성공후 실행 로직 %>
 Call FncNew()
End Function

'========================================================================================================= 
Function DbQuery() 
    
    Err.Clear                                                               <%'☜: Protect system from crashing%>
    
    DbQuery = False                                                         <%'⊙: Processing is NG%>

     
 If   LayerShowHide(1) = False Then
             Exit Function 
    End If

     
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       <%'☜: 비지니스 처리 ASP의 상태 %>
    strVal = strVal & "&txtConBp_cd=" & Trim(frm1.txtConBp_cd.value)    <%'☜: 조회 조건 데이타 %>
    strVal = strVal & "&txtConValidFromDt=" & Trim(frm1.txtConValidFromDt.Text)

 Call RunMyBizASP(MyBizASP, strVal)          <%'☜: 비지니스 ASP 를 가동 %>
 
    DbQuery = True                                                          <%'⊙: Processing is NG%>

End Function

'========================================================================================================= 
Function DbQueryOk()              <%'☆: 조회 성공후 실행로직 %>
 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE            <%'⊙: Indicates that current mode is Update mode%>
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")         <%'⊙: This function lock the suitable field%>
 '폴더/조회/입력 
 '/삭제/저장/한줄In
 '/한줄Out/취소/이전 
 '/다음/복사/엑셀 
 '/인쇄/찾기 
	Call SetToolBar("11111000001111")

	frm1.txtConBp_cd.focus

End Function


'========================================================================================================= 
Function DbSave() 

    Err.Clear                <%'☜: Protect system from crashing%>

	DbSave = False               <%'⊙: Processing is NG%>

     
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If

 
    Dim strVal

	With frm1
	 .txtMode.value = parent.UID_M0002           <%'☜: 비지니스 처리 ASP 의 상태 %>
	 .txtFlgMode.value = lgIntFlgMode
	 .txtInsrtUserId.value = parent.gUsrID 
	 .txtUpdtUserId.value = parent.gUsrID
	 '.txtRepre_Rgst.value = Trim(.txtRepre_Rgst1.value) + Trim(.txtRepre_Rgst2.value)
	 .txtContry_cd.value = Trim(parent.gCountry)
	 Call ExecMyBizASP(frm1, BIZ_PGM_ID)          
 
	End With
 
    DbSave = True                                                           <%'⊙: Processing is NG%>
    
End Function

'========================================================================================================= 
Function DbSaveOk()               <%'☆: 저장 성공후 실행 로직 %>

	With frm1
	 .txtConBp_cd.value = .txtBp_cd.value 
	 .txtConBp_nm.value = .txtBp_nm.value
	 .txtConValidFromDt.Text = .txtValidFromDt.Text
	End With
 
    Call InitVariables
    Call MainQuery()

End Function

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" --> 
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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>사업자이력</font></td>
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
                  <TD CLASS=TD5 NOWRAP>거래처</TD>
                  <TD CLASS=TD6><INPUT NAME="txtConBp_cd" ALT="거래처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConBp_cd()">&nbsp;<INPUT NAME="txtConBp_nm" TYPE="Text" SIZE=20 tag="14"></TD>
                  <TD CLASS=TD5 NOWRAP>적용일</TD>
					<TD CLASS=TD6 NOWRAP>
					           <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtConValidFromDt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="적용일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
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
     <TD WIDTH=100% VALIGN=TOP>
      <TABLE <%=LR_SPACE_TYPE_60%>>
       <TR>
                <TD CLASS="TD5" NOWRAP>거래처</TD>
                <TD CLASS="TD6"><INPUT NAME="txtBp_cd" ALT="거래처" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp_cd 0">&nbsp;<INPUT NAME="txtBp_nm" TYPE="Text" SIZE=20 tag="24"></TD>
                <TD CLASS="TD5" NOWRAP>적용일</TD>
        <TD CLASS=TD6 NOWRAP>
                  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtValidFromDt" style="HEIGHT: 20px; WIDTH: 100px" tag="23X1" ALT="적용일" Title="FPDATETIME"></OBJECT>');</SCRIPT>
        </TD>
       </TR>
       <TR>
                <TD CLASS=TD5 NOWRAP>사업자등록번호</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_Rgst_No" ALT="사업자등록번호" TYPE="Text" MAXLENGTH="20" SIZE=35 tag="22"></TD>
                <TD CLASS=TD5 NOWRAP>거래처전명</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust_full_nm" ALT="거래처전명" TYPE="Text" MAXLENGTH="120" SIZE=35 tag="22"></TD>
       </TR>
       <TR>
                <TD CLASS=TD5 NOWRAP>거래처약칭</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust_nm" ALT="거래처약칭" TYPE="Text" MAXLENGTH="50" SIZE=35 tag="22"></TD>
                <TD CLASS=TD5 NOWRAP>거래처영문명</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust_eng_nm" ALT="거래처영문명" TYPE="Text" MAXLENGTH="50" SIZE=35 tag="21" ></TD>
       </TR>
       <TR>
                <TD CLASS=TD5 NOWRAP>대표자명</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepre_nm" ALT="대표자명" TYPE="Text" MAXLENGTH="50" SIZE=35 tag="22"></TD>
                <TD CLASS=TD5 NOWRAP>대표자주민등록번호</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepre_Rgst" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="21"></TD>       
       </TR>
       <TR>
                <TD CLASS=TD5 NOWRAP>업태</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_Class" ALT="업태" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp_cd 1">&nbsp;<INPUT NAME="txtInd_ClassNm" TYPE="Text" SIZE=20 tag="24"></TD>                     
                <TD CLASS=TD5 NOWRAP>업종</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_Type" ALT="업종" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp_cd 2">&nbsp;<INPUT NAME="txtInd_TypeNm" TYPE="Text" SIZE=20 tag="24"></TD>
       </TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>우편번호</TD>			
			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtZIP_cd" TYPE="Text" ALT="우편번호" MAXLENGTH="12" SIZE=20 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnZIP_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenZip" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()"></TD>
			<TD CLASS=TD5 NOWRAP></TD>
			<TD CLASS=TD6 NOWRAP></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>주소</TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR1" TYPE="Text" ALT="주소" MAXLENGTH="100" SIZE=80 tag="24"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP></TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR2" TYPE="Text" ALT="주소" MAXLENGTH="100" SIZE=80 tag="21"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>영문주소</TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR1_Eng" TYPE="Text" ALT="영문주소" MAXLENGTH="50" SIZE=80 tag="21"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP></TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR2_Eng" TYPE="Text" ALT="영문주소" MAXLENGTH="50" SIZE=80 tag="21"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP></TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR3_Eng" TYPE="Text" ALT="영문주소" MAXLENGTH="50" SIZE=80 tag="21"></TD>
		</TR>
       <TR>
                <TD CLASS=TD5 NOWRAP>변경사유</TD>
                <TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtChgCause" TYPE="Text" ALT="변경사유" MAXLENGTH="120" SIZE=80 tag="21"></TD>
       </TR>
       <TR>
            <TD CLASS=TD5 NOWRAP>종사업장번호</TD>
            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSubBizArea" TYPE="Text" MAXLENGTH="4" SIZE=20 tag="21"></TD>       
            <TD CLASS=TD5 NOWRAP>E-Mail</TD>
            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtEMail" TYPE="Text" MAXLENGTH="40" SIZE=20 tag="21"></TD>       
       </TR>
       <TR>
			<TD CLASS=TD5 NOWRAP>종사업장상호및주소</TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtSubBizDesc" TYPE="Text" ALT="종사업장상호및주소" MAXLENGTH="150" SIZE=80 tag="21"></TD>
       </TR>
       <%Call SubFillRemBodyTD5656(6)%>
      </TABLE>
     </TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD <%=HEIGHT_TYPE_01%>></TD>
 </TR>
 <TR HEIGHT=20>
  <TD WIDTH=100%>
   <TABLE <%=LR_SPACE_TYPE_30%>>
       <TR>
     <TD WIDTH=10>&nbsp;</TD>
          <TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_ID1)">거래처등록</a>&nbsp;|&nbsp;<a href = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_ID2)">사업자이력조회</a></TD>
     <TD WIDTH=10>&nbsp;</TD>
    </TR>
   </TABLE>
  </TD>
 </TR>
 <TR >
  <TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0  TABINDEX = -1 ></IFRAME>
  </TD>
 </TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"  TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24" TABINDEX = -1>
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = -1>

<INPUT TYPE=HIDDEN NAME="txtContry_cd" tag="24" TABINDEX = -1>
</FORM>

<DIV ID="MousePT" NAME="MousePT">
 <iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

</BODY>
</HTML> 

