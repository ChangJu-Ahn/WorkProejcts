<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%
'********************************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        :                   
'*  3. Program ID           : B1263MA1                    
'*  4. Program Name         : ������̷µ��                  
'*  5. Program Desc         : ������̷µ��               
'*  6. Comproxy List        : PB5CS41.dll, PB5CS44.dll, PB5CS45.dll
'*  7. Modified date(First) : 2001/01/05                
'*  8. Modified date(Last)  : 2001/01/05                
'*  9. Modifier (First)     : Kim Hyungsuk                
'* 10. Modifier (Last)      : Sonbumyeol  
'* 11. Comment              :               
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"									
'*                            this mark(��) Means that "may  change"									
'*                            this mark(��) Means that "must change"									
'* 13. History              : 2002/12/02 : INCLUDE �߰� ���� ����, Kang Jun Gu
'*                            2002/12/09 : INCLUDE �ٽ� ���� ����, Kang Jun Gu
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
Option Explicit                                                             '��: indicates that All variables must be declared in advance

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
<% '== ��� == %>
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

 arrParam(0) = "�ŷ�ó"			<%' �˾� ��Ī %>
 arrParam(1) = "B_BIZ_PARTNER"      <%' TABLE ��Ī %>
 arrParam(2) = Trim(frm1.txtConBp_cd.Value)		<%' Code Condition%>
 arrParam(3) = ""								<%' Name Cindition%>
 arrParam(4) = ""								<%' Where Condition%>
 arrParam(5) = "�ŷ�ó"						<%' TextBox ��Ī %>
 
 arrField(0) = "BP_CD"        <%' Field��(0)%>
 arrField(1) = "BP_NM"        <%' Field��(1)%>
    
 arrHeader(0) = "�ŷ�ó"				<%' Header��(0)%>
 arrHeader(1) = "�ŷ�ó��Ī"           <%' Header��(1)%>
    
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
	arrParam(0) = "�ŷ�ó"			<%' �˾� ��Ī %>
	arrParam(1) = "B_BIZ_PARTNER"					<%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtBp_cd.Value)			<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = ""								<%' Where Condition%>
	arrParam(5) = "�ŷ�ó"					<%' TextBox ��Ī %>
	 
	arrField(0) = "BP_CD"        <%' Field��(0)%>
	arrField(1) = "BP_NM"        <%' Field��(1)%>
	    
	arrHeader(0) = "�ŷ�ó"				<%' Header��(0)%>
	arrHeader(1) = "�ŷ�ó��Ī"           <%' Header��(1)%>
	
	frm1.txtBp_cd.focus
	
 Case 1            <%' ���� %>
	arrParam(0) = "����"      <%' �˾� ��Ī %>
	arrParam(1) = "B_MINOR"       <%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtInd_Class.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("B9003", "''", "S") & " "    <%' Where Condition%>
	arrParam(5) = "����"            <%' TextBox ��Ī %>
  
	arrField(0) = "MINOR_CD"      <%' Field��(0)%>
	arrField(1) = "MINOR_NM"      <%' Field��(1)%>
		     
	arrHeader(0) = "����"			<%' Header��(0)%>
	arrHeader(1) = "���¸�"           <%' Header��(1)%>

	frm1.txtInd_Class.focus
 Case 2            <%' ���� %>
	arrParam(0) = "����"      <%' �˾� ��Ī %>
	arrParam(1) = "B_MINOR"       <%' TABLE ��Ī %>
	arrParam(2) = Trim(frm1.txtInd_Type.value)		<%' Code Condition%>
	arrParam(3) = ""								<%' Name Cindition%>
	arrParam(4) = "MAJOR_CD=" & FilterVar("B9002", "''", "S") & " "    <%' Where Condition%>
	arrParam(5) = "����"            <%' TextBox ��Ī %>
		  
	arrField(0) = "MINOR_CD"      <%' Field��(0)%>
	arrField(1) = "MINOR_NM"      <%' Field��(1)%>
		     
	arrHeader(0) = "����"			<%' Header��(0)%>
	arrHeader(1) = "������"           <%' Header��(1)%>

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
' Function Desc : �ѱ��� �����Ѵ�.
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
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    
     
 If   LayerShowHide(1) = False Then
             Exit Function 
    End If

     
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "BizPartLookUp"        <%'��: �����Ͻ� ó�� ASP�� ���� %>
    strVal = strVal & "&txtBp_cd=" & Trim(frm1.txtBp_cd.value)    <%'��: ��ȸ ���� ����Ÿ %>

 Call RunMyBizASP(MyBizASP, strVal)          <%'��: �����Ͻ� ASP �� ���� %>
 
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

 '************ �̱��� ��� **************
 If lgBlnFlgChgValue = True Then
  IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
    'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)
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
'    �ֹε�� üũ ��� 
'
'    Ex) 680312-1532520
'
'        6,  8,  0,  3,  1,  2,  1,  5,  3,  2,  5,  2
'    x)  2,  3,  4,  5,  6,  7,  8,  9,  2,  3,  4,  5
'    --------------------------------------------------
'    +) 12  24   0  15   6  14   8  45   6   6  20  10  = 166
'
'    11 - ( 166 / 11 ) = 11 - 1 = 10
'    ���� 680312-153252(0)
'    If [11-2=9] Then 680312-153252(9)


'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
%>

    Dim arrID(1, 5)      ' �� �ֹε�Ϲ�ȣ �޴� �迭 
    Dim seqNum
    Dim logNum1, logNum2
    Dim TotalSum

    logNum1 = 1: logNum2 = 7   ' �ֹε�� ������ �ʿ��� �������� �ʱ�ȭ....
    
    For seqNum = 0 To 5
        
        logNum1 = logNum1 + 1   ' ������� �� �ڸ��� �迭�� ���� 
        arrID(0, seqNum) = CInt(Mid(intIDFirst, seqNum + 1, 1)) * logNum1
        
        logNum2 = logNum2 + 1   ' �� 7�ڸ��� �� 6�ڸ��� �迭�� ���� 
        arrID(1, seqNum) = CInt(Mid(intIDSecond, seqNum + 1, 1)) * logNum2
        If logNum2 = 9 Then logNum2 = 1  '������ ����.... �ֹε�� �������� �ʿ�.... 
    
    Next

    For seqNum = 0 To 5     ' �� �迭�� ���� �ڸ����� ���Ѵ�....

        TotalSum = TotalSum + arrID(0, seqNum) + arrID(1, seqNum)

    Next

    IDCheck = 11 - (TotalSum Mod 11) ' �ֹε�� �ǵ��ڸ� ����....(���� �߿� ����)

End Function


'========================================================================================================= 
Function Check_ENTP_RGST(ByVal sNumber)


 Check_ENTP_RGST = False
<% 
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'
'    ����ڵ�Ϲ�ȣüũ���� (2002-06-14 sonbumyeol - ���ο� ����ڵ�Ϲ�ȣüũ����)
'
'    �ŷ�ó���� ����� ��Ϲ�ȣüũ 
'    (�ش�ŷ�ó�� ���� �ڵ尡 �ѱ�(KR)�� �ƴҰ�� üũ�������� 
'
'    Ex) 603-81-13055
'
'    1. Ȯ�κ��� 0,3,7,0,3,7,0,3,0.5,0
'    2. Ȯ�κ����� '0'�ϰ���  ���ϰ�, '0'�̿��� ����� ���ڴ� ���� 
'    3. Ȯ�κ��� 0.5�� ���� ���Ͽ� ���¼��� �����ο� �Ҽ��� �� ���� 
'    4. ��������� �հ������ ���ڸ��� '0'�� �Ǹ� ��Ȯ�� ����� ��ȣ�� 
' 
'
'    <����� ��ȣ ������>
'    Ex) 603-81-13055
' 
'        Ȯ�κ���      
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
'    �հ�              60     
'
'    --> �հ��� ���ڸ����� '0'�̹Ƿ� ��Ȯ�� ����� ��ȣ�� 
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
' Function Desc : ���ڸ� �Է¹޴� ���� üũ 
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

 Call LoadInfTB19029              '��: Load table , B_numeric_format
 Call ggoOper.FormatField(Document, "A", ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec) 
 Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field

 Call SetDefaultVal
 Call InitVariables              '��: Initializes local global variables
    '----------  Coding part  -------------------------------------------------------------

 '����/��ȸ/�Է� 
 '/����/����/����In
 '/����Out/���/���� 
 '/����/����/���� 
 '/�μ�/ã�� 
    Call SetToolBar("11101000000011")         '��: ��ư ���� ���� 
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
    
    FncQuery = False                                                        <%'��: Processing is NG%>
    
    Err.Clear                                                               <%'��: Protect system from crashing%>

    '-----------------------
    'Check previous data area
    '----------------------- %>
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
	'-----------------------
    'Erase contents area
    '----------------------- 
    Call ggoOper.ClearField(Document, "2")          '��: Clear Contents  Field
    Call InitVariables               '��: Initializes local global variables
    
	'-----------------------
    'Check condition area
    '----------------------- 
    If Not chkField(Document, "1") Then         <%'��: This function check indispensable field%>
       Exit Function
    End If

 Call ggoOper.LockField(Document, "N")                                        <%'��: Lock  Suitable  Field%>
 Call SetToolBar("11101000000011")

    '-----------------------
    'Query function call area
    '----------------------- 
    Call DbQuery                '��: Query db data
       
    FncQuery = True                '��: Processing is OK
        
End Function

'========================================================================================================= 
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                          '��: Processing is NG
    
  '-----------------------
    'Check previous data area
    '-----------------------
    If lgBlnFlgChgValue = True Then
        IntRetCD = DisplayMsgBox("900015", parent.VB_YES_NO,"x","x") 
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. �ű��۾��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
		    Exit Function
		End If
    End If
    
    '-----------------------
    'Erase condition area
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "A")                                      '��: Clear Condition,Contents Field
    Call ggoOper.LockField(Document, "N")                                       '��: Lock  Suitable  Field

 '����/��ȸ/�Է� 
 '/����/����/����In
 '/����Out/���/���� 
 '/����/����/���� 
 '/�μ�/ã�� 
    Call SetToolBar("11101000000011")         '��: ��ư ���� ���� 
    Call SetDefaultVal
    Call InitVariables                '��: Initializes local global variables
    
    FncNew = True                 '��: Processing is OK

End Function


'========================================================================================================= 
Function FncDelete() 
    
    Dim IntRetCD
    
    FncDelete = False              '��: Processing is NG
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgIntFlgMode <> parent.OPMD_UMODE Then      
        Call DisplayMsgBox("900002","x","x","x")
        'Call MsgBox("��ȸ���Ŀ� ������ �� �ֽ��ϴ�.", vbInformation)
        Exit Function
    End If
    
    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO,"x","x")
    If IntRetCD = vbNo then exit function
    
	'-----------------------
    'Delete function call area
    '-----------------------
    Call DbDelete               '��: Delete db data
    
    FncDelete = True                                                        <%'��: Processing is OK%>
    
End Function


'========================================================================================================= 
Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                                         <%'��: Processing is NG%>
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    
    '-----------------------
    'Precheck area
    '-----------------------
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001","x","x","x")
        'Call MsgBox("No data changed!!", vbInformation)
        Exit Function
    End If


	'-----------------------
	    'Check CodeSect �ŷ�ó������ üũ 
    '-----------------------
	 if CodeSect(frm1.txtCust_eng_nm.value ) = "0" then
	'  if msgbox  ("�ŷ�ó������ �ѱ��� �Էµƽ��ϴ�. ����Ͻðڽ��ϱ�" , vbYesNo) = vbNo then
	  Dim Check_CodeSect
	  Check_CodeSect = DisplayMsgBox("126144", parent.VB_YES_NO, "X", "X")
	   If Check_CodeSect = vbNo Then                                                        <%'��: Processing is OK%>
	    Exit Function
	    FncSave = True
	   End If   
	'  end if 
	 end if


	'-----------------------
    'Check Check_ENTP_RGST(����� ��Ϲ�ȣ üũ)
    '-----------------------

	If UCase(parent.gCountry) <> "KR"  Then
	 
	Elseif UCase(parent.gCountry) = "KR"  Then
	  if Check_ENTP_RGST(Trim(frm1.txtBp_Rgst_No.value)) = False then 
	  Dim Check_ENTP
	  Check_ENTP = DisplayMsgBox("126140", parent.VB_YES_NO, "X", "X")
	   If Check_ENTP = vbNo Then                                                        <%'��: Processing is OK%>
	    Exit Function
	    FncSave = True
	   End If 
	  End If
	   
	End If
 
	'-----------------------
    'Check Check_INDI_RGST (�ֹε�Ϲ�ȣ üũ)
    '-----------------------

	If UCase(parent.gCountry) <> "KR"  Then
	 
	Elseif UCase(parent.gCountry) = "KR"  Then
	  IF Check_INDI_RGST(Trim(frm1.txtRepre_Rgst.value)) = False then
	  Dim Check_INDI
	  Check_INDI = DisplayMsgBox("126139", parent.VB_YES_NO, "X", "X")
	   If Check_INDI = vbNo Then                                                       <%'��: Processing is OK%>
	    Exit Function
	    FncSave = True   
	   End If
	  End if 
	End If
    
    '-----------------------
    'Check CodeSect �����ּ� üũ 
    '-----------------------
    
	if CodeSect(frm1.txtADDR1_Eng.value ) = "0" Or CodeSect(frm1.txtADDR2_Eng.value ) = "0" Or CodeSect(frm1.txtADDR3_Eng.value ) = "0" then
		
		IntRetCD = DisplayMsgBox("126314", parent.VB_YES_NO,"x","x")
		'IntRetCD = msgbox  ("�����ּҿ� �ѱ��� �Էµƽ��ϴ�. �����Ͻðڽ��ϱ�?" , vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If	
		
	End if


<%  '-----------------------
    'Check content area
    '-----------------------%>
    If Not chkField(Document, "2") Then                             <%'��: Check contents area%>
       Exit Function
    End If

<%  '-----------------------
    'Save function call area
    '-----------------------%>
    CAll DbSave                                                    <%'��: Save db data%>
    
    FncSave = True                                                          <%'��: Processing is OK%>
    
End Function


'========================================================================================================= 
Function FncCopy() 
 Dim IntRetCD

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", parent.VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ��� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = parent.OPMD_CMODE            <%'��: Indicates that current mode is Crate mode%>
    
    <% ' ���Ǻ� �ʵ带 �����Ѵ�. %>
    Call ggoOper.ClearField(Document, "1")                                  <%'��: Clear Condition Field%>
    Call ggoOper.LockField(Document, "N")         <%'��: This function lock the suitable field%>
    Call InitVariables               <%'��: Initializes local global variables%>
 '����/��ȸ/�Է� 
 '/����/����/����In
 '/����Out/���/���� 
 '/����/����/���� 
 '/�μ�/ã�� 
    Call SetToolBar("11101000000011")         '��: ��ư ���� ���� 
    
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
    Call Parent.FncFind(parent.C_SINGLE , False)                                     <%'��:ȭ�� ����, Tab ���� %>
End Function

'========================================================================================================= 
Function FncExit()
 Dim IntRetCD

 FncExit = False

    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", parent.VB_YES_NO,"x","x")
		'IntRetCD = MsgBox("����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?", vbYesNo)
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    FncExit = True

End Function

'========================================================================================================= 
Function DbDelete() 
    Err.Clear                                                               <%'��: Protect system from crashing%>

    DbDelete = False              <%'��: Processing is NG%>

     
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If
    

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0003       <%'��: �����Ͻ� ó�� ASP�� ���� %>
    strVal = strVal & "&txtBp_cd=" & Trim(frm1.txtBp_cd.value)  <%'��: ���� ���� ����Ÿ %>
    strVal = strVal & "&txtValidFromDt=" & Trim(frm1.txtValidFromDt.Text)
    strVal = strVal & "&txtZIP_cd=" & Trim(frm1.txtZIP_cd.value)
        
	Call RunMyBizASP(MyBizASP, strVal)          <%'��: �����Ͻ� ASP �� ���� %>
 
    DbDelete = True                                                         <%'��: Processing is NG%>

End Function

'========================================================================================================= 
Function DbDeleteOk()              <%'��: ���� ������ ���� ���� %>
 Call FncNew()
End Function

'========================================================================================================= 
Function DbQuery() 
    
    Err.Clear                                                               <%'��: Protect system from crashing%>
    
    DbQuery = False                                                         <%'��: Processing is NG%>

     
 If   LayerShowHide(1) = False Then
             Exit Function 
    End If

     
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & parent.UID_M0001       <%'��: �����Ͻ� ó�� ASP�� ���� %>
    strVal = strVal & "&txtConBp_cd=" & Trim(frm1.txtConBp_cd.value)    <%'��: ��ȸ ���� ����Ÿ %>
    strVal = strVal & "&txtConValidFromDt=" & Trim(frm1.txtConValidFromDt.Text)

 Call RunMyBizASP(MyBizASP, strVal)          <%'��: �����Ͻ� ASP �� ���� %>
 
    DbQuery = True                                                          <%'��: Processing is NG%>

End Function

'========================================================================================================= 
Function DbQueryOk()              <%'��: ��ȸ ������ ������� %>
 
    '-----------------------
    'Reset variables area
    '-----------------------
    lgIntFlgMode = parent.OPMD_UMODE            <%'��: Indicates that current mode is Update mode%>
    lgBlnFlgChgValue = False
    
    Call ggoOper.LockField(Document, "Q")         <%'��: This function lock the suitable field%>
 '����/��ȸ/�Է� 
 '/����/����/����In
 '/����Out/���/���� 
 '/����/����/���� 
 '/�μ�/ã�� 
	Call SetToolBar("11111000001111")

	frm1.txtConBp_cd.focus

End Function


'========================================================================================================= 
Function DbSave() 

    Err.Clear                <%'��: Protect system from crashing%>

	DbSave = False               <%'��: Processing is NG%>

     
	If   LayerShowHide(1) = False Then
             Exit Function 
    End If

 
    Dim strVal

	With frm1
	 .txtMode.value = parent.UID_M0002           <%'��: �����Ͻ� ó�� ASP �� ���� %>
	 .txtFlgMode.value = lgIntFlgMode
	 .txtInsrtUserId.value = parent.gUsrID 
	 .txtUpdtUserId.value = parent.gUsrID
	 '.txtRepre_Rgst.value = Trim(.txtRepre_Rgst1.value) + Trim(.txtRepre_Rgst2.value)
	 .txtContry_cd.value = Trim(parent.gCountry)
	 Call ExecMyBizASP(frm1, BIZ_PGM_ID)          
 
	End With
 
    DbSave = True                                                           <%'��: Processing is NG%>
    
End Function

'========================================================================================================= 
Function DbSaveOk()               <%'��: ���� ������ ���� ���� %>

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
        <td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>������̷�</font></td>
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
                  <TD CLASS=TD5 NOWRAP>�ŷ�ó</TD>
                  <TD CLASS=TD6><INPUT NAME="txtConBp_cd" ALT="�ŷ�ó" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenConBp_cd()">&nbsp;<INPUT NAME="txtConBp_nm" TYPE="Text" SIZE=20 tag="14"></TD>
                  <TD CLASS=TD5 NOWRAP>������</TD>
					<TD CLASS=TD6 NOWRAP>
					           <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtConValidFromDt" CLASS=FPDTYYYYMMDD tag="12X1" ALT="������" Title="FPDATETIME"></OBJECT>');</SCRIPT>
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
                <TD CLASS="TD5" NOWRAP>�ŷ�ó</TD>
                <TD CLASS="TD6"><INPUT NAME="txtBp_cd" ALT="�ŷ�ó" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="23XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp_cd 0">&nbsp;<INPUT NAME="txtBp_nm" TYPE="Text" SIZE=20 tag="24"></TD>
                <TD CLASS="TD5" NOWRAP>������</TD>
        <TD CLASS=TD6 NOWRAP>
                  <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 NAME="txtValidFromDt" style="HEIGHT: 20px; WIDTH: 100px" tag="23X1" ALT="������" Title="FPDATETIME"></OBJECT>');</SCRIPT>
        </TD>
       </TR>
       <TR>
                <TD CLASS=TD5 NOWRAP>����ڵ�Ϲ�ȣ</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtBp_Rgst_No" ALT="����ڵ�Ϲ�ȣ" TYPE="Text" MAXLENGTH="20" SIZE=35 tag="22"></TD>
                <TD CLASS=TD5 NOWRAP>�ŷ�ó����</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust_full_nm" ALT="�ŷ�ó����" TYPE="Text" MAXLENGTH="120" SIZE=35 tag="22"></TD>
       </TR>
       <TR>
                <TD CLASS=TD5 NOWRAP>�ŷ�ó��Ī</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust_nm" ALT="�ŷ�ó��Ī" TYPE="Text" MAXLENGTH="50" SIZE=35 tag="22"></TD>
                <TD CLASS=TD5 NOWRAP>�ŷ�ó������</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtCust_eng_nm" ALT="�ŷ�ó������" TYPE="Text" MAXLENGTH="50" SIZE=35 tag="21" ></TD>
       </TR>
       <TR>
                <TD CLASS=TD5 NOWRAP>��ǥ�ڸ�</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepre_nm" ALT="��ǥ�ڸ�" TYPE="Text" MAXLENGTH="50" SIZE=35 tag="22"></TD>
                <TD CLASS=TD5 NOWRAP>��ǥ���ֹε�Ϲ�ȣ</TD>
        <TD CLASS=TD6 NOWRAP><INPUT NAME="txtRepre_Rgst" TYPE="Text" MAXLENGTH="20" SIZE=20 tag="21"></TD>       
       </TR>
       <TR>
                <TD CLASS=TD5 NOWRAP>����</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_Class" ALT="����" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp_cd 1">&nbsp;<INPUT NAME="txtInd_ClassNm" TYPE="Text" SIZE=20 tag="24"></TD>                     
                <TD CLASS=TD5 NOWRAP>����</TD>
                <TD CLASS=TD6 NOWRAP><INPUT NAME="txtInd_Type" ALT="����" TYPE="Text" MAXLENGTH="5" SIZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnBp_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenBp_cd 2">&nbsp;<INPUT NAME="txtInd_TypeNm" TYPE="Text" SIZE=20 tag="24"></TD>
       </TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>�����ȣ</TD>			
			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtZIP_cd" TYPE="Text" ALT="�����ȣ" MAXLENGTH="12" SIZE=20 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnZIP_Cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenZip" OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()"></TD>
			<TD CLASS=TD5 NOWRAP></TD>
			<TD CLASS=TD6 NOWRAP></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>�ּ�</TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR1" TYPE="Text" ALT="�ּ�" MAXLENGTH="100" SIZE=80 tag="24"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP></TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR2" TYPE="Text" ALT="�ּ�" MAXLENGTH="100" SIZE=80 tag="21"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP>�����ּ�</TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR1_Eng" TYPE="Text" ALT="�����ּ�" MAXLENGTH="50" SIZE=80 tag="21"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP></TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR2_Eng" TYPE="Text" ALT="�����ּ�" MAXLENGTH="50" SIZE=80 tag="21"></TD>
		</TR>
		<TR>
			<TD CLASS=TD5 NOWRAP></TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtADDR3_Eng" TYPE="Text" ALT="�����ּ�" MAXLENGTH="50" SIZE=80 tag="21"></TD>
		</TR>
       <TR>
                <TD CLASS=TD5 NOWRAP>�������</TD>
                <TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtChgCause" TYPE="Text" ALT="�������" MAXLENGTH="120" SIZE=80 tag="21"></TD>
       </TR>
       <TR>
            <TD CLASS=TD5 NOWRAP>��������ȣ</TD>
            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtSubBizArea" TYPE="Text" MAXLENGTH="4" SIZE=20 tag="21"></TD>       
            <TD CLASS=TD5 NOWRAP>E-Mail</TD>
            <TD CLASS=TD6 NOWRAP><INPUT NAME="txtEMail" TYPE="Text" MAXLENGTH="40" SIZE=20 tag="21"></TD>       
       </TR>
       <TR>
			<TD CLASS=TD5 NOWRAP>��������ȣ���ּ�</TD>
			<TD CLASS=TD6 NOWRAP COLSPAN=3><INPUT NAME="txtSubBizDesc" TYPE="Text" ALT="��������ȣ���ּ�" MAXLENGTH="150" SIZE=80 tag="21"></TD>
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
          <TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_ID1)">�ŷ�ó���</a>&nbsp;|&nbsp;<a href = "VBSCRIPT:JumpChgCheck(BIZ_PGM_JUMP_ID2)">������̷���ȸ</a></TD>
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

