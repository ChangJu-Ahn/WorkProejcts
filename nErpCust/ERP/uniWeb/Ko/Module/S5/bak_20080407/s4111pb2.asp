<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111PA1
'*  4. Program Name         : 출하관리번호 팝업 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/19	Date표준적용 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strPoType	                                                           '⊙ : 발주형태 
Dim strPoFrDt	                                                           '⊙ : 발주일 
Dim strPoToDt	                                                           '⊙ :
Dim strSpplCd	                                                           '⊙ : 공급처 
Dim strPurGrpCd	                                                           '⊙ : 구매그룹 
Dim strItemCd	                                                           '⊙ : 품목 
Dim strTrackNo	                                                           '⊙ : Tracking No
Dim BlankchkFlg
'----------------------- 추가된 부분 ----------------------------------------------------------------------
Dim arrRsVal(5)								'* : 화면에 조회해온 Name을 담아놓기 위해 만든 Array
'----------------------------------------------------------------------------------------------------------
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
	Call LoadBasisGlobalInf()
    Call HideStatusWnd 


    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
    lgMaxCount     = 30							                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '날짜 
					iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' 금액 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggAmtOfMoney.DecPoint, 0)
               Case "F3"  '수량 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '단가 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggUnitCost.DecPoint  , 0)
               Case "F5"   '환율 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(rs0(ColCnt))
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '☜: Check if next data exists
        lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Dim arrVal(2)															
    Dim MajorCd
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(3,2)

    UNISqlId(0) = "S4111pa101"									'* : 데이터 조회를 위한 SQL문 

    UNISqlId(1) = "S0000QA002"									'* : 각각의 조회조건부마다 Name 을 가져오는 SQL 문을 만듬 
    UNISqlId(2) = "S0000QA005"
    UNISqlId(3) = "S0000QA000"
 
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strVal = " "

	If Len(Request("txtBpCd")) Then
		strVal = " AND SHIP_TO_PARTY = " & FilterVar(Request("txtBpCd"), "''", "S") & " "		
	Else
		strVal = ""
	End If
	arrVal(0) = FilterVar(Trim(Request("txtBpCd")), "", "S")

	If Len(Request("txtSalesGroup")) Then		
		strVal = strVal & " AND SALES_GRP = " & FilterVar(Request("txtSalesGroup"), "''", "S") & " "		
	End If
	arrVal(1) = FilterVar(Trim(Request("txtSalesGroup")), "", "S")	
	
 	If Len(Request("txtMovType")) Then		
		strVal = strVal & " AND MOV_TYPE = " & FilterVar(Request("txtMovType"), "''", "S") & " "		
		MajorCd = "I0001"
	End If
	arrVal(2) = FilterVar(Trim(Request("txtMovType")), "", "S")		
    
	If Trim(Request("txtRadio")) = "Y" Then
		strVal = strVal & " AND POST_FLAG =" & FilterVar("Y", "''", "S") & " "
	ElseIf Trim(Request("txtRadio")) = "N" Then
		strVal = strVal & " AND POST_FLAG =" & FilterVar("N", "''", "S") & " "
	End If			
		
    If Len(Request("txtPromiseFrDt")) Then
		strVal = strVal & " AND PROMISE_DT >= " & FilterVar(UNIConvDate(Request("txtPromiseFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Request("txtPromiseToDt")) Then
		strVal = strVal & " AND PROMISE_DT <= " & FilterVar(UNIConvDate(Request("txtPromiseToDt")), "''", "S") & ""		
	End If

	strVal = strVal & " AND EXCEPT_DN_FLAG = " & FilterVar("Y", "''", "S") & " "

    UNIValue(0,1) = strVal   '---발주일 
    UNIValue(1,0) = arrVal(0)  
    UNIValue(2,0) = arrVal(1)  
    UNIValue(3,0) = FilterVar(MajorCd, "''", "S")  
    UNIValue(3,1) = arrVal(2)  
   
'================================================   
   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3) '* : Record Set 의 갯수 조정 

    iStr = Split(lgstrRetMsg,gColSep)

    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtBpCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "납품처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
%>
<Script Language=vbscript>
			parent.frm1.txtBpCd.focus
</Script>	       	
<%	       
		End If	
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing

	 	If Len(Request("txtSalesGroup")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True
%>
<Script Language=vbscript>
			parent.frm1.txtSalesGroup.focus
</Script>	       	
<%
		End If	
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If
    
    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing

	 	If Len(Request("txtMovType")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "출하형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       BlankchkFlg = True	
%>
<Script Language=vbscript>
			parent.frm1.txtMovType.focus
</Script>	       	
<%	       
		End If	
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
%>
<Script Language=vbscript>
			parent.frm1.txtBpCd.focus
</Script>	       	
<%
		Else    
		    Call  MakeSpreadSheetData()
		End If
	End If	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub


%>
<Script Language=vbscript>
    With parent
		.frm1.vspdData.Redraw = False
        .ggoSpread.Source    = .frm1.vspdData 
        .ggoSpread.SSShowData "<%=lgstrData%>", "F"                            '☜: Display data 
        .lgStrPrevKey					=  "<%=lgStrPrevKey%>"                       '☜: set next data tag
  		.frm1.txtBpNm.value				=  "<%=ConvSPChars(arrRsVal(1))%>" 	
  		.frm1.txtSalesGroupNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		.frm1.txtMovTypeNm.value		=  "<%=ConvSPChars(arrRsVal(5))%>" 	
		.frm1.vspdData.Redraw = True
        .DbQueryOk
	End with
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
