<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : ADO Template (Save)
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2000/11/01
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : KimTaeHyun
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Expires = -1                                                       '☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True                                                     '☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/IncServer.asp"  -->
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
Dim strFrPaymDt	                                                           
Dim strToPaymDt
Dim strFrPaymNo	                                                           
Dim strToPaymNo
Dim strdeptcd
	                                                           '⊙ : 발주일 
Dim strCond

Dim strMsgCd
Dim strMsg1
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub  MakeSpreadSheetData()
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
Sub  FixUNISQLData()
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(1,2)

    UNISqlId(0) = "A4104RA101"
	UNISqlId(1) = "ADEPTNM"
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,1) = strCond
    
    UNIValue(1,0) = UCase(FilterVar(strDeptCd, "''", "S"))
    UNIValue(1,1) = UCase(FilterVar(gChangeOrgID, "''", "S"))
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub  QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If (rs1.EOF And rs1.BOF) Then
		If strMsgCd = "" And strDeptCd <> "" Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtDeptCd_Alt")
		End If
    Else
%>
		<Script Language=vbScript>
		With parent
			.frm1.txtDeptCd.value = "<%=Trim(rs1(0))%>"
			.frm1.txtDeptNm.value = "<%=Trim(rs1(1))%>"
		End With
		</Script>
<%
    End If
    
	rs1.Close
	Set rs1 = Nothing 
	    
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("990007", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
     strFrPaymDt     = UniConvDate(Trim(Request("txtFrPaymDt")))
     strToPaymDt     = UniConvDate(Trim(Request("txtToPaymDt")))
     strFrPaymNo	 = UCase(Trim(Request("txtFrPaymNo")))                                                          
     strToPaymNo     = UCase(Trim(Request("txtToPaymNo")))
     strDeptCd       = UCase(Trim(Request("txtDeptCd")))
     
     strCond  = " and A.ALLC_TYPE = " & FilterVar("H", "''", "S") & "  "
     
     If strFrPaymDt <> "" Then
		strCond = strCond & " and A.PAYM_DT >= " & FilterVar(strFrPaymDt, "''", "S")
     End If
     
     If strToPaymDt <> "" Then
		strCond = strCond & " and A.PAYM_DT <= " & FilterVar(strToPaymDt, "''", "S")
     End If
     
     If strFrPaymNo <> "" Then
		strCond = strCond & " and A.PAYM_NO >= " & FilterVar(strFrPaymNo, "''", "S")
     End If
     
     If strToPaymNo <> "" Then
		strCond = strCond & " and A.PAYM_NO <= " & FilterVar(strToPaymNo, "''", "S")
     End If
     
     If strdeptcd <> "" Then
		strCond = strCond & " and A.dept_cd = " & FilterVar(strdeptcd, "''", "S")
     End If     
     
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub
'----------------------------------------------------------------------------------------------------------
' Trim string and set string to space if string length is zero
'----------------------------------------------------------------------------------------------------------
'Function FilterVar(Byval str,Byval strALT)
'     Dim strL
'     strL = UCASE(Trim(str))
'     If Len(strL) Then
'        FilterVar = "'" & strL  & "'"
'     Else
'        FilterVar = strALT   
'     End If
'End Function

%>

<Script Language=vbscript>
    With parent
		 IF Trim(.frm1.txtDeptcd.value) = "" Then	
			.frm1.txtDeptNm.Value = ""
		 END IF	
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
         .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                       '☜: set next data tag
         .DbQueryOk
	End with
</Script>	

