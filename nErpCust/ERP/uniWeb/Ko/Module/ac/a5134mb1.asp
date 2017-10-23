<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : a5133mb1
'*  4. Program Name         : 기초계정집계표 조회 
'*  5. Program Desc         : 
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2002.02.01
'*  8. Modified date(Last)  : 2002.02.01
'*  9. Modifier (First)     : Kim, Sang Joong
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================
Response.Expires = -1                                                       '☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True                                                     '☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
																		   '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
'On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim  UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4         '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgPageNo

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strDt, strBizAreaCd, strClassType, strToAcctCd							'사용자정의 변수 
Dim strFrDt, strToDt														'사용자정의 변수 

Dim TDrLocAmt,TCrLocAmt,NDrLocAmt,NCrLocAmt
Dim TTotSumAmt,STotSumAmt,SDrAmt, SCrAmt,DTotSumAmt,CTotSumAmt 
Dim strMsgCd, strMsg1, strMsg2 												'사용자정의 변수 
Dim strWhere0, strWhere1													'사용자정의 변수 
Dim strBizAreaNm															'사용자정의 변수 
Dim strCompYr,strCompMnth,strCompDt											'사용자정의 변수 
Dim strDtYr, strDtMnth, strDtDay											'사용자정의 변수 
Dim strCompFiscStartDt														'사용자정의 변수 
Dim strToGlDts																'사용자정의 변수 


'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  '##
    Call HideStatusWnd 

    lgPageNo   = Request("lgPageNo")                               '☜ : Next key flag
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

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    lgstrData      = ""
    iCnt = 0

    If Len(Trim(lgPageNo)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          Response.Write lgpageno
          iCnt = CInt(lgPageNo)
       End If   
    Else
		lgPageNo	= 0
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    'rs0에 대한 결과 
    rs0.PageSize     = lgMaxCount                                                'Seperate Page with page count (MA : C_SHEETMAXROWS_D )
    rs0.AbsolutePage = lgPageNo + 1

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
               Case "F3"  '수량7
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '단가 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggUnitCost.DecPoint  , 0)
               Case "F5"   '환율 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case Else
                   iStr = iStr & Chr(11) & ConvSPChars(Trim("" & rs0(ColCnt)))
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
'            lgPageNo = CStr(iCnt)
			lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
	Set rs0 = Nothing 
'	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "A5134MA101"	'일계표조회 
	UNISqlId(1) = "A5134MA102"	'현금차변잔액 
	UNISqlId(2) = "A5134MA103"  '현금대변잔액 
	UNISqlId(3) = "A5134MA104"	'기초입력경로    
	UNISqlId(4) = "ABIZNM"		'계정코드    
	
	
	Redim UNIValue(4,4)
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,1) = FilterVar(UCase(gCurrency), "''", "S")
    UNIValue(0,2) = FilterVar(UCase(gCurrency), "''", "S")
    
    UNIValue(0,3) = UCase(Trim(strWhere0))
		
	UNIValue(1,0) = Trim(strWhere0)		
	
	UNIValue(2,0) = Trim(strWhere0)		
	
	
	UNIValue(3,0) = " " & FilterVar(strClassType, "''", "S") & ""
	
	UNIValue(4,0) = " " & FilterVar(strBizAreaCd, "''", "S") & ""
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    'UNIValue(0,2) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)    

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If        	   

   
	If Not(rs1.EOF And rs1.BOF) Then
	
		If IsNull(rs1(0)) = False Then DTotSumAmt    = rs1(0)
	End If
		
	rs1.Close
	Set rs1 = Nothing	

	If Not(rs2.EOF And rs2.BOF) Then
		If IsNull(rs2(0)) = False Then CTotSumAmt    = rs2(0)
	End If
	
	rs2.Close
	Set rs2 = Nothing	
	

	If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If    
        	
    %>
    
    <Script Language=vbscript>
		With parent
    '	.frm1.txtYAmt.value		= "<%=UNINumClientFormat(TTotSumAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtDAmt.value		= "<%=UNINumClientFormat(DTotSumAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtCAmt.value		= "<%=UNINumClientFormat(CTotSumAmt, ggAmtOfMoney.DecPoint, 0)%>"		
		End With
	</script>
	<%
	
	If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And strClassType <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtClassType_Alt")
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtClassType.value = "<%=ConvSPChars(strClassType)%>"
			.txtClassTypeNm.value = "<%=ConvSPChars(Trim(rs3(1)))%>"
			End With
		</Script>
<%			
	End If
	
	rs3.Close
	Set rs3 = Nothing

	If  rs4.EOF And rs4.BOF Then        
		If strMsgCd = "" And strBizAreaCd <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtBizAreaCd_Atl")
		end if
    Else		
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaCd.value = "<%=ConvSPChars(strBizAreaCd)%>"
			.txtBizAreaNm.value = "<%=ConvSPChars(Trim(rs4(1)))%>"
			End With
		</Script>
<%			
    End If  
    
    rs4.Close
    Set rs4 = Nothing
    
	'rs0.Close
	'Set rs0 = Nothing 	
	Set lgADF = Nothing  	
	                                                  '☜: ActiveX Data Factory Object Nothing
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
	
	End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  

    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strDt = Request("txtDateYr")
	strBizAreaCd = UCase(Request("txtBizAreaCd"))
	strClassType = UCase(Request("txtClassType"))	
		
	Call ExtractDateFrom(strCompFiscStartDt,gAPDateFormat,gApDateSeperator,strCompYr,strCompMnth,strCompDt)
	
	strFrDt = strDt +  strCompMnth + strCompDt
			
	strToDt = CStr(CInt(strDt) + 1) + strCompMnTh + strCompDt

	strFrDt = UniConvDateToYYYYMMDD(strFrDt, gServerDateFormat,"-") 
	strToDt = UniDateAdd("D", -1, strToDt, gServerDateFormat)	
	
	
	strWhere0 = ""

	strWhere0 = strWhere0 & " B.gl_dt between  " & FilterVar(strFrDt, "''", "S") & " and  " & FilterVar(strToDt, "''", "S") & " "

		
	If strBizAreaCd <> "" Then		
		strWhere0 = strWhere0 & " AND B.biz_area_cd =  " & FilterVar(strBizAreaCd , "''", "S") & " " 		
	End If

	If strClassType <> "" Then		
		strWhere0 = strWhere0 & " AND B.gl_input_type =  " & FilterVar(strClassType , "''", "S") & " " 
	End If

	

End Sub

%>


<Script Language=vbscript>
	With parent
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
'		.lgPageNo =  "<%=ConvSPChars(lgPageNo)%>"                       '☜: set next data tag
		.lgPageNo = "<%=lgPageNo%>"
		.DbQueryOk
	End with
	
</Script>	

<%
	Response.End 
%>
