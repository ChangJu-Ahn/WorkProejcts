<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M3113RB1
'*  4. Program Name         : 반품출고참조 
'*  5. Program Desc         : 반품출고참조 
'*  6. Component List       : 
'*  7. Modified date(First) : 2002/03/21	
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Oh chang won
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
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
  
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
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
	Const C_SHEETMAXROWS_D  = 100            
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr
	Dim PvArr

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
   ReDim PvArr(C_SHEETMAXROWS_D - 1)
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"  '날짜 
					iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  '금액 
                    iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggAmtOfMoney.DecPoint, ggAmtOfMoney.RndPolicy, ggAmtOfMoney.RndUnit, 0)
               Case "F3"  '수량 
                    iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '단가 
                    iStr = iStr & Chr(11) & UniConvNumberDBToCompany(rs0(ColCnt), ggUnitCost.DecPoint, ggUnitCost.RndPolicy, ggUnitCost.RndUnit, 0)
               Case "F5"  '환율 
                    iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case "F2D" '거래금액 
                    iStr = iStr & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(ColCnt), lgCurrency, ggAmtOfMoneyNo,"X","X")
               Case "F4D" '거래단가 
                    iStr = iStr & Chr(11) & UNIConvNumDBToCompanyByCurrency(rs0(ColCnt), lgCurrency, ggUnitCostNo,"X","X")
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(rs0(ColCnt)) 
            End Select
		Next
 
        If  iRCnt < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If

        PvArr(iRCnt) = lgstrData
        lgstrData=""
        rs0.MoveNext
	Loop
    lgstrData = Join(PvArr,"")

    If  iRCnt < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
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
    Dim arrVal(3)															
    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(1,2)

    UNISqlId(0) = "M4132RA101"									'* : 데이터 조회를 위한 SQL문 만듬 
 
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------


	If Trim(Request("txtMvmtNo")) <> "" Then
		strVal = " AND A.MVMT_RCPT_NO >= " & FilterVar(Trim(UCase(Request("txtMvmtNo"))), " " , "S") & "  AND A.MVMT_RCPT_NO <=  " & FilterVar(Trim(UCase(Request("txtMvmtNo"))), " " , "S") & " "
	Else
		strVal = ""
	End If

  	If Len(Trim(Request("txtFrIvDt"))) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT >= " & FilterVar(UNIConvDate(Request("txtFrIvDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToIvDt"))) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT <= " & FilterVar(UNIConvDate(Request("txtToIvDt")), "''", "S") & ""		
	End If

    If Trim(Request("txtSupplierCd")) <> "" Then
        strVal = strVal & " AND A.BP_CD = " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & " "		
    ELSE
        strVal = strVal & " AND A.BP_CD = " & FilterVar("zzzzzzzzz", "''", "S") & ""		
    End If

    If Trim(Request("txtSubcontra2flg")) <> "" Then
		strVal = strVal & " AND G.SUBCONTRA2_FLG = " & FilterVar(Trim(UCase(Request("txtSubcontra2flg"))), "''", "S") & " "			' 외주가공여부 
    End If

    UNIValue(0,1) = strVal 

'================================================================================================================   
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    'UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNIValue(0,2) = " ORDER BY A.MVMT_RCPT_NO DESC "
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	BlankchkFlg = False
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0) '* : Record Set 의 갯수 조정 
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If BlankchkFlg = False Then	
		If  rs0.EOF And rs0.BOF Then
		    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		    rs0.Close
		    Set rs0 = Nothing
    
			' 이 위치에 있는 Response.End 를 삭제하여야 함. Client 단에서 Name을 모두 뿌려준 후에 Response.End 를 기술함.
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
        .ggoSpread.Source    	= .frm1.vspdData 
        .ggoSpread.SSShowData "<%=lgstrData%>"                            	'☜: Display data 
        .lgStrPrevKey		    =  "<%=lgStrPrevKey%>"                      '☜: set next data tag
  		.frm1.txtMvmtNo.value	=  "<%=ConvSPChars(Request("txtMvmtNo"))%>" 	
        .DbQueryOk
	End with
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>

