<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>


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

%>

<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
'On Error Resume Next
'Err.Clear

Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")
Call HideStatusWnd 

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
	                                                           '⊙ : 발주일 
Dim strCond

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    Const C_SHEETMAXROWS_D  = 30 
   
    lgMaxCount = CInt(C_SHEETMAXROWS_D)                     '☜: Max fetched data at a time


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
    
'    strDeptNm = rs0(1)
    
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
                    iStr = iStr & Chr(11) & rs0(ColCnt) 
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
	Dim strDiffer
	
	strDiffer = Trim(Request("txtdiffer"))
	
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	If Trim(strDiffer) = "1"  Then
		UNISqlId(0) = "f3101RA101"
	Else 
		UNISqlId(0) = "f3101RA201"
	End If 

	Redim UNIValue(0,2)


    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    
    'UNIValue(0,2) = UCase(Trim(strtotempgldt))
    

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
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
Sub TrimData()
	Dim strCd
	Dim strNm
	Dim strDiffer
 
    strCd     = Trim(Request("txtcd"))
    strNm     = Trim(Request("txtNm"))
    strDiffer = Trim(Request("txtdiffer"))
     
	' 권한관리 추가 
	lgAuthBizAreaCd	= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd	= Trim(Request("lgInternalCd"))
	lgSubInternalCd	= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))     

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		lgBizAreaAuthSQL		= " AND E.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		lgInternalCdAuthSQL		= " AND E.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		lgSubInternalCdAuthSQL	= " AND E.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		lgAuthUsrIDAuthSQL		= " AND E.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If

    If strDiffer = "3" Then 
		If strCd = "" and strNm <> "" Then
			strCond = " and A.BANK_CD >=  " & FilterVar(strNm , "''", "S") & ""
		ElseIf strCd <> "" and strNm = ""  Then      
			strCond =  " and B.BANK_ACCT_NO >=  " & FilterVar(strCd , "''", "S") & ""
		ElseIf strCd = "" and strNm = "" Then
			strCond =  " and B.BANK_ACCT_NO >=  " & FilterVar(strCd , "''", "S") & ""
		Elseif strCd <> "" and strNm <> "" Then
			strCond =  " and B.BANK_ACCT_NO >=  " & FilterVar(strCd , "''", "S") & ""
		End if

	    '2008.04.25 자계좌만 나타나게 수정	>>air 
	 	'strCond = strCond & " and (ISNULL(B.BANK_ACCT_PRNT,'N') = 'N' OR B.BANK_ACCT_PRNT = '') "
		''strCond = strCond & " and ISNULL(A.PAR_BANK_CD,'') <> '' "
		
		' 권한관리 추가 
		strCond	= strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
    Elseif strDiffer = "2" Then
		If strCd = "" and strNm <> "" Then
			strCond = " and A.BANK_NM >=  " & FilterVar(strNm , "''", "S") & ""
			'-----------------------------------------------------------<<2004.04.14>>
			lgTailList = " Order By A.BANK_NM ASC "                   'Bank_nm 이 조건으로 있을 경우만 
			'--------------------------
		Elseif strCd <> "" and strNm = "" then
			strCond =  " and A.BANK_CD >=  " & FilterVar(strCd , "''", "S") & ""
		Elseif strCd = "" and strNm = "" Then
			strCond =  " and A.BANK_CD >=  " & FilterVar(strCd , "''", "S") & ""
		Elseif strCd <> "" and strNm <> "" Then
			strCond =  " and A.BANK_CD >=  " & FilterVar(strCd , "''", "S") & ""
		End if
		
		'2008.04.25 자계좌만 나타나게 수정	>>air 
	 	'strCond = strCond & " and (ISNULL(B.BANK_ACCT_PRNT,'N') = 'N' OR B.BANK_ACCT_PRNT = '') "		
	 	''strCond = strCond & " and ISNULL(A.PAR_BANK_CD,'') <> '' "
		
		' 권한관리 추가 
		strCond	= strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL
	Else
		If strCd = "" and strNm <> "" Then
			strCond = " and A.BANK_NM >=  " & FilterVar(strNm , "''", "S") & ""
			'-----------------------------------------------------------<<2004.04.14>>
			lgTailList = " Order By A.BANK_NM ASC "                   'Bank_nm 이 조건으로 있을 경우만 
			'--------------------------
		Elseif strCd <> "" and strNm = "" then
			strCond =  " and A.BANK_CD >=  " & FilterVar(strCd , "''", "S") & ""
		Elseif strCd = "" and strNm = "" Then
			strCond =  " and A.BANK_CD >=  " & FilterVar(strCd , "''", "S") & ""
		Elseif strCd <> "" and strNm <> "" Then
			strCond =  " and A.BANK_CD >=  " & FilterVar(strCd , "''", "S") & ""
		End if	
		
		'2008.04.25 자계좌만 나타나게 수정	>>air 
	 	'strCond = strCond & " and (ISNULL(B.BANK_ACCT_PRNT,'N') = 'N' OR B.BANK_ACCT_PRNT = '') "
	 	''strCond = strCond & " and ISNULL(A.PAR_BANK_CD,'') <> '' "
	 				
    End If
 	
End Sub


%>

<Script Language=vbscript>
    With parent
	 
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"                            '☜: Display data 
         .lgStrPrevKey        =  "<%=ConvSPChars(lgStrPrevKey)%>"                       '☜: set next data tag
         .DbQueryOk
	End with
</Script>
