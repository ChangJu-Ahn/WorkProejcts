
<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2106mb1
'*  4. Program Name         : 예산실적조회 
'*  5. Program Desc         : Query of Budget Result
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.02.12
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'*   - 2003/06/12 lgMaxCount 변수값 수정(MA의 C_SHEETMAXROWS_D 삭제)
'=======================================================================================================


%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
	Call LoadBasisGlobalInf() 
    Call LoadInfTB19029B("Q", "A","NOCOOKIE","QB")
    Call LoadBNumericFormatB("Q", "A","NOCOOKIE","QB")
%>	
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3               '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strBdgYymmFr, strBdgYymmTo
Dim strDeptCd
Dim strBdgCdFr, strBdgCdTo
Dim strInternalCd
Dim strWhere
Dim strMsgCd, strMsg1, strMsg2
Dim strColYymm, strDateType


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
Const C_SHEETMAXROWS_D = 100
    Call HideStatusWnd 

    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
    lgMaxCount     = C_SHEETMAXROWS_D
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    
	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))		
	lgInternalCd		= Trim(Request("lgInternalCd"))	
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))	
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))


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
				iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
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

End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "F2106MA101"
    UNISqlId(1) = "F2106MA103"	'부서명 
    UNISqlId(2) = "F2106MA104"	'예산코드명 
    UNISqlId(3) = "F2106MA104"	'예산코드명 

    Redim UNIValue(3,2)		'(Sql, Parameter)

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = UCase(Trim(strWhere))
    UNIValue(1,0) = FilterVar(strDeptCd  , "''", "S") 
    UNIValue(1,1) = FilterVar(Request("OrgChangeId"), "''", "S")   'GetGlobalInf("gChangeOrgId")
    UNIValue(2,0) = FilterVar(strBdgCdFr  , "''", "S") 
    UNIValue(3,0) = FilterVar(strBdgCdTo  , "''", "S") 
    
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
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
				.frm1.txtDeptCd.value = "<%=ConvSPChars(strDeptCd)%>"
				.frm1.txtDeptNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
			End With
		</Script>
<%
    End If

	rs1.Close
	Set rs1 = Nothing

    If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strBdgCdFr <> "" Then 
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBdgCdFr_Alt")
		End If
    Else
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtBdgCdFr.value = "<%=ConvSPChars(strBdgCdFr)%>"
				.frm1.txtBdgNmFr.value = "<%=ConvSPChars(Trim(rs2(0)))%>"
			End With
		</Script>
<%
    End If

	rs2.Close
	Set rs2 = Nothing
    
    If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And strBdgCdTo <> "" Then 
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBdgCdTo_Alt")
		End If
    Else
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtBdgCdTo.value = "<%=ConvSPChars(strBdgCdTo)%>"
				.frm1.txtBdgNmTo.value = "<%=ConvSPChars(Trim(rs3(0)))%>"
			End With
		</Script>
<%
    End If

	rs3.Close
	Set rs3 = Nothing

    If  rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found.
'		rs0.Close
'		Set rs0 = Nothing
'		Set lgADF = Nothing
'		Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strBdgYymmFr = Request("txtBdgYymmFr")
	strBdgYymmTo = Request("txtBdgYymmTo")
	strDeptCd    = Request("txtDeptCd")
	strBdgCdFr   = Request("txtBdgCdFr")
	strBdgCdTo   = Request("txtBdgCdTo")
	strColYymm   = Request("txtColYymm")
	strDateType  = Request("txtDateType")
	
	strWhere = ""
	strWhere = strWhere & " and A.bdg_yyyymm between  " & FilterVar(strBdgYymmFr, "''", "S") & " and  " & FilterVar(strBdgYymmTo, "''", "S") & " "
	If strDeptCd <> "" Then
		strInternalCd = fnGetInternalCd
		strWhere = strWhere & " and A.internal_cd =  " & FilterVar(strInternalCd , "''", "S") & " "
	End If
	If strBdgCdFr <> "" Then strWhere = strWhere & " and A.bdg_cd >=  " & FilterVar(strBdgCdFr , "''", "S") & " "
	If strBdgCdTo <> "" Then strWhere = strWhere & " and A.bdg_cd <=  " & FilterVar(strBdgCdTo , "''", "S") & " "



	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then			
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	


	strWhere	= strWhere	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	


    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
End Sub

'내부부서코드 select
Function fnGetInternalCd()
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,1)

    UNISqlId(0) = "F2106MA103"

    UNIValue(0,0) = FilterVar(strDeptCd  , "''", "S") 
    UNIValue(0,1) = FilterVar(Request("OrgChangeId"), "''", "S") 'GetGlobalInf("gChangeOrgId")
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
	
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        fnGetInternalCd = ""
        rs0.Close
        Set rs0 = Nothing
    Else    
        fnGetInternalCd = rs0(0)
    End If
End Function

%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData  "<%=lgstrData%>"                          '☜: Display data 
         .lgStrPrevKey_A      = "<%=ConvSPChars(lgStrPrevKey)%>"                       '☜: set next data tag
         Call .DbQueryOk
	End with
</Script>	

