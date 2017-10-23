<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2109pa1
'*  4. Program Name         : 예산상세내역보기 
'*  5. Program Desc         : Popup of Budget Detail
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.04.01
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : 
'* 11. Comment              :
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
Dim strCond
Dim strBdgYymm, strDeptCd, strBdgCd
Dim strColYymm, strDateType
Dim strMsgCd, strMsg1, strMsg2
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
					If ColCnt = CInt(strColYymm) Then		'년월 Mask
						iStr = iStr & Chr(11) & Trim(Left(rs0(ColCnt),4) & strDateType & Right(rs0(ColCnt),2))
					Else
						iStr = iStr & Chr(11) & ConvSPChars(Trim("" & rs0(ColCnt)))
					End If
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
  	
'	rs0.Close
'	Set rs0 = Nothing 
'	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(1,13)

    UNISqlId(0) = "F2109PA101"
    UNISqlId(1) = "F2109PA102"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNIValue(0,1) = FilterVar(strBdgCd, "''", "S")
    UNIValue(0,2) = FilterVar(strDeptCd, "''", "S")
    UNIValue(0,3) = FilterVar(gChangeOrgId, "''", "S")
    UNIValue(0,4) = FilterVar(strBdgYymm, "''", "S")
    UNIValue(0,5) = FilterVar(strBdgCd, "''", "S")
    UNIValue(0,6) = FilterVar(strDeptCd, "''", "S")
    UNIValue(0,7) = FilterVar(gChangeOrgId, "''", "S")
    UNIValue(0,8) = FilterVar(strBdgYymm, "''", "S")
    UNIValue(0,9) = FilterVar(strBdgCd, "''", "S")
    UNIValue(0,10) = FilterVar(strDeptCd, "''", "S")
    UNIValue(0,11) = FilterVar(gChangeOrgId, "''", "S")
    UNIValue(0,12) = FilterVar(strBdgYymm, "''", "S")

    UNIValue(1,0) = FilterVar(strBdgCd, "''", "S")
    UNIValue(1,1) = FilterVar(strDeptCd, "''", "S")
    UNIValue(1,2) = gChangeOrgId
    UNIValue(1,3) = FilterVar(strBdgYymm, "''", "S")
    
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
	
	If rs1.EOF And rs1.BOF Then
	Else
%>
		<Script Language="vbScript">
			With parent.frm1
				.txtBdgPlanAmt.value = "<%=UNINumClientFormat(rs1(0), ggAmtOfMoney.DecPoint, 0)%>"
				.txtBdgAmt.value     = "<%=UNINumClientFormat(rs1(1), ggAmtOfMoney.DecPoint, 0)%>"
			End With
		</Script>
<%
	End If
	
	rs1.Close
	Set rs1 = Nothing
	
    If  rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Set lgADF = Nothing
'		Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
	Dim strInternalCd
	
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    strBdgYymm = Request("txtBdgYymm")
    strDeptCd  = UCase(Request("txtDeptCd"))
    strBdgCd   = UCase(Request("txtBdgCd"))
	strColYymm   = Request("txtColYymm")
	strDateType  = Request("txtDateType")
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

'내부부서코드 select
Function fnGetInternalCd()
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,1)

    UNISqlId(0) = "F2109PA102"

    UNIValue(0,0) = strDeptCd
    UNIValue(0,1) = gChangeOrgId
    
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
        fnGetInternalCd = rs0(2)
    End If
End Function

'----------------------------------------------------------------------------------------------------------
' Trim string and set string to space if string length is zero
'----------------------------------------------------------------------------------------------------------
'2004.8.19 comment처리 
'Function FilterVar(Byval str,Byval strALT)
'     Dim strL
'	    strL = UCase(Trim(str))
'  If Len(strL) Then
'        FilterVar = " " & FilterVar(strL , "''", "S") & ""
'     Else
'        FilterVar = strALT   
'     End If
'End Function

%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
         .lgStrPrevKey        = "<%=ConvSPChars(lgStrPrevKey)%>"                       '☜: set next data tag

         With .frm1
			.hBdgYymm.value = strBdgYymm
			.hDeptCd.value  = strDeptCd
			.hBdgCd.value   = strBdgCd
         End With
         
         Call .DbQueryOk
	End with
</Script>	

<%
	Response.End 
%>
