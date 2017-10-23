<%
'======================================================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 수주관리 
'*  3. Program ID           : S3135QB1
'*  4. Program Name         : 수주진행별조회 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/08
'*  8. Modified date(Last)  : 2002/02/15
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : Ahn tae hee
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Date표준적용 
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "MB")
On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgPageNo																'☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Const C_SHEETMAXROWS_D  = 100                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim arrRsVal(1)
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
Call HideStatusWnd 

lgPageNo	   = UNICInt(Trim(Request("lgPageNo")),0)
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
	Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgstrData      = ""
    
    If CInt(lgPageNo) > 0 Then
       rs0.Move  =  C_SHEETMAXROWS_D * CInt(lgPageNo)
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                             '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(1,2)

    UNISqlId(0) = "S3135QA101"
    UNISqlId(1) = "s0000qa001"
   

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "
	
	strVal = "AND A.SO_QTY > 0 "
	
	If Len(Request("txtTrackingNo")) Then
		strVal = strVal & "AND A.TRACKING_NO = " & FilterVar(Request("txtTrackingNo"), "''", "S") & " "
	End If

	If Len(Request("txtItemCode")) Then
		strVal = strVal & " AND A.ITEM_CD = " & FilterVar(Request("txtItemCode"), "''", "S") & " "		
		arrVal = Request("txtItemCode")
	End If		
		   
	If Len(Request("txtSoNo")) Then
		strVal = strVal & " AND A.SO_NO = " & FilterVar(Request("txtSoNo"), "''", "S") & " "
	End If		

    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar(arrVal, "" , "S")
        
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
	Dim FalsechkFlg
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

    FalsechkFlg = False
    
    iStr = Split(lgstrRetMsg,gColSep)
	
	If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtItemCode")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtItemCode.focus    
            </Script>
        <%	       	
		End If	
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtTrackingNo.focus    
            </Script>
        <%        
    Else   
        Call  MakeSpreadSheetData()
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
        .ggoSpread.Source    = .frm1.vspdData 
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>"			'☜: Display data 
        .lgPageNo        =  "<%=lgPageNo%>"		 '☜: set next data tag         
      	.frm1.txtHTrackingNo.value = "<%=ConvSPChars(Request("txtTrackingNo"))%>"
		.frm1.txtHSoNo.value = "<%=ConvSPChars(Request("txtSoNo"))%>"
		.frm1.txtHItemCode.value = "<%=ConvSPChars(Request("txtItemCode"))%>"            
        .frm1.txtItemCodeNm.value = "<%=ConvSPChars(arrRsVal(1))%>"  
		.DbQueryOk      
    End with
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
