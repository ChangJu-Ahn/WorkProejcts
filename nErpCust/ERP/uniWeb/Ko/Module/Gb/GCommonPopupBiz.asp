<!--
======================================================================================================
*  1. Module Name          : P&L Mgmt.
*  2. Function Name        :
*  3. Program ID           : GCommonPopup
*  4. Program Name         : 경영손익 작업실행 에러 조회화면
*  5. Program Desc         : 경영손익 작업실행 에러 조회화면
*  6. Comproxy List        :
*  7. Modified date(First) : 2002/01/04
*  8. Modified date(Last)  : 2002/01/04
*  9. Modifier (First)     : Kwon Ki Soo
* 10. Modifier (Last)      : Kwon Ki Soo
* 11. Comment              :
* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
*                            this mark(⊙) Means that "may  change"
*                            this mark(☆) Means that "must change"
* 13. History              :
=======================================================================================================-->
<!-- #Include file="../../inc/IncServer.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다
    Dim ADF														'ActiveX Data Factory 지정 변수선언
    Dim strRetMsg												'Record Set Return Message 변수선언
    Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'DBAgent Parameter 선언
	Dim StrData
	Dim iLoop,jLoop
	Dim isOverFlowKey
	Dim isOverFlowName

    Const C_SHEETMAXROWS = 30									'한화면에 보일수 있는 최대 Row 수


    Call HideStatusWnd

If Request("arrField") <> "" Then
	Dim strSelect					'SELECT 할 Field 선언위한 변수
	Dim strTable					'SELECT 하고자하는 Table을 위한 변수
	Dim strWhere					'SELECT 하고자하는 SQL문장의 WHERE 조건을 위한 변수
	Dim intDataCount

	Redim UNISqlId(0)
	Redim UNIValue(0, 2)

	intDataCount = Request("gintDataCnt")
	strTable     = Request("txtTable")
	strWhere     = Request("txtWhere")

    strSelect = replace(Request("arrField"),gColSep,",")
    strSelect = Left(strSelect,Len(Trim(strSelect)) - 1)

	UNISqlId(0) = "compopup"
	UNIValue(0, 0) = strSelect
	UNIValue(0, 1) = strTable
	UNIValue(0, 2) = strWhere

	UNILock = DISCONNREAD :	UNIFlag = "1"

    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
	If Not (rs0.EOF And rs0.BOF) Then

       isOverFlowKey  = ""
       isOverFlowName = ""
       strData        = ""

       For iLoop = 0 to rs0.RecordCount-1
         If iLoop < C_SHEETMAXROWS Then
		    For jLoop = 0 To intDataCount - 1
                strData = strData & Chr(11) & rs0(jLoop)
            Next
			strData = strData & Chr(11) & Chr(12)
         Else
		    isOverFlowKey  = rs0(0)
			isOverFlowName = rs0(1)
			Exit For
		End If
        rs0.MoveNext
	   Next
	End If

    rs0.Close
    Set rs0 = Nothing
    Set ADF = Nothing
	
End If
%>
<Script Language="vbscript">
  On Error Resume Next
	With parent	
        .ggoSpread.SSShowData  "<%=ConvSPChars(strData)%>"               
        .lgStrCodeKey        = "<%=ConvSPChars(isOverFlowKey)%>"       
        .lgStrNameKey        = "<%=ConvSPChars(isOverFlowName)%>"       
        .vspdData.focus
        .DbQueryOk()
	End With

</Script>

