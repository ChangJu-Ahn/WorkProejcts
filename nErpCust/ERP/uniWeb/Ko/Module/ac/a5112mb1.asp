<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<%

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1 , rs2                   '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim txtBizArea
Dim lgBizAreaCd
Dim lgBizAreaNm

Dim txtClassType
Dim lgClassTypeCd
Dim lgClassTypeNm

Dim lgSp_Id

Const C_SHEETMAXROWS_D  = 100

Call HideStatusWnd 

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					


lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
lgDataExist    = "No"

'''uniCode 관련 수정 
txtBizArea	   = Trim(Request("txtBizArea"))
txtClassType   = Trim(Request("txtClassType"))

lgSp_Id   = Trim(Request("strSp_Id"))
	
' 권한관리 추가 
lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))		
lgInternalCd		= Trim(Request("lgInternalCd"))	
lgSubInternalCd		= Trim(Request("lgSubInternalCd"))	
lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

Call FixUNISQLData()
Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr

    lgDataExist    = "Yes"
    lgstrData      = ""

    If CDbl(lgPageNo) > 0 Then
       rs0.Move     = CDbl(C_SHEETMAXROWS_D) * CDbl(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If

	rs0.Close
    Set rs0 = Nothing 

    If Not( rs1.EOF OR rs1.BOF) Then

   		lgBizAreaCd = rs1(0)
		lgBizAreaNm = rs1(1)
    End IF

    rs1.Close
    Set rs1= Nothing

    If Not( rs2.EOF OR rs2.BOF) Then

   		lgClassTypeCd = rs2(0)
		lgClassTypeNm = rs2(1)
    End IF
    rs2.Close
    Set rs2= Nothing

End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)

    Redim UNIValue(2,2)

    UNISqlId(0) = "a5112QA101"
    UNISqlId(1) = "A_GetBiz"
    UNISqlId(2) = "A_CLSTYPE"


    UNIValue(0,0) = lgSelectList
    UNIValue(0,1) = FilterVar(lgSp_Id, "''", "S")
    UNIValue(0,2) = " ORDER BY LIST_SEQ " 
    

    If txtBizArea = "" Then
	 	UNIValue(1,0) = FilterVar("", "''", "S")
	Else
		UNIValue(1,0) = FilterVar(txtBizArea, "''", "S")
	End If

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then			
		UNIValue(1,0)  = UNIValue(1,0) & " AND BIZ_AREA_CD LIKE " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			
	
	UNIValue(2,0) = FilterVar(txtClassType, "''", "S") 

    'UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg
    Dim iStr
    Dim lgADF

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)

    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing

    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else
        Call  MakeSpreadSheetData()
    End If
End Sub

%>


<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Show multi spreadsheet data from this line

       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
       parent.frm1.txtBizAreaNm.value = "<%=ConvSPChars(lgBizAreaNm)%>"
       parent.frm1.txtClassTypeNm.value = "<%=ConvSPChars(lgClassTypeNm)%>"
       Parent.lgPageNo      =  "<%=lgPageNo%>"

       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			Parent.Frm1.txtBizArea.Value      = Parent.Frm1.txtBizArea.Value                  'For Next Search
			Parent.Frm1.txtClassType.Value    = Parent.Frm1.txtClassType.Value
		End If

       Parent.DbQuery2Ok

    End If
</Script>
