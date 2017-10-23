<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% 


On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2
Dim lgstrData																		'☜ : data for spreadsheet data
Dim lgStrPrevKey																	'☜ : 이전 값 
Dim lgMaxCount																		'☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList																		'☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo



Dim txtYyyymm
Dim txtModuleCd
Dim txtBizAreaCd

Dim strMsgCd, strMsg1, strMsg2



Dim iPrevEndRow
Dim iEndRow

' 권한관리 추가
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					


    Call LoadBasisGlobalInf()    

    Call LoadInfTB19029B("Q","A","NOCOOKIE","QB")   
    Call LoadBNumericFormatB("Q", "A","NOCOOKIE","QB") 

    Call HideStatusWnd 


    lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount		= CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList	= Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList		= Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist		= "No"

    txtYyyymm		= Trim(Request("txtYyyymm"))
    txtModuleCd		= Trim(Request("txtModuleCd"))
    txtBizAreaCd	= Trim(Request("txtBizAreaCd"))


    If txtBizAreaCd = "" then
       txtBizAreaCd = "%"
    End If

    If txtModuleCd = "" then
       txtModuleCd = "%"
    End If

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

    iPrevEndRow = 0

    If CDbl(lgPageNo) > 0 Then
		iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)    
		rs0.Move= iPrevEndRow                 
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
 
        If  iLoopCount < lgMaxCount Then
            lgstrData		=	lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop
    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
                                                 '☜: SQL ID 저장을 위한 영역확보 
    Dim strWhere,strWhere2

    Redim UNISqlId(7)   
    Redim UNIValue(1,2)                                                 '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "A5312MA441"
    UNISqlId(1) = "A5312MA441S"

    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
    strWhere = ""
    strWhere2 = ""



    strWhere =  FilterVar(txtYyyymm ,"''","S") 
    strWhere = strWhere & " and k.module_cd like " & FilterVar(txtModuleCd ,"''"	,"S") 
    strWhere = strWhere & " and k.biz_area_cd like  " & FilterVar(txtBizAreaCd ,"''"	,"S") 

    strWhere2 =  FilterVar(txtYyyymm ,"''","S") 
    strWhere2 = strWhere2 & " and module_cd like " & FilterVar(txtModuleCd ,"''"	,"S") 
    strWhere2 = strWhere2 & " and biz_area_cd like  " & FilterVar(txtBizAreaCd ,"''"	,"S") 

      
 
    UNIValue(0,1)  = strWhere

    UNIValue(1,1)  = strWhere2

     '   Call ServerMesgBox(UNIValue(0,1), vbInformation, I_MKSCRIPT)
	




    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        

	If Not (rs1.EOF And rs1.BOF) Then
%>
		<Script Language=vbScript>
			With Parent
				.frm1.txtAmtSum1.value = "<%=UNINumClientFormat(rs1(0), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum2.value = "<%=UNINumClientFormat(rs1(1), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum3.value = "<%=UNINumClientFormat(rs1(2), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum4.value = "<%=UNINumClientFormat(rs1(3), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum5.value = "<%=UNINumClientFormat(rs1(4), ggAmtOfMoney.DecPoint, 0)%>"



			End With
		</Script>
<%
	End If

    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If

End sub

%>

<Script Language=vbscript>
 
	With Parent
	

		If "<%=lgDataExist%>" = "Yes" Then
		   'Show multi spreadsheet data from this line
		   
	
		   .ggoSpread.Source  = Parent.frm1.vspdData
		   .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
		   .lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		   .DbQueryOk
		   .frm1.vspdData.Redraw = True
		End If   
    
    End With

</Script>	
	

