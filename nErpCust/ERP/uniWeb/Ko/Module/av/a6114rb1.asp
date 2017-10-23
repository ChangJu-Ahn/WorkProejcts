<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% 

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim rdoInOut
Dim txtFrIssuedDt
Dim txtToIssuedDt
Dim rdoIn
Dim rdoOut
Dim txtBPCd
Dim txtBPNm
Dim txtFrVatLocAmt
Dim txtToValLocAmt
Dim txtTempGlNo
Dim txtGlNo
Dim lgMaxCount

	Const C_SHEETMAXROWS_D = 40

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgPageNo       = Request("lgPageNo")                               '☜ : Next key flag
    lgMaxCount     = CInt(C_SHEETMAXROWS_D)                            '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"

	txtFrIssuedDt		= FilterVar(UNIConvDate(Request("txtFrIssuedDt")), "''", "S")
	txtToIssuedDt		= FilterVar(UNIConvDate(Request("txtToIssuedDt")), "''", "S")
	rdoInOut			= FilterVar(Request("rdoInOut"), "''", "S")
	txtBPCd				= FilterVar(Request("txtBPCd"), "''", "S")
	txtTempGlNo			= FilterVar(Request("txtTempGlNo"), "''", "S")
	txtGlNo				= FilterVar(Request("txtGlNo"), "''", "S")
	
	if Trim(Request("txtFrVatLocAmt")) = "" then 
		txtFrVatLocAmt = -999999999999
	else 
		txtFrVatLocAmt		= UniConvNum(Trim(Request("txtFrVatLocAmt")),0)
	end if 

	if Trim(Request("txtToValLocAmt")) = "" then 
			txtToValLocAmt = 999999999999
	else 
			txtToValLocAmt		= UniConvNum(Trim(Request("txtToValLocAmt")),0)
	end if 
	
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

    If Len(Trim(lgPageNo)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          lgPageNo = CInt(lgPageNo)
       End If   
    Else   
       lgPageNo = 0
    End If   

    rs0.PageSize     = lgMaxCount                                                'Seperate Page with page count (MA : C_SHEETMAXROWS_D )
    rs0.AbsolutePage = lgPageNo + 1



    iLoopCount = -1
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < lgMaxCount Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(0,9)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "A6114RAQ01"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	 UNIValue(0,1)  = txtFrIssuedDt
	 UNIValue(0,2)  = txtToIssuedDt
	 
	If Trim(Request("rdoInOut")) = "" Then
		UNIValue(0,3)  = ""
	Else 
		UNIValue(0,3)  = "  AND A.IO_FG = " & rdoInOut
	End If
	
	If Trim(Request("txtBPCd")) = "" then
		UNIValue(0,4)  = ""
	Else 
		UNIValue(0,4)  = "  AND B.BP_CD  = " & txtBPCd
	End If
	
		UNIValue(0,5)  = "  AND A.VAT_LOC_AMT >= " & txtFrVatLocAmt
	
		UNIValue(0,6)  = "  AND A.VAT_LOC_AMT <= " & txtToValLocAmt
	
	If Trim(Request("txtTempGlNo")) = "" Then
		UNIValue(0,7) = ""
	Else 
		UNIValue(0,7) = "  AND A.TEMP_GL_NO = " & txtTempGlNo	
	End If
	
	If Trim(Request("txtGlNo")) = "" Then
		UNIValue(0,8) = ""
	Else
		UNIValue(0,8) = "  AND A.GL_NO  = " & txtGlNo
	End If				

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
on error resume next
    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    Set lgADF = Nothing  
    
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

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
  '        Parent.Frm1.txtFrIssuedDt.Value      = Parent.Frm1.txtSchoolCd.Value                  'For Next Search
  '        Parent.Frm1.htxtHanyun.Value        = Parent.Frm1.txtHanyun.Value
  '        Parent.Frm1.htxtBan.Value           = Parent.Frm1.txtBan.Value
  '        Parent.Frm1.hfpdtFromEnterDt.Value  = Parent.Frm1.fpdtFromEnterDt.Text
  '        Parent.Frm1.hfpdtToEnterDt.Value    = Parent.Frm1.fpdtToEnterDt.Text
       End If
       
       'Show multi spreadsheet data from this line
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
    End If   

</Script>	

