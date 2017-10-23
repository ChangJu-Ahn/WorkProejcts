
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%             

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("I", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...

Call HideStatusWnd 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                        '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim iPrevEndRow
Dim iEndRow	

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim LngRow
Dim GroupCount    
Dim strVal

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 

'Dim strFrArDt	                                                           
'Dim strToArDt
Dim strDocCur                                                          
Dim strPayBpCd
Dim strBizCd
Dim strAllcDt

Dim strCond

' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 
	
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    iPrevEndRow = 0
    iEndRow = 0
       
    Call TrimData()
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
		iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)    
       rs0.Move= iPrevEndRow                   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If

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

   Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "A4112RA101"
    UNISqlId(1) = "COMMONQRY"
    UNISqlId(2) = "COMMONQRY"

    Redim UNIValue(2,2)

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	UNIValue(0,1) = strCond
	UNIValue(1,0) = "SELECT BP_NM FROM B_BIZ_PARTNER WHERE BP_CD =  " & FilterVar(UCase(Request("txtBpCd")), "''", "S") & " "
    UNIValue(2,0) = "SELECT BIZ_AREA_NM FROM B_BIZ_AREA WHERE BIZ_AREA_CD =  " & FilterVar(UCase(Request("txtBizCd")), "''", "S") & " "                 
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub



'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim strMsg
    Dim strMsg1
    Dim strMsgCd
    Dim strMsgCd1
    
    strMsg = Trim(Request("txtBpcd_Alt"))
    strMsg1 = Trim(Request("txtBizCd_Alt"))
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    Set lgADF = Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If
	
	IF NOT (rs1.EOF or rs1.BOF) then
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = "<%=Trim(ConvSPChars(rs1(0)))%>"   
		End With
		</Script>
<%			
	ELSE
		if Trim(Request("txtBpCd")) <> "" Then
			strMsgCd = "970000"
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = ""   
		End With
		</Script>
<%	
		Else 
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBpNm.Value  = ""   
		End With
		</Script>
<%			
		End if
	End if
    rs1.Close
    Set rs1 = Nothing 
    
    'rs2에 대한 결과 
    IF NOT (rs2.EOF or rs2.BOF) then
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBizNm.Value = "<%=Trim(ConvSPChars(rs2(0)))%>"   
		End With
		</Script>
<%			    
	ELSE
		if Trim(Request("txtBizCd")) <> "" Then
			strMsgCd1 = "970000"
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBizNm.Value = ""   
		End With
		</Script>
<%	
		Else
%>
		<Script Language=vbScript>
		With parent
			.Frm1.txtBizNm.Value = ""   
		End With
		</Script>
<%		
		End if
    END IF
    rs2.Close
    Set rs2 = Nothing
	
	If  "" & Trim(strMsgCd) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg, "", I_MKSCRIPT)
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    End If
    
    If  "" & Trim(strMsgCd1) <> "" Then
		Call DisplayMsgBox("970000", vbOKOnly, strMsg1, "", I_MKSCRIPT)
        Response.End													'☜: 비지니스 로직 처리를 종료함 
    End If
	
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Response.End													'☜: 비지니스 로직 처리를 종료함 
	Else
		Call  MakeSpreadSheetData()
    End If				
    
    Set rs0 = Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub  TrimData()
	
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
'     strFrArDt     = Ucase(Trim(UNIConvDate(Request("txtArDt"))))
 '    strToArDt     = Ucase(Trim(UNIConvDate(Request("txtToArDt"))))
    strDocCur	   = UCase(Trim(Request("txtDocCur")))
    strPayBpCd    = UCase(Trim(Request("txtBpCd")))
    strBizCd	   = UCase(Trim(Request("txtBizCd")))
    strAllcDt		= UNIConvDate(Trim(Request("txtAllcDt"))) 
     
	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))     
     
    If strFrArDt <> "" Then	:		strCond = strCond & " and A.AR_DT >=  " & FilterVar(strFrArDt , "''", "S") & ""
     
    If strToArDt <> "" Then	:		strCond = strCond & " and A.AR_DT <=  " & FilterVar(strToArDt , "''", "S") & ""

    If strDocCur <> "" Then	:		strCond = strCond & " and A.doc_cur =  " & FilterVar(strDocCur , "''", "S") & ""

    If strPayBpCd <> "" Then	:		strCond = strCond & " and A.DEAL_BP_CD =  " & FilterVar(strPayBpCd , "''", "S") & ""

    If strBizCd <> "" Then	:		strCond = strCond & " and A.biz_area_cd =  " & FilterVar(strBizCd , "''", "S") & ""

    strCond = strCond & " AND A.bal_amt <> 0 AND A.gl_no <> '' "    
    strCond = strCond & " AND A.ar_dt <=  " & FilterVar(strAllcDt , "''", "S") & "" 
     
	If "" & Trim(Request("txtArDt")) <> "" Then	
		strCond = strCond & " AND A.ar_dt >=  " & FilterVar(UNIConvDate(Trim(Request("txtArDt"))), "''", "S") & "" 
	End if

	If "" & Trim(Request("txtToArDt")) <> "" Then	
		strCond = strCond & " AND A.ar_dt <=  " & FilterVar(UNIConvDate(Trim(Request("txtToArDt"))), "''", "S") & "" 
	End if

	If "" & Trim(Request("txtArDueDt")) <> "" Then	
		strCond = strCond & " AND A.ar_Due_dt >=  " & FilterVar(UNIConvDate(Trim(Request("txtArDueDt"))), "''", "S") & "" 
	End if

	If "" & Trim(Request("txtToArDueDt")) <> "" Then	
		strCond = strCond & " AND A.ar_Due_dt <=  " & FilterVar(UNIConvDate(Trim(Request("txtToArDueDt"))), "''", "S") & "" 
	End if

	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then
		strCond = strCond & " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
	End If
	
	If lgInternalCd <> "" Then
		strCond = strCond & " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
	End If
	
	If lgSubInternalCd <> "" Then
		strCond = strCond & " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
	End If
	
	If lgAuthUsrID <> "" Then
		strCond = strCond & " AND A.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
	End If     
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

End Sub

%>
<Script Language=vbscript>
    If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
          parent.Frm1.htxtBizCd.Value		= Parent.Frm1.txtBizCd.Value
          Parent.Frm1.htxtBpCd.Value        = Parent.Frm1.txtBpCd.Value
          Parent.Frm1.htxtArDt.Value		= Parent.Frm1.txtArDt.Text
          Parent.Frm1.htxtToArDt.Value		= Parent.Frm1.txtToArDt.Text
		  Parent.frm1.htxtArDueDt.Value		= Parent.frm1.txtArDueDt.Text
		  Parent.frm1.htxtToArDueDt.Value		= Parent.frm1.txtToArDueDt.Text          
          Parent.Frm1.htxtDocCur.Value		= Parent.Frm1.txtDocCur.Value
       End If
       
       'Show multi spreadsheet data from this line
       
       Parent.ggoSpread.Source  = Parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data

       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",10),Parent.GetKeyPos("A",12),"A", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",10),Parent.GetKeyPos("A",14),"A", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",10),Parent.GetKeyPos("A",18),"A", "I" ,"X","X")
       Parent.frm1.vspdData.Redraw = True
       Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       Parent.DbQueryOk
    End If   

</Script>
	
