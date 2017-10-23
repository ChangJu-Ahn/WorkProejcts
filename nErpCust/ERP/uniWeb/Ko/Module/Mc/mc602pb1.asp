<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Procuremant
'*  2. Function Name        : 
'*  3. Program ID           : mc602pa1
'*  4. Program Name         : L/C Reference ASP		
'*  5. Program Desc         : L/C Reference ASP		
'*  6. Component List       : 
'*  7. Modified date(First) : 2003/02/28	
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Ahn Jung Je	
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
<%

	On Error Resume Next													   '실행 오류가 발생할 때 오류가 발생한 문장 바로 다음에 실행이 계속될 수 있는 문으로 컨트롤을 옮길 수 있도록 지정합니다.				

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '☜ : DBAgent Parameter 선언 
	Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
	Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim iTotstrData

	Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "PB")
	
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist      = "No"
 
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
' Query하기 전에  DB Agent 배열을 이용하여 Query문을 만드는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
	
	Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
	Redim UNIValue(0,2)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 
	
	
    UNISqlId(0) = "M4111PA301"											' main query(spread sheet에 뿌려지는 query statement)
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
	
	strVal = ""
	If Len(Request("txtFrRcptDt")) Then 
		strVal = strVal & " AND A.MVMT_RCPT_DT >=  " & FilterVar(UNIConvDate(Request("txtFrRcptDt")), "''", "S") & " "
	End If		
	
	If Len(Request("txtToRcptDt")) Then
		strVal = strVal & " AND A.MVMT_RCPT_DT <=  " & FilterVar(UNIConvDate(Request("txtToRcptDt")), "''", "S") & " "
	End If
	
	If Len(Request("cboMvmtType")) Then
		strVal = strVal & " AND B.IO_TYPE_CD =  " & FilterVar(Request("cboMvmtType"), "''", "S") & " "
	end if	    
    		
	If Len(Trim(Request("txtSupplier"))) Then
		strVal = strVal & " AND A.BP_CD =  " & FilterVar(Request("txtSupplier"), "''", "S") & " "
	End If
	
	IF LEN(Trim(Request("txtGroup"))) THEN
		strVal = strVal & " AND D.PUR_GRP =  " & FilterVar(Request("txtGroup"), "''", "S") & " "
	END IF	 

	IF Request("txtFlag") = "MC" THEN
		strVal = strVal & " AND A.DLVY_ORD_FLG = " & FilterVar("Y", "''", "S") & "  "
	END IF

	
	UNIValue(0,1) = strVal											'	UNISqlId(0)의 두번째 ?에 입력됨	
    UNIValue(0,2) = UCase(Trim(lgTailList))							'	UNISqlId(0)의 마지막 ?에 입력됨	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                '☜: set ADO read mode
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If  
   
End Sub
    
'----------------------------------------------------------------------------------------------------------
'QueryData()에 의해서 Query가 되면 MakeSpreadSheetData()에 의해서 데이터를 스프레드시트에 뿌려주는 프로시져 
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    Dim PvArr
    Const C_SHEETMAXROWS_D = 100 
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = C_SHEETMAXROWS_D * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
           PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")

    If iLoopCount < C_SHEETMAXROWS_D Then                                 '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

	Response.Write "<Script Language=vbscript> " & vbCr   
	Response.Write " With Parent "               & vbCr
	
	Response.Write "	If """ & lgDataExist & """  = ""Yes"" Then " & vbCr  
	Response.Write "		If """ & lgPageNo & """ = ""1"" Then " & vbCr  
	Response.Write "			.frm1.hdnMvmtType.value  = """ & ConvSPChars(Request("cboMvmtType")) & """" & vbCr
	Response.Write "			.frm1.hdnSupplier.value  = """ & ConvSPChars(Request("txtSupplier")) & """" & vbCr
	Response.Write "			.frm1.hdnFrRcptDt.value  = """ & Request("txtFrRcptDt") & """" & vbCr
	Response.Write "			.frm1.hdnToRcptDt.value = """ & Request("txtToRcptDt") & """" & vbCr
	Response.Write "			.frm1.hdnGroup.value  = """ & ConvSPChars(Request("txtGroup")) & """" & vbCr
	Response.Write "		End If       " & vbCr                    
	Response.Write "	End If       " & vbCr                    	    
	
	Response.Write "	.ggoSpread.Source     = .frm1.vspdData "    & vbCr
	Response.Write "	.ggoSpread.SSShowData """ & iTotstrData  & """" & vbCr
	Response.Write "	.lgPageNo  = """ & lgPageNo & """" & vbCr
	Response.Write "	.DbQueryOk "                                                                                                   & vbCr  
	Response.Write "End With       " & vbCr                    
	Response.Write "</Script>      " & vbCr   
	    
	Response.End 

%>
