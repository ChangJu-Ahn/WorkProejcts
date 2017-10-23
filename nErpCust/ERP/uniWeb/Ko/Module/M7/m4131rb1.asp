<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3112rb5.asp															*
'*  4. Program Name         : 발주내역참조(입고등록)ADO																			*
'*  5. Program Desc         : Purchase Order Detail 참조 PopUp ASP										*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2003/05/27																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : Kim Jin HaL																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"										*
'*                            this mark(☆) Means that "must change"										*
'* 13. History              : 1. 2000/04/08 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%

On Error Resume Next

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3           '☜ : DBAgent Parameter 선언 
   Dim lgstrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
   Dim iTotstrData
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   
   Dim iPrevEndRow
   Dim iEndRow
    
   Dim strShiptoPartyNm
   Dim strPlantNm
   Dim strSlNm
   
    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
     
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
    lgDataExist      = "No"
	iPrevEndRow = 0
    iEndRow = 0
    
    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim PvArr
    
    Const C_SHEETMAXROWS_D = 100     
    
    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow                 

    End If

    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
            PvArr(iLoopCount) = lgstrData	
		    lgstrData = ""
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")
	
    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        iEndRow = iPrevEndRow + iLoopCount + 1
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
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

    Dim strVal
    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(0,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
   
	IF  Trim(Request("txtChkFlg")) = "Y" THEN                                                                           '    parameter의 수에 따라 변경함 
		UNISqlId(0) = "m4131ra101" 
	ELSE 
		UNISqlId(0) = "m4131ra102" 
	END IF     
     
     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	strVal = " "

	If Len(Request("txtRcptNo")) Then
		strVal = " AND A.MVMT_RCPT_NO = " & FilterVar(Request("txtRcptNo"), "''", "S") & " "
	Else
		strVal = ""
	End If

    UNIValue(0,1) = strVal   
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If rs0.EOF And rs0.BOF Then
		
		Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
			 
		lgStrSQL = "Select MVMT_RCPT_NO " 
		lgStrSQL = lgStrSQL & " From  M_PUR_GOODS_MVMT "
		lgStrSQL = lgStrSQL & " WHERE MVMT_RCPT_NO=  " & FilterVar(Request("txtRcptNo"), "''", "S") & "  "
		

		call FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") 

		If lgObjRs.EOF AND lgObjRs.BOF Then
			  Call DisplayMsgBox("174110", vbOKOnly, "", "", I_MKSCRIPT)		      
		Else		  
			  Call DisplayMsgBox("179027", vbOKOnly, "", "", I_MKSCRIPT)		 
		End If
		
		Call SubCloseRs(lgObjRs) 		
		Set rs0 = Nothing
		Exit Sub
			
    Else    
        Call  MakeSpreadSheetData()
    
    End If
    
End Sub

%>
<Script Language=vbscript>

    If "<%=lgDataExist%>" = "Yes" Then
       If "<%=lgPageNo%>" = "1" Then  
			parent.frm1.hdnRcptNo.value	= "<%=ConvSPChars(Request("txtRcptNo"))%>"   
       End If
       
       parent.ggoSpread.Source  = parent.frm1.vspdData
       parent.frm1.vspdData.Redraw = false
       parent.ggoSpread.SSShowData "<%=iTotstrData%>"          '☜ : Display data
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       parent.DbQueryOk
       parent.frm1.vspdData.Redraw = true
    End If   
</Script>	
