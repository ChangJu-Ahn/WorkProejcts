<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3112rb7.asp																*
'*  4. Program Name         :																			*
'*  5. Program Desc         : 반품에서 발주참조 
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2003/05/28																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"										*
'*                            this mark(☆) Means that "must change"										*
'* 13. History              : 1. 2000/04/08 : Coding Start												*
'********************************************************************************************************
On Error Resume Next

   Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1           '☜ : DBAgent Parameter 선언 
   Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
   Dim iTotstrData
   Dim lgTailList
   Dim lgSelectList
   Dim lgSelectListDT
   Dim lgDataExist
   Dim lgPageNo
   
   Dim iPrevEndRow
   Dim iEndRow

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
    Redim UNISqlId(1)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(1,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
    UNISqlId(0) = "M3112RA701" 
    UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 
	strVal = " "
	
	If Len(Request("txtClsflg")) Then
		strVal =  strVal & " AND A.CLS_FLG = " & FilterVar(Trim(UCase(Request("txtClsflg"))), " " , "S") & " "
	End If

	If Len(Request("txtreleaseflg")) Then
		strVal =  strVal & " AND B.RELEASE_FLG = " & FilterVar(Trim(UCase(Request("txtreleaseflg"))), " " , "S") & " "
	End If

	If Trim(Request("txtRcptflg")) = "Y" Then
		strVal =  strVal & " AND B.RCPT_TYPE = " & FilterVar(Trim(UCase(Request("txtRcptType"))), " " , "S") & " "
		strVal =  strVal & " AND B.RET_FLG =  " & FilterVar(Trim(UCase(Request("txtRetflg"))), " " , "S") & " "
		strVal =  strVal & " AND B.RCPT_FLG = " & FilterVar(Trim(UCase(Request("txtRcptflg"))), " " , "S") & " "
		strVal =  strVal & " AND A.AFTER_LC_FLG <> " & FilterVar("N", "''", "S") & " " 
	
	ElseIf 	Trim(Request("txtRcptflg")) = "N" Then 'Trim(Request("hdnRetflg")) = "N" Then
		strVal =  strVal & " AND B.RCPT_TYPE = " & FilterVar(Trim(UCase(Request("txtRcptType"))), " " , "S") & " "
		strVal =  strVal & " AND B.RET_FLG =  " & FilterVar(Trim(UCase(Request("txtRetflg"))), " " , "S") & " "
		strVal =  strVal & " AND B.RCPT_FLG = " & FilterVar(Trim(UCase(Request("txtRcptflg"))), " " , "S") & " "
	End If		


	If Len(Request("txtSupplier")) Then
		strVal = strVal & " AND B.BP_CD = " & FilterVar(Trim(UCase(Request("txtSupplier"))), " " , "S") & " "
	End If

    If Len(Trim(Request("txtPoNo"))) Then
		strVal = strVal & " AND B.PO_NO = " & FilterVar(Trim(UCase(Request("txtPoNo"))), " " , "S") & " "		
	End If		
	
    If Len(Trim(Request("txtFrPoDt"))) Then
		strVal = strVal & " AND B.PO_DT >= " & FilterVar(UNIConvDate(Request("txtFrPoDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToPoDt"))) Then
		strVal = strVal & " AND B.PO_DT <= " & FilterVar(UNIConvDate(Request("txtToPoDt")), "''", "S") & ""		
	End If

     If Request("gPlant") <> "" Then
        strVal = strVal & " AND a.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND b.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND b.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     If Request("gBizArea") <> "" Then
        strVal = strVal & " AND b.PUR_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If   

    UNIValue(0,1) = strVal 
   
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
        
    Set lgADF   = Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
 
    If rs0.EOF And rs0.BOF Then
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
With Parent
    If "<%=lgDataExist%>" = "Yes" Then
       
       If "<%=lgPageNo%>" = "1" Then   
			.frm1.hdnFrPoDt.Value 		= "<%=Request("txtFrPoDt")%>"
		    .frm1.hdnToPoDt.Value 		= "<%=Request("txtToPoDt")%>"
		    .frm1.hdnPoNo.Value 		= "<%=ConvSPChars(Request("txtPoNo"))%>"
       End If
       
       .ggoSpread.Source  = .frm1.vspdData
       .frm1.vspdData.Redraw = False
       .ggoSpread.SSShowData "<%=iTotstrData%>","F"          '☜ : Display data
       
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,.GetKeyPos("A",13),.GetKeyPos("A",11),"C","I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,.GetKeyPos("A",13),.GetKeyPos("A",12),"A","I","X","X")
	   .lgPageNo      =  "<%=lgPageNo%>"               
       .DbQueryOk
       .frm1.vspdData.Redraw = True
    End If  
End With 
</Script>	
