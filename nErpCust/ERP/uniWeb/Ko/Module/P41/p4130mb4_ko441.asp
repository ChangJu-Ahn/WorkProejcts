<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name			: Production
'*  2. Function Name		: 
'*  3. Program ID			: p4213mb4.asp
'*  4. Program Name			: Cancel Release Production Order
'*  5. Program Desc			: Cancel Release Production Order
'*  6. Comproxy List		: PP4G255.cPCnclRlse
'*  7. Modified date(First) : 2000/05/25
'*  8. Modified date(Last)  : 2003/02/04
'*  9. Modifier (First)     : Im, Hyun Soo
'* 10. Modifier (Last)      : Chen, Jae Hyun
'* 11. Comment              :
'**********************************************************************************************
'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%														'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
Call LoadBasisGlobalInf
																	
Call HideStatusWnd

On Error Resume Next

Dim oPP4G255						'PP4G255.cPCnclRlse
Dim iErrorPosition
Dim itxtSpread
Dim itxtSpreadArr
Dim itxtSpreadArrCount
Dim iCUCount
Dim ii
Dim if_error                    '20080115::hanc::에러발생여부 ('E' : error 발생)

if_error    = ""
itxtSpread  = ""
             
iCUCount = Request.Form("txtCUSpread").Count
             
itxtSpreadArrCount = -1
             
ReDim itxtSpreadArr(iCUCount)
             
For ii = 1 To iCUCount
    itxtSpreadArrCount = itxtSpreadArrCount + 1
    itxtSpreadArr(itxtSpreadArrCount) = Request.Form("txtCUSpread")(ii)
Next

itxtSpread = Join(itxtSpreadArr,"")


Err.Clear																		'☜: Protect system from crashing

    Call SubOpenDB(lgObjConn)                                                  '20080114::hanc  '☜: Make a DB Connection
    Call SubBizSaveMulti()                                                     '20080114::hanc
    Call SubCloseDB(lgObjConn)                                                 '20080114::hanc      '☜: Close DB Connection

'20080114::hanc =============================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()

    Dim arrRowVal
    Dim arrColVal
    Dim lgErrorPos
    Dim iDx

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    lgErrorPos        = ""                                                           '☜: Set to space

    
    If itxtSpread = "" Then
       Exit Sub
    End If   
    
  	lgStrSQL = " BEGIN TRAN "
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords

	arrRowVal = Split(itxtSpread, gRowSep)                                 '☜: Split Row    data

	
    For iDx = 0 To UBound(arrRowVal,1) - 1
        arrColVal = Split(arrRowVal(iDx), gColSep)                                 '☜: Split Column data
        
        Call SubBizSaveMultiCreate(arrColVal)                        '☜: Create
        
        If if_error    = "E" Then
           Exit For
        End If

    Next

    If 	if_error    = "" Then                    
        lgStrSQL = "COMMIT TRAN  "
        lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
        Response.Write  " <Script Language=vbscript> " & vbCr
        Response.Write  " Parent.DBSaveOk            " & vbCr
        Response.Write  " </Script>                  " & vbCr
  	else
        lgStrSQL = "ROLLBACK TRAN "
        lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
        Response.Write  " <Script Language=vbscript> " & vbCr
        Response.Write  " Parent.DBSaveOk            " & vbCr
'        Response.Write  " Parent.SubSetErrPos(Trim(""" & lgErrorPos & """))" & vbCr
        Response.Write  " </Script>                  " & vbCr
    End If
    
End Sub  


'20080114::hanc =============================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveMultiCreate(arrColVal)

    Dim lgStrSQL
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to create record
    '--------------------------------------------------------------------------------------------------------

    lgStrSQL = " EXEC USP_4130ma1_ko441 "   & FilterVar(arrcolval(00), "''", "S") & ", " _
                                            & FilterVar(arrcolval(01), "''", "S") & ", " _
                                            & FilterVar(arrcolval(02), "''", "S") & ", " _
                                            & FilterVar(REPLACE(arrcolval(03), ",", ""), "''", "D") & ", " _
                                            & FilterVar(arrcolval(04), "''", "S") & ", " _
                                            & FilterVar(arrcolval(05), "''", "S") & ", " _
                                            & FilterVar(arrcolval(06), "''", "S") & ", " _
                                            & FilterVar(arrcolval(07), "''", "S") & ", " _
                                            & FilterVar(arrcolval(08), "''", "S") & " "
    
    
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    

    If CheckSYSTEMError(Err,True) = True Then
        if_error = "E"

'		lgStrSQL = " Rollback Tran "
'		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
    Else
        if_error = ""

'		lgStrSQL = " Commit Tran "		
'		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords    
'		Response.Write  " <Script Language=vbscript> " & vbCr
'		Response.Write  "	Parent.DBSaveOk			" & vbCr
'		Response.Write  " </Script>                  " & vbCr
    End If   

End Sub

%>

