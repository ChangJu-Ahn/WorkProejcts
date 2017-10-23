<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p1201mb4_ko441.asp
'*  4. Program Name         : Routing Component Allocation
'*  5. Program Desc         :
'*  6. Component List       : PP1S506.cPMngCmpReqByRtng
'*  7. Modified date(First) : 2000/03/27
'*  8. Modified date(Last)  : 2008/01/31
'*  9. Modifier (First)     : Im Hyun Soo
'* 10. Modifier (Last)      : HAN cheol
'* 11. Comment              :
'**********************************************************************************************

On Error Resume Next

Call HideStatusWnd
Err.Clear

Call LoadBasisGlobalInf

Dim pPP1S506
Dim I1_plant_cd, I2_item_cd, I3_rout_no, I4_opr_no, iErrorPosition, strSpread
Dim arrColVal, arrRowVal       '20080131::hanc
Dim iDx


strSpread   = Request("txtSpread")
arrRowVal   = Split(strSpread, gRowSep)                                 '20080131::hanc¢Ð: Split Row    data
'arrRowVal   = Split(strSpread, "BB")                                 '20080131::hanc¢Ð: Split Row    data
I1_plant_cd	= Trim(UCase(Request("txtPlantCd")))
I2_item_cd	= Trim(UCase(Request("txtItemCd")))
I3_rout_no	= Trim(UCase(Request("txtRoutNo")))
I4_opr_no	= Trim(UCase(Request("hOprNo")))
lgLngMaxRow = Request("txtMaxRows")     '20080131::hanc                                        '¢Ð: Read Operation Mode (CRUD)


If I1_plant_cd = "" Then
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)            
	Response.End 
ElseIf I3_rout_no = "" Then
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)             
	Response.End 
ElseIf I4_opr_no = "" Then
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)       
	Response.End 
ElseIf I2_item_cd = "" Then
	Call DisplayMsgBox("700112", vbOKOnly, "", "", I_MKSCRIPT)              
	Response.End 
End If

    Call SubOpenDB(lgObjConn)                                                        '¢Ð: Make a DB Connection


    For iDx = 1 To lgLngMaxRow 
        arrColVal = Split(arrRowVal(iDx-1), gColSep)                                 '¢Ð: Split Column data iDx-1

        Select Case arrColVal(0)
            Case "C"
	                Call SubBizSaveMultiCreate(arrColVal)                            '¢Ð: Create
            Case "U"
                    Call SubBizSaveMultiUpdate(arrColVal)                            '¢Ð: Update
            Case "D"				
                    Call SubBizSaveMultiDelete(arrColVal)                            '¢Ð: Delete
        End Select

        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next


'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal)


    Dim lgStrSQL
    
    On Error Resume Next                                                             '¢Ð: Protect system from crashing
    Err.Clear                                                                        '¢Ð: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------
    lgStrSQL =              " UPDATE	P_ROUTING_DETAIL "
    lgStrSQL = lgStrSQL &   " SET		MACHINE_CD		= " & FilterVar(            arrColVal(3) ,"","S")  & ","
    lgStrSQL = lgStrSQL &   " 		    MACHINE_NM		= " & FilterVar(            arrColVal(4) ,"","S")  & ","
    lgStrSQL = lgStrSQL &   " 		    REWORK_YN		= " & FilterVar(            arrColVal(5) ,"","S")  & " "
    lgStrSQL = lgStrSQL &   " WHERE	    PLANT_CD        = " & FilterVar(I1_plant_cd,"","S")                   & " "      
    lgStrSQL = lgStrSQL &   " AND		ITEM_CD         = " & FilterVar(I2_item_cd,"","S")                    & " "      
    lgStrSQL = lgStrSQL &   " AND		ROUT_NO         = " & FilterVar(I3_rout_no,"","S")                    & " "      
    lgStrSQL = lgStrSQL &   " AND		OPR_NO          = " & FilterVar(I4_opr_no,"","S")                     & " "      


    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords


    If CheckSYSTEMError(Err,True) = True Then
       lgErrorStatus    = "YES"
       ObjectContext.SetAbort
    End If
End Sub

Response.Write "<Script Language=VBScript>" & vbCrLf
Response.Write "	parent.DbSaveOk" & vbCrLf
Response.Write "</Script>" & vbCrLf
Response.End 

    Call SubCloseDB(lgObjConn)                                                       '¢Ð: Close DB Connection

%>