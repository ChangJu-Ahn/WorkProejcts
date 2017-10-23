<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : BA
'*  2. Function Name        : System Management
'*  3. Program ID           : za009mb2
'*  4. Program Name         : Audit Info Overview Query
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000.03.13
'*  8. Modified date(Last)  : 2002.06.21
'*  9. Modifier (First)     : LeeJaeJoon
'* 10. Modifier (Last)      : LeeJaeWan
'* 11. Comment              :
'**********************************************************************************************-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>


<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">        

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>

<%
Call HideStatusWnd
Call LoadBasisGlobalInf

On Error Resume Next
Err.Clear


Dim iDx

Dim iZa023
Dim I2_z_audit_mast_dtl
Dim E2_z_audit_mast_dtl
Dim E3_bk_column

Const ZA23_I2_table_id = 0
Const ZA23_I2_date = 1
Const ZA23_I2_time = 2
Const ZA23_I2_tran_id = 3
Const ZA23_I2_usr_id = 4

If Request("txtTable") <> "" Then

    ReDim I2_z_audit_mast_dtl(ZA23_I2_usr_id)
    I2_z_audit_mast_dtl(ZA23_I2_table_id) = Request("txtTable")
    I2_z_audit_mast_dtl(ZA23_I2_date) = Request("txtOccurDt")
    I2_z_audit_mast_dtl(ZA23_I2_time) = Request("txtOccurTm")
    I2_z_audit_mast_dtl(ZA23_I2_tran_id) = Request("txtTransCd")
    I2_z_audit_mast_dtl(ZA23_I2_usr_id) = Request("txtUser")

    Set iZa023 = Server.CreateObject("PZAG023.cListAuditMast")
    Call iZa023.ZA_Lookup_Audit_Mast_Dtl(gStrGlobalCollection, I2_z_audit_mast_dtl, E2_z_audit_mast_dtl, E3_bk_column)

    If CheckSYSTEMError(Err,True) = True Then
       Set iZa023 = Nothing                                                         
    End If
    
End If
%>
<SCRIPT LANGUAGE=VBSCRIPT>
Sub Window_onLoad()
    Dim iDx
    Dim iEnvInf

    gMouseClickStatus = "N"

    gTabMaxCnt = 0
    gIsTab     = "N"
    gPageNo    =  1
    
    Call parent.Parent.AdjustStyleSheet(Document)

    iEnvInf =           parent.parent.gADODBConnString   & Chr(12)   '0
    iEnvInf = iEnvInf & parent.parent.gLang              & Chr(12)   '1
    iEnvInf = iEnvInf & gStrRequestMenuID         & Chr(12)   '2
    iEnvInf = iEnvInf & parent.parent.gUsrId             & Chr(12)   '3
    iEnvInf = iEnvInf & parent.parent.gClientNm          & Chr(12)   '4
    iEnvInf = iEnvInf & parent.parent.gClientIp          & Chr(12)   '5
    iEnvInf = iEnvInf & parent.parent.gUsrId             & Chr(12)   '6
    iEnvInf = iEnvInf & parent.parent.gSeverity          & Chr(12)   '7

    parent.Parent.gStrRequestMenuID = gStrRequestMenuID
    parent.Parent.gEnvInf           = iEnvInf
    
    Call Form_Load()
' KIT
    If Trim(gStrRequestUpperMenuID) <> "" Then
       top.document.title = parent.parent.gLogo & " - " & "[" & gStrRequestUpperMenuID & "][" & gStrRequestMenuID & "][" & document.title & "]"  
    Else
       top.document.title = parent.parent.gLogo & " - " & "[" & document.title & "]"  
    End If
' KIT    
    window.status      = ""

    Set gActiveElement        = document.activeElement 
    gLookUpEnable      = True    

    'iDx = Instr(UCase(document.location.href),"MODULE")

    'If iDx > 0 And Trim(gStrRequestMenuID) > ""  Then   ' This means that if current program id is not popup    
    '   Document.Cookie = "gActivePgmID" & "=" & Mid(document.location.href,iDx ) & "; path=" & "/"    
    'End If       
    
End Sub
'=========================================================================================================
Sub Form_Load()
    Call Parent.ggoOper.LockField(Document, "N") 
End Sub
</SCRIPT>

<!-- #Include file="../../inc/uni2kcm.inc" -->    

</HEAD>
<BODY TABINDEX="-1" SCROLL=Auto>
<FORM NAME="frm1">
<TABLE CLASS="BasicTB">
    <TR>
        <TD CLASS="TD5" STYLE="TEXT-ALIGN:Center" HEIGHT=20>Value</TD>
    </TR>
    <TR HEIGHT=10>
        <TD>
            <TABLE CLASS="BasicTB" CELLSPACING=0>
                <TR>
                    <TD CLASS="TD5" STYLE="WIDTH:43%; HEIGHT:20">발생날짜</TD>
                    <TD CLASS="TD6" STYLE="WIDTH:57%"><INPUT TYPE=TEXT NAME="txtOccurDT" SIZE=22 tag="14X" VALUE="<%=UNIDateClientFormat(E2_z_audit_mast_dtl(0,0))%>"></TD>
                </TR>
                <TR>
                    <TD CLASS="TD5" HEIGHT=20>발생시간</TD>
                    <TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtOccurTm" SIZE=22 tag="14X" VALUE="<%=E2_z_audit_mast_dtl(1,0)%>"></TD>
                </TR>
                <TR>
                    <TD CLASS="TD5" HEIGHT=20>트랜잭션</TD>
                    <TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtTrans" SIZE=22 tag="14X" VALUE="<%=E2_z_audit_mast_dtl(UBound(E2_z_audit_mast_dtl,1),0)%>"></TD>
                </TR>
                <TR>
                    <TD CLASS="TD5" HEIGHT=20>사용자</TD>
                    <TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtUsr" SIZE=22 tag="14X" VALUE="<%=Replace(E2_z_audit_mast_dtl(4,0),"""","&quot;")%>"></TD>
                </TR>
                <TR>
                    <TD CLASS="TD5" HEIGHT=20>테이블</TD>
                    <TD CLASS="TD6"><INPUT TYPE=TEXT NAME="txtTable" SIZE=22 tag="14X" VALUE="<%=Replace(I2_z_audit_mast_dtl(ZA23_I2_table_id),"""","&quot;")%>"></TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
    <TR HEIGHT=*>
        <TD>
            <TABLE CLASS="BasicTB" CELLSPACING=0>
<%
If IsEmpty(E3_bk_column) Then
    Response.Write "<TR>" & vbCr
    Response.Write "<TD CLASS=""TD5"" STYLE=""WIDTH:43%; HEIGHT:20""></TD>" & vbCr
    Response.Write "<TD CLASS=""TD6"" STYLE=""WIDTH:57%""></TD>" & vbCr
    Response.Write "</TR>" & vbCr
    
    Response.End
End If

For iDx = 5 To UBound(E3_bk_column,1) - 1
    Response.Write "<TR>" & vbCr
    Response.Write "<TD CLASS=""TD5"" STYLE=""WIDTH:43%; HEIGHT:20"">" & E3_bk_column(iDx) & "</TD>" & vbCr
    If IsDate(E2_z_audit_mast_dtl(iDx, 0)) Then
    Response.Write "<TD CLASS=""TD6"" STYLE=""WIDTH:57%""><INPUT TYPE=TEXT NAME=""arrCond"" SIZE=22 tag=""14X"" VALUE=""" & UNIDateClientFormat(E2_z_audit_mast_dtl(iDx,0)) & """></TD>" & vbCr
    ElseIf IsNULL(E2_z_audit_mast_dtl(iDx, 0)) Then
    Response.Write "<TD CLASS=""TD6"" STYLE=""WIDTH:57%""><INPUT TYPE=TEXT NAME=""arrCond"" SIZE=22 tag=""14X"" VALUE=""""></TD>" & vbCr
    Else
    Response.Write "<TD CLASS=""TD6"" STYLE=""WIDTH:57%""><INPUT TYPE=TEXT NAME=""arrCond"" SIZE=22 tag=""14X"" VALUE=""" & Replace(E2_z_audit_mast_dtl(iDx,0),"""","&quot;") & """></TD>" & vbCr
    End If
    Response.Write "</TR>" & vbCr
Next
%>
                <TR>
                    <TD CLASS="TD5" STYLE="WIDTH:43%" HEIGHT=*></TD>
                    <TD CLASS="TD6" STYLE="WIDTH:57%"></TD>
                </TR>
            </TABLE>
        </TD>
    </TR>
</TABLE>
</FORM>
</BODY>
</HTML>


