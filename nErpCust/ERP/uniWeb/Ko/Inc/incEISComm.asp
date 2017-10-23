<%

Const  TABSTYLE01 = "CELLSPACING = 1 CELLPADDING = 1 width=100% border=0 STYLE = ""filter=progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr=#EAEADA, EndColorStr=WHITE)"""

Sub  PrintTitle(ByVal pTitle)

     Response.Write "<TABLE BORDER =0  WIDTH=100% CELLPADDING=0 CELLSPACING=1 STYLE = ""filter=progid:DXImageTransform.Microsoft.Gradient(GradientType=1,StartColorStr=white, EndColorStr=#EAEADA)"">" & vbCrLf
     Response.Write "    <TR>" & vbCrLf
     Response.Write "       <td width=61><img src=""../../../CShared/EISImage/Main/title_icon.gif""></td>      " & vbCrLf
     Response.Write "       <td width=150 valign=bottom ><b>" & pTitle & "</b></td>          " & vbCrLf
     Response.Write "	    <TD ALIGN=RIGHT VALIGN=BOTTOM>  " & vbCrLf
     Response.Write "          <table BORDER = 0 CELLPADDING=0 CELLSPACING=1>  " & vbCrLf
     Response.Write "              <tr>   " & vbCrLf
     Response.Write "                  <td></td>  " & vbCrLf
     Response.Write "                  <td><a href='vbscript:DoSpExec(""Preview"")' onMouseOut=""javascript:MM_swapImgRestore()"" onMouseOver=""javascript:MM_swapImage('Image10','','../../image/EIS/Button/bu_r_02.gif',1)""><img src=""../../image/EIS/Button/bu_02.gif"" name=Image10 width=61 height=22 border=0></a></td>  " & vbCrLf
     Response.Write "                  <td width=5></td>  " & vbCrLf
'     Response.Write "                  <td><a href='vbscript:DoSpExec(""Print"")'   onMouseOut=""javascript:MM_swapImgRestore()"" onMouseOver=""javascript:MM_swapImage('Image9' ,'','../../image/EIS/Button/bu_r_01.gif',1)""><img src=""../../image/EIS/Button/bu_01.gif"" name=Image9  width=61 height=22 border=0></a></td>  " & vbCrLf
'     Response.Write "                  <td width=5></td>  " & vbCrLf
'     Response.Write "                  <td><a href='vbscript:DoMagnify()' onMouseOut=""javascript:MM_swapImgRestore()"" onMouseOver=""javascript:MM_swapImage('Image11','','../../image/EIS/Button/bu_r_03.gif',1)""><img src=""../../image/EIS/Button/bu_03.gif"" name=Image11 width=61 height=22 border=0></a></td>  " & vbCrLf
'     Response.Write "                  <td width=5></td>   " & vbCrLf
'     Response.Write "                  <td><a href='vbscript:DoReduce()'  onMouseOut=""javascript:MM_swapImgRestore()"" onMouseOver=""javascript:MM_swapImage('Image12','','../../image/EIS/Button/bu_r_04.gif',1)""><img src=""../../image/EIS/Button/bu_04.gif"" name=Image12 width=61 height=22 border=0></a></td>  " & vbCrLf
     Response.Write "               </tr>  " & vbCrLf
     Response.Write "            </table>   " & vbCrLf
     Response.Write "       </TD>  " & vbCrLf
     Response.Write "     </TR>                " & vbCrLf   
     Response.Write "     <TR>  " & vbCrLf
     Response.Write "	   <td colspan=3 height=2 background=""../../../CShared/EISImage/Main/path_line.gif""></td>					  " & vbCrLf
     Response.Write "     </TR>  " & vbCrLf
     Response.Write "</TABLE>  " & vbCrLf

End Sub 

%>