<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="insa_report.aspx.cs" Inherits="ERPAppAddition.ERPAddition.INSA.insa_report"  EnableEventValidation="false"%>
<%@ Register Assembly="Microsoft.ReportViewer.WebForms, Version=11.0.0.0, Culture=neutral, PublicKeyToken=89845DCD8080CC91" Namespace="Microsoft.Reporting.WebForms" TagPrefix="rsweb" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
      <link rel="stylesheet" href="//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css"/>
      <script src="//code.jquery.com/jquery-1.10.2.js"></script>
      <script src="//code.jquery.com/ui/1.11.4/jquery-ui.js"></script>
      <link rel="stylesheet" href="/resources/demos/style.css"/>
      <style type="text/css">
          .custom-combobox {
              position: relative;
              display: inline-block;
          }

          .custom-combobox-toggle {
              position: absolute;
              top: 0;
              bottom: 0;
              margin-left: -1px;
              padding: 0;
          }

          .custom-combobox-input {
              margin: 0;
              padding: 5px 10px;
          }
        .title {
            font-family: 굴림체;
            font-size: 10pt;
            text-align: left;
            font-weight: bold;
            background-color: #EAEAEA;
            color: Blue;
            vertical-align: middle;
            display: table-cell;
            line-height: 25px;
            height: 25px;
        }

        .style22 {
            height: 30px;
            width: 61px;
        }

        .dt {
            font-family: 굴림체;
            font-size: 10pt;
            text-align: center;
            margin-left: 0px;
        }

        .style15 {
            height: 30px;
            width: 352px;
        }

        .style5 {
            height: 30px;
            width: 63px;
        }

        .style9 {
            height: 30px;
            width: 173px;
        }

        .style14 {
            height: 30px;
        }

        .style23 {
            width: 61px;
            height: 53px;
        }

        .style19 {
            width: 352px;
            height: 53px;
        }

        .style20 {
            width: 63px;
            height: 53px;
        }

        .style21 {
            width: 173px;
            height: 53px;
        }

        .style12 {
            width: 120px;
            background-color: #99CCFF;
            font-weight: bold;
            font-family: 굴림체;
            font-size: 10pt;
            text-align: center;
        }
          .auto-style1 {
              width: 148px;
          }
    </style>
    <script type ="text/javascript">
        $(document).ready(function () {

            $('#ButtonAdd').click(function (e) {
                var selectedOpts = $('#list_select option:selected');
                if (selectedOpts.length == 0) {
                    alert("Nothing to move.");
                    e.preventDefault();
                }
                $('#list_add').append($(selectedOpts).clone());
                $(selectedOpts).remove();
                e.preventDefault();
            });

            $('#ButtonRemove').click(function (e) {
                var selectedOpts = $('#list_add option:selected');
                if (selectedOpts.length == 0) {
                    alert("Nothing to move.");
                    e.preventDefault();
                }
                $('#list_select').append($(selectedOpts).clone());
                $(selectedOpts).remove();
                e.preventDefault();
            });

            $('#ButtonAddAll').click(function (e) {
                var selectedOpts = $('#list_select option');
                $('#list_add').append($(selectedOpts).clone());
                $(selectedOpts).remove();
                e.preventDefault();
            });

            $('#ButtonRemoveAll').click(function (e) {
                var selectedOpts = $('#list_add option');
                $('#list_select').append($(selectedOpts).clone());
                $(selectedOpts).remove();
                e.preventDefault();
            });

            $('#btn_insert').bind("click", function () {
                $('#list_add option').attr("selected", "selected");
            });

            $("#combobox").combobox({
                select: function (event, ui) {
                    $('#btn_hidden').trigger('click');
                }
            });
            $('#bt_retrieve').click(function (e) {

                if ($("#tb_yyyymm").val() == "") {
                    alert("기준일을 입력해주세요.");
                    e.preventDefault();
                }
            });
          
        });

    </script>
    <script type ="text/javascript">
        (function ($) {
            $.widget("custom.combobox", {
                _create: function () {
                    this.wrapper = $("<span>")
                      .addClass("custom-combobox")
                      .insertAfter(this.element);

                    this.element.hide();
                    this._createAutocomplete();
                    this._createShowAllButton();
                },

                _createAutocomplete: function () {
                    var selected = this.element.children(":selected"),
                      value = selected.val() ? selected.text() : "";

                    this.input = $("<input>")
                      .appendTo(this.wrapper)
                      .val(value)
                      .attr("title", "")
                      .addClass("custom-combobox-input ui-widget ui-widget-content ui-state-default ui-corner-left")
                      .autocomplete({
                          delay: 0,
                          minLength: 0,
                          source: $.proxy(this, "_source")
                      })
                      .tooltip({
                          tooltipClass: "ui-state-highlight"
                      });

                    this._on(this.input, {
                        autocompleteselect: function (event, ui) {
                            ui.item.option.selected = true;
                            this._trigger("select", event, {
                                item: ui.item.option
                            });
                        },

                        autocompletechange: "_removeIfInvalid"
                    });
                },

                _createShowAllButton: function () {
                    var input = this.input,
                      wasOpen = false;

                    $("<a>")
                      .attr("tabIndex", -1)
                      .attr("title", "Show All Items")
                      .tooltip()
                      .appendTo(this.wrapper)
                      .button({
                          icons: {
                              primary: "ui-icon-triangle-1-s"
                          },
                          text: false
                      })
                      .removeClass("ui-corner-all")
                      .addClass("custom-combobox-toggle ui-corner-right")
                      .mousedown(function () {
                          wasOpen = input.autocomplete("widget").is(":visible");
                      })
                      .click(function () {
                          input.focus();

                          // Close if already visible
                          if (wasOpen) {
                              return;
                          }

                          // Pass empty string as value to search for, displaying all results
                          input.autocomplete("search", "");
                      });
                },

                _source: function (request, response) {
                    var matcher = new RegExp($.ui.autocomplete.escapeRegex(request.term), "i");
                    response(this.element.children("option").map(function () {
                        var text = $(this).text();
                        if (this.value && (!request.term || matcher.test(text)))
                            return {
                                label: text,
                                value: text,
                                option: this
                            };
                    }));
                },

                _removeIfInvalid: function (event, ui) {

                    // Selected an item, nothing to do
                    if (ui.item) {
                        return;
                    }

                    // Search for a match (case-insensitive)
                    var value = this.input.val(),
                      valueLowerCase = value.toLowerCase(),
                      valid = false;
                    this.element.children("option").each(function () {
                        if ($(this).text().toLowerCase() === valueLowerCase) {
                            this.selected = valid = true;
                            return false;
                        }
                    });

                    // Found a match, nothing to do
                    if (valid) {
                        return;
                    }

                    // Remove invalid value
                    this.input
                      .val("")
                      .attr("title", value + " didn't match any item")
                      .tooltip("open");
                    this.element.val("");
                    this._delay(function () {
                        this.input.tooltip("close").attr("title", "");
                    }, 2500);
                    this.input.autocomplete("instance").term = "";
                },

                _destroy: function () {
                    this.wrapper.remove();
                    this.element.show();
                }
            });
        })(jQuery);
    </script>
</head>
<body>
    <form id="form2" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>
        <div>
            <table>
                <tr>
                    <td>
                        <asp:Image ID="Image1" runat="server" ImageUrl="~/img/folder.gif" />
                    </td>
                    <td style="width: 100%;">
                        <asp:Label ID="Label2" runat="server" Text="실시간인원현황" CssClass="title" Width="100%"></asp:Label></td>
                </tr>
            </table>
        </div>
        ▶ 총인원 : 해당 사업장의 재직 인원
        <br />
        <br />
        ▶ 현재인원 : 해당 사업장에서 ID카드, 경비실 근태로 출근체크를 한 인원(퇴근 시 제외)
        <br />
        <br />
        <table style="border: thin solid #000080; height: 31px;">
            <tr>
                <td class="style12">
                    <asp:Label ID="Label17" runat="server" Text="조회구분" BackColor="#99CCFF"
                        Font-Bold="True" Style="text-align: center; font-size: small"></asp:Label>
                </td>
                <td class="style3">
                    <asp:RadioButtonList ID="rbl_view_type" runat="server" Font-Size="Small"
                        RepeatDirection="Horizontal"
                        AutoPostBack="True" Width="234px" Style="margin-left: 0px; font-weight: 700;"
                        BackColor="White" Height="21px">
                        <asp:ListItem Value="A" Selected="True">조 회</asp:ListItem>
                        <asp:ListItem Value="B">기준정보</asp:ListItem>
                    </asp:RadioButtonList>
                </td>
            </tr>
        </table>
        <br />
        <div id ="search_table" runat ="server" visible ="false">
            <table style="border: thin solid #000080">
                <tr>
                    <td class="style12">
                        <asp:Label ID="lb_yyyymm" runat="server" Text="기 준 일"></asp:Label>
                    </td>
                    <td class="auto-style1">
                        <asp:TextBox ID="tb_yyyymm" runat="server" MaxLength="8" BackColor ="Yellow" ClientIDMode ="Static"></asp:TextBox>
                    <td>
                        EX) 20150701 (YYYYMMDD)
                    </td>
                    </td>
                    <td>
                        <asp:Button ID="bt_retrieve" runat="server" OnClick="bt_retrieve_Click" Text="조 회" Width="100px" ClientIDMode ="Static"/></td>
                </tr>
            </table>
        </div>
        <br />
        <div id ="chk_standard" runat ="server" visible ="false">
            <table>
                <tr>
                    <td>
                        성 명 :  <select id="combobox" runat ="server" datasourceid ="SqlDataSource2" datatextfield ="pr_name" datavaluefield ="pr_sno"></select>
                        <asp:SqlDataSource ID="SqlDataSource2" runat="server" ConnectionString="<%$ ConnectionStrings:nepes_test1 %>" 
                            SelectCommand="select pr_name,pr_sno from INSADB.INBUS.DBO.h_person where pr_outdate = '' and pr_notdp = '0'"></asp:SqlDataSource>
                    </td>
                    <td><asp:Button ID="btn_insert" runat="server" OnClick="bt_insert_click" Text="입 력" Width="100px" ClientIDMode ="Static"/></td>
                    <td><asp:Button ID="btn_hidden" runat="server" OnClick="btn_hidden_Click" Text="숨 김" Width="100px" ClientIDMode ="Static"/></td>
                </tr>
                <tr> 
                    <td>
                        <asp:ListBox ID="list_select" runat="server" SelectionMode="Multiple" Width="500px" Height="500px" DataSourceID="SqlDataSource1" DataTextField="bu_name" DataValueField="bu_code" ClientIDMode ="Static"></asp:ListBox>
                        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:nepes_test1 %>" SelectCommand="select bu_name,bu_code from INSADB.INBUS.DBO.c_busor WHERE bu_edate = ''"></asp:SqlDataSource>
                    </td>
                    <td valign="middle" align="center" style="width: 100px">
                        <asp:Button ID="ButtonAdd" runat="server" Text=">"  Width="50px" ClientIDMode ="Static"/>
                        <br />
                        <asp:Button ID="ButtonRemove" runat="server" Text="<" Width="50px" ClientIDMode ="Static" />
                        <br />
                        <asp:Button ID="ButtonAddAll" runat="server" Text=">>>"  Width="50px" ClientIDMode ="Static"/>
                        <br />
                        <asp:Button ID="ButtonRemoveAll" runat="server" Text="<<<"  Width="50px" ClientIDMode ="Static"/>
                    </td>
                    <td>
                        <asp:ListBox ID="list_add" runat="server" SelectionMode="Multiple" Width="500px" Height="500px" ClientIDMode ="Static"></asp:ListBox>
                    </td>
                </tr>
            </table>
        </div>
        <br />
        <script type="text/javascript">
            var ModalProgress = '<%= ModalProgress.ClientID %>';

            Sys.WebForms.PageRequestManager.getInstance().add_beginRequest(beginReq);
            Sys.WebForms.PageRequestManager.getInstance().add_endRequest(endReq);
            function beginReq(sender, args) {
                //show the Popup
                $find(ModalProgress).show()
            }
            function endReq(sender, args) {
                //hide the Popup
                $find(ModalProgress).hide();
            }
        </script>
        <asp:UpdatePanel ID="UpdatePanel2" runat="server">
            <Triggers>
                <asp:AsyncPostBackTrigger ControlID="bt_retrieve" />
            </Triggers>
        </asp:UpdatePanel>
        <asp:UpdateProgress ID="UpdateProg1" runat="server" DisplayAfter="200">
            <ProgressTemplate>
                <asp:Image ID="Image3_1" runat="server" CssClass="updateProgress" ImageUrl="~/img/loading9_mod.gif" />
                <br />
                <br />
                <br />
                <br />
                <asp:Image ID="Image2_1" runat="server" ImageUrl="~/img/ajax-loader.gif" />
            </ProgressTemplate>
        </asp:UpdateProgress>
        <cc1:ModalPopupExtender ID="ModalProgress" runat="server" PopupControlID="UpdateProg1" TargetControlID="UpdateProg1">
        </cc1:ModalPopupExtender>
        <rsweb:ReportViewer ID="ReportViewer1" runat="server" AsyncRendering="False" Height="200px" SizeToReportContent="True" WaitControlDisplayAfter="600000" Width="934px">
        </rsweb:ReportViewer>
        <rsweb:ReportViewer ID="ReportViewer2" runat="server" AsyncRendering="False" Height="430px" SizeToReportContent="True" WaitControlDisplayAfter="600000" Width="934px">
        </rsweb:ReportViewer>
    </form>
</body>
</html>
