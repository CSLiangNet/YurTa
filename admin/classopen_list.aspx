<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" EnableEventValidation="false" Debug="true" %>

<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>
<% Session.Contents.Remove("se_class_id")%>

<%--<asp:sqldatasource id="Da_classopen" runat="server"
    connectionstring="<%$ ConnectionStrings:yurta_DBConnectionString %>"
    selectcommand="SELECT class_id, class_name, class_style, class_allowno, CONVERT(char(10), class_opendate,111)as class_opendate, CONVERT(char(10), class_enddate,111)as class_enddate, class_opend, CONVERT(char(10), class_endsign,111)as class_endsign, class_publish, class_counter, (SELECT COUNT(*) FROM dbo.yurta_sign where dbo.yurta_classopen.class_id = dbo.yurta_sign.sign_flowid AND dbo.yurta_sign.sign_cancel = @sign_cancel) AS NO_count FROM dbo.yurta_classopen where class_style=@class_style ORDER BY class_opendate DESC">
            <SelectParameters>
                 <asp:Parameter DefaultValue="初訓" Name="class_style" Type="String" />
                <asp:Parameter DefaultValue="0" Name="sign_cancel" Type="String" />
            </SelectParameters>
        </asp:sqldatasource>--%>



<%--<asp:sqldatasource id="Da_web" runat="server"
    connectionstring="<%$ ConnectionStrings:yurta_DBConnectionString %>"
    selectcommand="SELECT *  FROM dbo.yurta_ser  WHERE ser_item = @ser_item AND ser_enable = @ser_enable">
            <SelectParameters>
                 <asp:Parameter DefaultValue="16" Name="ser_item" Type="String" />
                <asp:Parameter DefaultValue="1" Name="ser_enable" Type="String" />
            </SelectParameters>
        </asp:sqldatasource>--%>


<MM:DataSet
    ID="Da_classopen"
    runat="Server"
    IsStoredProcedure="false"
    ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_conn_iia") %>'
    DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_conn_iia") %>'
    CommandText='<%# "SELECT class_id, class_name, class_style, class_allowno, CONVERT(char(10), class_opendate,111)as class_opendate, CONVERT(char(10), class_enddate,111)as class_enddate, class_opend, CONVERT(char(10), class_endsign,111)as class_endsign, class_publish, class_counter, (SELECT COUNT(*) FROM dbo.yurta_sign where dbo.yurta_classopen.class_id = dbo.yurta_sign.sign_flowid AND dbo.yurta_sign.sign_cancel = @sign_cancel) AS NO_count FROM dbo.yurta_classopen where class_style=@class_style ORDER BY class_opendate DESC" %>'
    Debug="true">
    <parameters> 
    <Parameter  Name="@class_style"  Value='<%# "初訓" %>'  Type="VarChar"   />    <Parameter  Name="@sign_cancel"  Value='<%# 0 %>'  Type="NChar"   /> 
  </parameters>
</MM:DataSet>
<MM:DataSet
    ID="Da_web"
    runat="Server"
    IsStoredProcedure="false"
    ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_conn_iia") %>'
    DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_conn_iia") %>'
    CommandText='<%# "SELECT *  FROM dbo.yurta_ser  WHERE ser_item = @ser_item AND ser_enable = @ser_enable" %>'
    Debug="true">
    <parameters>
    <Parameter  Name="@ser_item"  Value='<%# "16" %>'  Type="NChar"   />
    <Parameter  Name="@ser_enable"  Value='<%# "1" %>'  Type="SmallInt"   />
  </parameters>
</MM:DataSet>

<MM:PageBind runat="server" PostBackBind="true" />
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <script runat="server">
        Sub Page_Load(Src As Object, E As EventArgs)
            '若PS是Nothign就不執行本頁
            Session("IIA_userac") = "OK"
            If Session("IIA_userac") = Nothing Then
                '移除目前的 Session
                Session.Abandon()
                Response.Redirect("login.aspx?erry=timeout")
            End If
            Session("IIA_syD") = "OK"
            If Session("IIA_syD") = Nothing Then
                '沒有IIA_syD的管理權限
                Response.Redirect("admin.aspx?erry=0")
            End If
            If Not IsPostBack Then
                DataGrid1.DataBind()
            End If
            '	session.Contents.Remove("Usign_id")
            '	session.Contents.Remove("Uclass_opendate")
            '	session.Contents.Remove("Uclass_enddate")
            '	session.Contents.Remove("Uclass_style")
        End Sub
    </script>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>祐太技術顧問股份有限公司</title>
    <style type="text/css">
        <!--
        body {
            margin-left: 0px;
            margin-top: 0px;
            margin-right: 0px;
            margin-bottom: 0px;
        }
        -->
    </style>
    <link rel="stylesheet" type="text/css" href="style.css">

    <script src="SpryAssets/SpryMenuBar.js" type="text/javascript"></script>
    <link href="SpryAssets/SpryMenuBarHorizontal.css" rel="stylesheet" type="text/css" />
    <style type="text/css">
        <!--
        .style10 {
            font-size: large;
            color: #003399;
            font-weight: bold;
        }
        -->
    </style>
    <script language="javascript">
        function ConfirmDeletion() {
            if (event.srcElement.value == "刪除")
                event.returnValue = confirm("確認要刪除嗎？");
        } document.onclick = ConfirmDeletion;
        function GP_popupConfirmMsg(msg) { //v1.0
            document.MM_returnValue = confirm(msg);
        }
    </script>
</head>
<body>
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr>
            <td bgcolor="#FFFFFF">&nbsp;</td>
            <td bgcolor="#014AA6">&nbsp;</td>
            <td bgcolor="#014AA6">
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr>

                        <td width="14" height="32" valign="bottom" bgcolor="#014AA6"></td>
                        <td width="100%" valign="top" bgcolor="#014AA6">
                            <!-- #include file="menu_top.aspx" -->
                        </td>
                    </tr>
                </table>
            </td>
            <td>
                <img src="images/home-5u.jpg" width="15" height="32" /></td>
        </tr>
        <tr>
            <td bgcolor="#FFFFFF">&nbsp;</td>
            <td width="11" height="11">
                <img src="images/test_11.gif" width="11" height="16" /></td>
            <td height="16">
                <img src="images/test_02.gif" width="100%" height="16" /></td>
            <td width="11" height="11">
                <img src="images/test_10.gif" width="11" height="16" /></td>
        </tr>
        <tr>
            <td width="6" bgcolor="#FFFFFF">&nbsp;</td>
            <td width="5" background="images/login_lf.gif">&nbsp;</td>
            <td valign="top">
                <div align="center">
                    <br />
                    <br />
                    <table width="98%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="11" height="11">
                                <img src="images/test_01.gif" width="11" height="11" /></td>
                            <td height="11" colspan="2">
                                <img src="images/test_02.gif" width="100%" height="11" /></td>
                            <td width="11" height="11">
                                <img src="images/test_03.gif" width="11" height="11" /></td>
                        </tr>
                        <tr>
                            <td width="11" valign="top">
                                <img src="images/test_04.gif" width="11" height="50" /></td>
                            <td colspan="2" bgcolor="#ffffff">
                                <table width="100%" border="0" cellpadding="0" cellspacing="0">
                                    <tr>
                                        <td width="48">
                                            <img src="images/png-0014-s.png" alt="開課管理" width="48" height="49" border="0" /></td>
                                        <td width="80%">
                                            <div align="left" class="itemtitle">
                                                初訓開課管理 		
                                            </div>
                                        </td>
                                        <td width="50"><% If Session("IIA_syD") > "1" Then%><div align="center">
                                            <a href="classopen_add_1.aspx">
                                                <img src="images/png-0084.png" alt="開課" width="32" height="32" border="0" /></a><br />
                                            <a href="classopen_add_1.aspx"><span class="en1">開課</span></a>
                                        </div>
                                            <% End If%></td>
                                        <td width="66">
                                            <div align="center">
                                                <a href="admin.aspx">
                                                    <img src="images/png-0097.png" alt="回總管理" width="32" height="32" border="0" longdesc="admin.aspx" /></a><br />
                                                <a href="admin.aspx"><span class="en1">回總管理</span></a>
                                            </div>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td width="11" height="38" valign="top">
                                <img src="images/test_06.gif" width="11" height="50" /></td>
                        </tr>

                        <tr>
                            <td width="11" height="11" valign="top">
                                <img src="images/test_20.gif" width="11" height="11" /></td>
                            <td height="11" colspan="2" valign="top">
                                <img src="images/test_21.gif" width="100%" height="11" /></td>
                            <td width="11" height="11" valign="top">
                                <img src="images/test_22.gif" width="11" height="11" /></td>
                        </tr>
                    </table>

                    <% If Request.QueryString("sendok") = "ok" Then%> 開課課程通知,寄送完成! <% End If%>
                    <% If Request.QueryString("sendok") = "ok_cancel" Then%> 取消課程通知,寄送完成! <% End If%>
                    <% If Request.QueryString("sendok") = "ok_change" Then%> 異動課程通知,寄送完成! <% End If%>
                    <% If Request.QueryString("sendok") = "bad" Then%> 寄送,失敗!沒有相關的郵寄名單 <% End If%>
                    <% If Request.QueryString("sendok_chech") = "ok" Then%> 修改課程起迄日<% '= Request.QueryString("c_open") %>完成！ <% End If%>
                    <br />
                    <br />
                    <table width="98%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                            <td width="11" height="11">
                                <img src="images/test_01.gif" width="11" height="11" /></td>
                            <td height="11" colspan="2">
                                <img src="images/test_02.gif" width="100%" height="11" /></td>
                            <td width="11" height="11">
                                <img src="images/test_03.gif" width="11" height="11" /></td>
                        </tr>
                        <tr>
                            <td width="11" valign="top" background="images/test_04.gif">&nbsp;</td>
                            <td colspan="2" valign="top" bgcolor="#ffffff">
                                <div align="center">
                                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                        <tr>
                                            <td>
                                                <div align="right"></div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div align="center">
                                                    <form runat="server">
                                                        <asp:DataGrid
                                                            AllowPaging="false"
                                                            AllowSorting="false"
                                                            AutoGenerateColumns="false"
                                                            CellPadding="3"
                                                            CellSpacing="0"
                                                            DataSource="<%# Da_classopen.DefaultView %>" ID="DataGrid1"
                                                            runat="server"
                                                            ShowFooter="false"
                                                            ShowHeader="true" Width="100%">
                                                            <HeaderStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
                                                            <ItemStyle BackColor="#F2F2F2" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
                                                            <AlternatingItemStyle BackColor="#E5E5E5" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
                                                            <FooterStyle HorizontalAlign="center" BackColor="#E8EBFD" ForeColor="#3D3DB6" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Bold="true" Font-Size="smaller" />
                                                            <PagerStyle BackColor="white" Font-Name="Verdana, Arial, Helvetica, sans-serif" Font-Size="smaller" />
                                                            <Columns>
                                                                <asp:TemplateColumn HeaderText="開課日期"
                                                                    Visible="True">
                                                                    <ItemTemplate><%# Da_classopen.FieldValue("class_opendate", Container)%></ItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:TemplateColumn HeaderText="課程名稱"
                                                                    Visible="True">
                                                                    <ItemTemplate>
                                                                        <a href="classopen_print.aspx?class_id=<%# Da_classopen.FieldValue("class_id", Container)%>" target="_blank">
                                                                            <div align="left">
                                                                                <%# Da_classopen.FieldValue("class_name", Container)%>
                                                                                <MM:If runat="server" Expression='<%# CDate(Trim(Da_classopen.FieldValue("class_endsign", Container))).AddDays(1).Tostring("yyyyMMdd") > DateTime.Now.ToString("yyyyMMdd") %>'>
                                                                                    <contentstemplate>                           
<MM:If runat="server" Expression='<%# CDate(Trim(Da_classopen.FieldValue("class_endsign", Container))).AddDays(-3).Tostring("yyyyMMdd") < DateTime.Now.ToString("yyyyMMdd") %>'>
                              <ContentsTemplate>
<img src="images/cutid-1.gif" width="30" height="30" title="報名截止日:<%# CDate(Trim(Da_classopen.FieldValue("class_endsign", Container))).Tostring("yyyy/MM/dd") %>"/>                              
                              </ContentsTemplate>
                            </MM:If> 
                            </contentstemplate>
                                                                                </MM:If>
                                                                            </div>
                                                                        </a>
                                                                    </ItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:TemplateColumn HeaderText="剩餘<br>名額"
                                                                    Visible="True">
                                                                    <ItemTemplate>
                                                                        <%# IIf((Int(Da_classopen.FieldValue("class_allowno", Container)) - Int(Da_classopen.FieldValue("NO_count", Container)) >= 0 ), Int(Da_classopen.FieldValue("class_allowno", Container)) - Int(Da_classopen.FieldValue("NO_count", Container)), "0") %>
                                                                    </ItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:TemplateColumn HeaderText="訓練<br>類別"
                                                                    Visible="True">
                                                                    <ItemTemplate>
                                                                        <MM:If runat="server" Expression='<%# Da_classopen.FieldValue("class_style", Container) = "初訓" %>'>
                                                                            <contentstemplate><div align="center"><%# Da_classopen.FieldValue("class_style", Container)%></div></contentstemplate>
                                                                        </MM:If>
                                                                        <MM:If runat="server" Expression='<%# Da_classopen.FieldValue("class_style", Container) = "複訓" %>'>
                                                                            <contentstemplate><div align="center"><B><%# Da_classopen.FieldValue("class_style", Container)%></B></div></contentstemplate>
                                                                        </MM:If>
                                                                    </ItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:BoundColumn DataField="class_publish"
                                                                    HeaderText="發佈否"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="NO_count"
                                                                    HeaderText="報名<br>人數"
                                                                    ReadOnly="true"
                                                                    Visible="True" />

                                                                <asp:BoundColumn DataField="class_opend"
                                                                    HeaderText="是否<br>開課"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="class_counter"
                                                                    HeaderText="點閱"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:TemplateColumn HeaderText="報名<br>管理"
                                                                    Visible="True">
                                                                    <ItemTemplate>
                                                                        <a href="classsign_list.aspx?class_id=<%# Da_classopen.FieldValue("class_id", Container)%>">
                                                                            <img src="images/class_sign.png" alt="報名管理" width="24" height="24" border="0" /></a>
                                                                        <a href="classsign_mail.aspx?class_id=<%# Da_classopen.FieldValue("class_id", Container)%>">
                                                                            <img src="images/mail.jpg" alt="Mail發送" width="24" height="24" border="0" />
                                                                    </ItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:TemplateColumn HeaderText="匯出作業"
                                                                    Visible="True">
                                                                    <ItemTemplate><a href="classsign_list_excel.aspx?class_id=<%# Da_classopen.FieldValue("class_id", Container)%>">報名表</a>&nbsp;&nbsp;<a href="classsign_list_excelfee.aspx?class_id=<%# Da_classopen.FieldValue("class_id", Container)%>">已結業</a></ItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:TemplateColumn HeaderText="更新<br>日期"
                                                                    Visible="True">
                                                                    <ItemTemplate>
                                                                        <a href="classsign_list_changeid.aspx?c_id=<%# Da_classopen.FieldValue("class_id", Container)%>&c_open=<%# Da_classopen.FieldValue("class_opendate", Container) %>&c_end=<%# Da_classopen.FieldValue("class_enddate", Container) %>&rs:ParameterLanguage=zh-TW">
                                                                            <img src="images/alarm-1.png" alt="修改報名人員的課程日期" width="24" height="24" border="0" onclick="GP_popupConfirmMsg('確定要修改報名人員的課程日期為<%# Trim(Da_classopen.FieldValue("class_opendate", Container)) %>?');return document.MM_returnValue" /></a>
                                                                    </ItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:TemplateColumn HeaderText="編輯"
                                                                    Visible="True">
                                                                    <ItemTemplate>
                                                                        <MM:If runat="server" Expression='<%# CDate(Trim(Da_classopen.FieldValue("class_opendate", Container))).AddDays(2).Tostring("yyyy/MM/dd") >= DateTime.Now.ToString("yyyy/MM/dd") %>'>
                                                                            <contentstemplate>
                                <% 'IIA_syD的權限劃分 %>
                                <% If Session("IIA_syD") = "1" Then%>
無編輯權限
<% End If%>
<% If Session("IIA_syD") > "1" Then%>
<a href="classopen_edit.aspx?class_id=<%# Da_classopen.FieldValue("class_id", Container) %>"><img src="images/edit.png" alt="修改" width="24" height="24" border="0"  /> </a>
<% End If%>
                              </contentstemplate>
                                                                        </MM:If>
                                                                        <MM:If runat="server" Expression='<%# CDate(Trim(Da_classopen.FieldValue("class_opendate", Container))).AddDays(2).Tostring("yyyy/MM/dd") < DateTime.Now.ToString("yyyy/MM/dd") %>'>
                                                                            <contentstemplate><% If Session("IIA_syD") > "1" Then%>
<a href="classopen_edit.aspx?class_id=<%# Da_classopen.FieldValue("class_id", Container) %>"><font color="#990000">過期</font></a>
<% End If%></contentstemplate>
                                                                        </MM:If>
                                                                    </ItemTemplate>
                                                                </asp:TemplateColumn>
                                                                <asp:TemplateColumn HeaderText="刪除"
                                                                    Visible="True">
                                                                    <ItemTemplate>
                                                                        <!-- 加防護:若此課程有報名人數,不執行 -->
                                                                        <MM:If runat="server" Expression='<%# Da_classopen.FieldValue("class_opendate", Container) >= DateTime.Now.ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) %>'>
                                                                            <contentstemplate>
                                <% If Session("IIA_syD") > "2" Then%>
                                <a href="classopen_del.aspx?class_id=<%# Da_classopen.FieldValue("class_id", Container)%>"><img src="images/del.png" alt="刪除" width="24" height="24" border="0" onclick="GP_popupConfirmMsg('確定要刪除<%# Trim(Da_classopen.FieldValue("class_opendate", Container)) %>課程?');return document.MM_returnValue"/></a>
                                <% End If%>
                              </contentstemplate>
                                                                        </MM:If>
                                                                        <MM:If runat="server" Expression='<%# Da_classopen.FieldValue("class_opendate", Container) < DateTime.Now.ToString("yyyy/MM/dd", System.Globalization.DateTimeFormatInfo.InvariantInfo) %>'>
                                                                            <contentstemplate><font color="#990000">過期</font></contentstemplate>
                                                                        </MM:If>
                                                                    </ItemTemplate>
                                                                </asp:TemplateColumn>
                                                            </Columns>
                                                        </asp:DataGrid>
                                                    </form>
                                                </div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div align="center"></div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div align="center" class="subtitle"></div>
                                            </td>
                                        </tr>
                                    </table>
                                </div>
                            </td>
                            <td width="11" valign="top" background="images/test_06.gif">&nbsp;</td>
                        </tr>
                        <tr>
                            <td width="11" height="11" valign="top">
                                <img src="images/test_20.gif" width="11" height="11" /></td>
                            <td height="11" colspan="2" valign="top">
                                <img src="images/test_21.gif" width="100%" height="11" /></td>
                            <td width="11" height="11" valign="top">
                                <img src="images/test_22.gif" width="11" height="11" /></td>
                        </tr>
                    </table>
                </div>
            </td>
            <td width="15" background="images/login_rt.gif">&nbsp;</td>
        </tr>

        <tr>
            <td width="6" height="11" valign="top">&nbsp;</td>
            <td width="5" height="11" valign="top">
                <img src="images/login_lf_bt_coer.gif" width="11" height="11" border="0" /></td>
            <td height="11" valign="top" background="images/login_bt_line.gif">&nbsp;</td>
            <td width="15" height="11" valign="top">
                <img src="images/login_rt_bt_coer.gif" width="11" height="11" border="0" /></td>
        </tr>
        <tr>
            <td height="11" valign="top">&nbsp;</td>
            <td valign="top">&nbsp;</td>
            <td height="11" valign="top">
                <div align="center" class="en1">Copyright © 2012 - <% = Year(DateTime.Now)  %>  <%# Da_web.FieldValue("ser_name", Container) %>  系統支援 <a href="http://www.mico.com.tw/" target="_blank" />MICO </a></div>
            </td>
            <td height="11" valign="top">&nbsp;</td>
        </tr>
    </table>

</body>
</html>
