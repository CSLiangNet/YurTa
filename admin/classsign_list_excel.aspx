<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="utf-8" EnableEventValidation="false" Debug="true" %>

<%@ Import Namespace="System.Web.Configuration" %>

<%@ Import Namespace="System.Data.SqlClient" %>
<%@ Register TagPrefix="MM" Namespace="DreamweaverCtrls" Assembly="DreamweaverCtrls,version=1.0.0.0,publicKeyToken=836f606ede05d46a,culture=neutral" %>


<%@ Import Namespace="System.Collections.Generic" %>
<%@ Import Namespace="System.Text" %>
<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="DocumentFormat.OpenXml.Packaging" %>

<%--
<%@ Import Namespace  = "Ap=DocumentFormat.OpenXml.ExtendedProperties" %>
<%@ Import Namespace  = "DocumentFormat.OpenXml.Wordprocessing" %>
<%@ Import Namespace  = "DocumentFormat.OpenXml"%>
<%@ Import Namespace  = "Ds=DocumentFormat.OpenXml.CustomXmlDataProperties"%>
<%@ Import Namespace  = "Ovml=DocumentFormat.OpenXml.Vml.Office"%>
<%@ Import Namespace  = "V=DocumentFormat.OpenXml.Vml"%>
<%@ Import Namespace  = "M=DocumentFormat.OpenXml.Math"%>
<%@ Import Namespace  = "W15=DocumentFormat.OpenXml.Office2013.Word"%>
<%@ Import Namespace  = "A=DocumentFormat.OpenXml.Drawing"%>
<%@ Import Namespace  = "Thm15=DocumentFormat.OpenXml.Office2013.Theme"%
>--%>




<MM:DataSet
    ID="Da_classopen"
    runat="Server"
    IsStoredProcedure="false"
    ConnectionString='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_STRING_conn_iia") %>'
    DatabaseType='<%# System.Configuration.ConfigurationSettings.AppSettings("MM_CONNECTION_DATABASETYPE_conn_iia") %>'
    CommandText='<%# "SELECT sign_id, sign_flowid, sign_flowsn, sign_classname, CONVERT(char(10), sign_joindate,111)as sign_joindate, CONVERT(char(10), sign_opendate,111)as sign_opendate, CONVERT(char(10), sign_enddate,111)as sign_enddate, sign_name, sign_member_idno, sign_sex, CONVERT(char(10), sign_birth,111)as sign_birth, sign_org, sign_job2, sign_job2_other, sign_school, sign_org_zip, sign_org_adds, substring(sign_sendcertzip,1,3) as mailzip, RTRIM(CAST(substring(sign_sendcertzip,5,20) as Varchar(20))) + RTRIM(CAST(sign_sendcertadds as Varchar(100))) AS home_adds, substring(sign_percertzip,1,3) as mailperzip, RTRIM(CAST(substring(sign_percertzip,5,20) as Varchar(20))) + RTRIM(CAST(sign_percertadds as Varchar(100))) AS home_peradds, sign_tel, sign_mobil, sign_mail, sign_fax,  sign_classfee, sign_fee, sign_payban, sign_feeno, sign_certificate, sign_certificate2, sign_certificate3, sign_contactname, sign_contacttel, sign_contactfax, sign_contactmail, sign_opinion, sign_meno, sign_cancel, (select class_name from yurta_classopen where class_id=@sign_id) AS class_name  FROM dbo.yurta_sign  WHERE sign_flowid = @sign_id AND sign_cancel = @sign_cancel  ORDER BY sign_joindate ASC, sign_id ASC" %>'
    Debug="true">
    <parameters>
<Parameter  Name="@sign_id"  Value='<%# IIf((Request.QueryString("class_id") <> Nothing), Request.QueryString("class_id"), "") %>'  Type="Int"   />
<Parameter  Name="@sign_cancel"  Value='<%# 0 %>'  Type="Int"   />
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
<script runat="server">
    Sub Page_Load(Src As Object, E As EventArgs)
        '若PS是Nothign就不執行本頁
        If Session("IIA_userac") = Nothing Then
            '移除目前的 Session
            Session.Abandon()
            Response.Redirect("login.aspx?erry=timeout")
        End If
        If Session("IIA_syD") = Nothing Then
            '沒有IIA_syD的管理權限
            Response.Redirect("admin.aspx?erry=0")
        End If
        If Not IsPostBack Then
            DataGrid1.DataBind()
        End If
    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs)
        
        
        Dim strDocid As String = Page.Request.QueryString("docids")

        'Connect to the database and bring back the image contents & MIME type for the specified picture 
        Using Conn As New SqlConnection(WebConfigurationManager.AppSettings("MM_CONNECTION_STRING_conn_iia"))
            'set param val 
            Const SQL As String = "SELECT [manager_id], [manager_name]      ,[manager_ac]  FROM [dbo].[manager]"
            Dim myCommand As New SqlCommand(SQL, Conn)
            'myCommand.Parameters.AddWithValue("@docids" strDocid)'

            Conn.Open()
            Dim myReader As SqlDataReader = myCommand.ExecuteReader

            If myReader.Read Then
                'Response.ContentType = myReader("DocMIMEType").ToString() 
                'Response.AddHeader("Content-Length", myReader("filesize").ToString()) 
                'Response.BinaryWrite(myReader("document")) 

                
                Dim bytes() As Byte = UnicodeStringToBytes("test")
                
                Response.Buffer = True
                Response.Clear()
                Response.Charset = ""
                'Response.Cache.SetCacheability(HttpCacheability.NoCache)
                Response.Cache.SetCacheability(HttpCacheability.ServerAndPrivate)
                
                'Response.ContentType = myReader("DocMIMEType").ToString()
                '& dt.Rows(0)("Name").ToString()) 
                Dim fs As Integer = bytes.Length - 1
                            
                
                 
                
                
                Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                Response.AddHeader("content-disposition", "attachment;filename=download.docx")
                Response.OutputStream.Write(bytes, 0, fs)
                Response.BinaryWrite(bytes)
                Response.Flush()
                Response.End()
                
               
                Conn.Close()
                

                myReader.Close()
            End If
        End Using
   
        
        
        
        'Dim template As String = Server.MapPath("Notice.docx")

        'Dim dct As New Dictionary(Of String, String)() From { _
        '    {"TODAY", DateTime.Today.ToString("yyyy-MM-dd")}, _
        '    {"NAME", "test for open xml sdk 2.5"}, _
        '    {"CONT", "TEST"}, _
        '    {"SALARY", "2000000"} _
        '}

        'Response.Clear()
        'Response.ContentType = "application/octet-stream"
        'Response.AddHeader("content-disposition", "attachment;filename=Notice.docx")
        'Response.BinaryWrite(DocxHelper.MakeDocx(Server.MapPath("Notice.docx"), dct))
        'Response.[End]()
        

    End Sub
    
 
    Private Function UnicodeStringToBytes(ByVal str As String) As Byte()

        Return System.Text.Encoding.Unicode.GetBytes(str)
    End Function
    
    
    
    
    
    
    
    
    
    


    ''' <summary>
    ''' Ver 1.0 By Jeffrey Lee, 2009-07-29
    ''' </summary>
    Public Class DocxHelper
        ''' <summary>
        ''' Replace the parser tags in docx document
        ''' </summary>
        ''' <param name="oxp">OpenXmlPart object</param>
        ''' <param name="dct">Dictionary contains parser tags to replace</param>
        Private Shared Sub parse(oxp As OpenXmlPart, dct As Dictionary(Of String, String))
            Dim xmlString As String = Nothing
            Using sr As New StreamReader(oxp.GetStream())
                xmlString = sr.ReadToEnd()
            End Using
            For Each key As String In dct.Keys
                xmlString = xmlString.Replace((Convert.ToString("[$") & key) + "$]", dct(key))
            Next
            Using sw As New StreamWriter(oxp.GetStream(FileMode.Create))
                sw.Write(xmlString)
            End Using
        End Sub

        ''' <summary>
        ''' Parse template file and replace all parser tags, return the binary content of
        ''' new docx file.
        ''' </summary>
        ''' <param name="templateFile">template file path</param>
        ''' <param name="dct">a Dictionary containing parser tags and values</param>
        ''' <returns></returns>
        Public Shared Function MakeDocx(templateFile As String, dct As Dictionary(Of String, String)) As Byte()
            Dim tempFile As String = Path.GetTempPath() + ".docx"
            File.Copy(templateFile, tempFile)

            Using wd As WordprocessingDocument = WordprocessingDocument.Open(tempFile, True)
                'Replace document body
                parse(wd.MainDocumentPart, dct)
                For Each hp As HeaderPart In wd.MainDocumentPart.HeaderParts
                    parse(hp, dct)
                Next
                For Each fp As FooterPart In wd.MainDocumentPart.FooterParts
                    parse(fp, dct)
                Next
            End Using
            Dim buff As Byte() = File.ReadAllBytes(tempFile)
            File.Delete(tempFile)
            Return buff
        End Function

    End Class
    
    
    '匯出word
    Protected Sub btnExportWord_Click(sender As Object, e As EventArgs)
        If Request.QueryString("class_id") <> Nothing And Request.QueryString("class_id") <> "" Then
            'use open xml sdk
            'Response.Redirect("~/admin/ExportWord/ExportWord.aspx")            
            
            'use fastreport.net
            Session("class_id") = Request.QueryString("class_id")
            'Response.Redirect("~/admin/ExportWordFr/FrExportWord.aspx")
            Session("ExportType") = "Word"
            Dim url As String = "/web_yurta/admin/ExportWordFr/FrExportWord.aspx"
            Dim s As String = "window.open('" & url + "', 'popup_window', 'width=1024,height=800,left=100,top=100,resizable=yes');"
            ClientScript.RegisterStartupScript(Me.GetType(), "script", s, True)
            
        End If
        
        

    End Sub

    '匯出pdf
    Protected Sub btnExportPDF_Click(sender As Object, e As EventArgs)
        If Request.QueryString("class_id") <> Nothing And Request.QueryString("class_id") <> "" Then
            'use open xml sdk
            'Response.Redirect("~/admin/ExportWord/ExportWord.aspx")            
            
            'use fastreport.net
            Session("class_id") = Request.QueryString("class_id")
            'Response.Redirect("~/admin/ExportWordFr/FrExportWord.aspx")
            Session("ExportType") = "Pdf"
            
            Dim url As String = "/web_yurta/admin/ExportWordFr/FrExportWord.aspx"
            Dim s As String = "window.open('" & url + "', 'popup_window', 'width=1024,height=800,left=100,top=100,resizable=yes');"
            ClientScript.RegisterStartupScript(Me.GetType(), "script", s, True)
            
            
            
        End If
        

    End Sub
</script>
<script runat="server">
    Sub Button1_Click(sender As Object, e As EventArgs)
        GJHRunReport()
    End Sub

    Private Sub GJHRemoveControls(ByVal control As Control)
        Dim i As Integer
        For i = control.Controls.Count - 1 To 0 Step -1
            GJHRemoveControls(control.Controls(i))
        Next i
        If Not TypeOf control Is TableCell Then
            If Not (control.GetType().GetProperty("SelectedItem") Is Nothing) Then
                control.Parent.Controls.Remove(control)
            Else
                If Not (control.GetType().GetProperty("Text") Is Nothing) Then
                    control.Parent.Controls.Remove(control)
                End If
            End If
        End If
    End Sub
    Private Sub GJHRunReport()
        Dim GJHXType, GJHXTension, GJHXContentType As String
        GJHXType = "Excel"
        If GJHXType = "Excel" Then
            GJHXTension = "xls"
            GJHXContentType = "application/vnd.ms-excel"
        ElseIf GJHXType = "Word" Then
            GJHXTension = "doc"
            GJHXContentType = "application/vnd.ms-word"
        End If
        GJHRemoveControls(DataGrid1)
        Response.Clear()
        Response.AddHeader("content-disposition", "attachment;filename=sign_A" + Request.QueryString("class_id") + "." + GJHXTension + "")
        Response.Charset = ""
        Response.Cache.SetCacheability(HttpCacheability.NoCache)
        Response.ContentType = GJHXContentType
        Dim stringWrite As System.IO.StringWriter = New System.IO.StringWriter
        Dim htmlWrite As HtmlTextWriter = New HtmlTextWriter(stringWrite)
        DataGrid1.RenderControl(htmlWrite)
        Response.Write(stringWrite.ToString())
        Response.End()
        Session.Abandon()
    End Sub
</script>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
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
                    <table width="100%" border="0" cellpadding="0" cellspacing="0">
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
                                        <td width="90%">
                                            <div align="left">
                                                <table width="100%" border="0">
                                                    <tr>
                                                        <td><span class="itemtitle">匯出報名</span><a href="javascript:history.back()"><img src="images/png-0048.png" alt="回前頁" width="32" height="32" border="0" /><span class="en1">back</span></a><a href="admin.aspx"><img src="images/png-0097.png" alt="回總管理" width="32" height="32" border="0" longdesc="admin.aspx" /><span class="en1">回總管理</span></a></td>
                                                    </tr>
                                                </table>
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
                    <br />
                    <table width="200%" border="0" align="left" cellpadding="0" cellspacing="0">
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
                                    <table border="0" align="left" cellpadding="0" cellspacing="0">
                                        <tr>
                                            <td>
                                                <div align="right"></div>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>
                                                <div align="left">
                                                    <form runat="server">
                                                        <asp:Button ID="Button1" OnClick="Button1_Click" runat="server" Text="匯出Excel"></asp:Button>
                                                        <asp:Button ID="Button2" runat="server" OnClick="Button2_Click" Text="OpenXmlSDKWord" Enabled="False" />
                                                        <asp:Button ID="btnExportWord" runat="server" OnClick="btnExportWord_Click" Text="匯出Word" />
                                                        <asp:Button ID="btnExportPDF" runat="server" OnClick="btnExportPDF_Click" Text="匯出PDF" />
                                                        <asp:DataGrid AllowPaging="false"
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
                                                                <asp:BoundColumn DataField="sign_id"
                                                                    HeaderText="signID"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_flowid"
                                                                    HeaderText="classID"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_flowsn"
                                                                    HeaderText="課程編號"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="class_name"
                                                                    HeaderText="課程名稱"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_opendate"
                                                                    HeaderText="課程起日"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_enddate"
                                                                    HeaderText="課程迄日"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_fee"
                                                                    HeaderText="結業"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_certificate"
                                                                    HeaderText="填寫證書字號"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_certificate2"
                                                                    HeaderText="產投證書字號"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_certificate3"
                                                                    HeaderText="結業證書字號"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_name"
                                                                    HeaderText="姓名"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_birth"
                                                                    HeaderText="生日"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_member_idno"
                                                                    HeaderText="身分證字號"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_job2"
                                                                    HeaderText="學歷"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_school"
                                                                    HeaderText="校系名稱"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_org"
                                                                    HeaderText="服務單位"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="mailzip"
                                                                    HeaderText="郵遞區號"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="home_adds"
                                                                    HeaderText="聯絡地址"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="mailperzip"
                                                                    HeaderText="戶籍郵遞區號"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="home_peradds"
                                                                    HeaderText="戶籍地址"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_mobil"
                                                                    HeaderText="手機"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_mail"
                                                                    HeaderText="Mail"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_fax"
                                                                    HeaderText="傳真"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_contactmail"
                                                                    HeaderText="聯絡人Mail"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_contactname"
                                                                    HeaderText="承辦人"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_contacttel"
                                                                    HeaderText="承辦人電話"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_payban"
                                                                    HeaderText="統編"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_feeno"
                                                                    HeaderText="收據編號"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
                                                                <asp:BoundColumn DataField="sign_meno"
                                                                    HeaderText="備註"
                                                                    ReadOnly="true"
                                                                    Visible="True" />
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
            <td width="6" height="22" valign="top">&nbsp;</td>
            <td width="5" height="22" valign="top">
                <img src="images/login_lf_bt_coer.gif" width="11" height="11" border="0" /></td>
            <td height="22" valign="top" background="images/login_bt_line.gif">&nbsp;</td>
            <td width="15" height="22" valign="top">
                <img src="images/login_rt_bt_coer.gif" width="11" height="11" border="0" /></td>
        </tr>
        <tr>
            <td height="11" valign="top">&nbsp;</td>
            <td valign="top">&nbsp;</td>
            <td height="11" valign="top">
                <div align="center" class="en1">Copyright © 2012 - <% = Year(DateTime.Now)  %><%# Da_web.FieldValue("ser_name", Container) %>系統支援 <a href="http://www.mico.com.tw/" target="_blank" />MICO </a></div>
            </td>
            <td height="11" valign="top">&nbsp;</td>
        </tr>
    </table>

</body>
</html>
