Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Xml
Imports System.Drawing
Imports System.Reflection
Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports FastReport
Imports FastReport.Export.Pdf
Imports FastReport.Export.OoXML
Imports FastReport.Export


Imports System.Web.Configuration



Partial Class admin_ExportWordFr_FrExportWord
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        'create fr
        Dim fr As New FastReport.Report
        'create dt
        Dim dt As New DataTable

        'connect with sign query
        Dim qry As String = <string>
                                SELECT 
 ROW_NUMBER() OVER(ORDER BY sign_joindate ASC, sign_id ASC) AS Number,
sign_id, sign_flowid, sign_flowsn, sign_classname, CONVERT(char(10), sign_joindate,111)as sign_joindate, 
CONVERT(char(10), sign_opendate,111)as sign_opendate, CONVERT(char(10), sign_enddate,111)as sign_enddate, 
sign_name, sign_member_idno, sign_sex, CONVERT(char(10), sign_birth,111)as sign_birth, 
sign_org, sign_job2, sign_job2_other, sign_school, sign_org_zip, sign_org_adds, substring(sign_sendcertzip,1,3) as mailzip, 
RTRIM(CAST(substring(sign_sendcertzip,5,20) as Varchar(20))) + RTRIM(CAST(sign_sendcertadds as Varchar(100))) AS home_adds, 
substring(sign_percertzip,1,3) as mailperzip, 
RTRIM(CAST(substring(sign_percertzip,5,20) as Varchar(20))) + RTRIM(CAST(sign_percertadds as Varchar(100))) AS home_peradds, 
sign_tel, sign_mobil, sign_mail, sign_fax,  sign_classfee, sign_fee, sign_payban, sign_feeno, sign_certificate, sign_certificate2, sign_certificate3, sign_contactname, sign_contacttel, sign_contactfax, sign_contactmail, sign_opinion, sign_meno, sign_cancel, 
(select class_name from yurta_classopen where class_id=@sign_id) AS class_name  FROM dbo.yurta_sign  
WHERE sign_flowid = @sign_id AND sign_cancel = @sign_cancel  
ORDER BY sign_joindate ASC, sign_id ASC


                            </string>
        Using cn As New SqlConnection
            cn.ConnectionString = WebConfigurationManager.AppSettings("MM_CONNECTION_STRING_conn_iia")
            cn.Open()
            Dim cmd As New SqlCommand
            cmd.Connection = cn
            cmd.CommandText = qry
            cmd.Parameters.Clear()
            cmd.Parameters.AddWithValue("sign_id", Session("class_id"))
            cmd.Parameters.AddWithValue("sign_cancel", 0)
            Dim dr As SqlDataReader
            dr = cmd.ExecuteReader()
            dt.Load(dr)
            dr.Close()
        End Using
        fr.RegisterData(dt, "sign")
        fr.Report.Load(HttpRuntime.AppDomainAppPath + "admin\FRX\sign.frx")


        Dim databand1 As FastReport.DataBand = fr.FindObject("Data1")
        databand1.DataSource = fr.GetDataSource("sign")

        Select Case Session("ExportType")
            Case "Pdf"
                'fr.Prepare(True)
                Dim expdf As New FastReport.Export.Pdf.PDFExport
                'fr.Export(expdf, fname)

                '內嵌字型
                expdf.EmbeddingFonts = Not Request.UserAgent.ToLower.Contains("windows")

                If fr.Prepare() Then
                    Dim St As IO.MemoryStream = New IO.MemoryStream
                    fr.Export(expdf, St)
                    St.Position = 0
                    Dim buff(St.Length) As Byte
                    St.Read(buff, 0, St.Length)
                    St.Close()
                    Response.ContentType = "application/pdf"
                    Response.Expires = 0
                    Response.BinaryWrite(buff)
                    Response.End()
                    St.Close()
                    St.Dispose()
                End If
                fr.Dispose()
            Case "Word"
                'fr.Prepare(True)
                Dim exword As New FastReport.Export.OoXML.Word2007Export

                '內嵌字型
                'exword.EmbeddingFonts = Not Request.UserAgent.ToLower.Contains("windows")

                If fr.Prepare() Then
                    Dim St As IO.MemoryStream = New IO.MemoryStream
                    fr.Export(exword, St)
                    St.Position = 0
                    Dim buff(St.Length - 1) As Byte
                    St.Read(buff, 0, St.Length)
                    St.Close()
                    Response.Clear()
                    Response.Charset = ""
                    Response.Cache.SetCacheability(HttpCacheability.ServerAndPrivate)
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    Response.AddHeader("Content-Disposition", "attachment ;filename=ContentDocument.docx")
                    Response.AddHeader("Content-Length", buff.Length.ToString())


                    Response.Expires = 0
                    Response.BinaryWrite(buff)
                    Response.End()
                    Response.Flush()


                    'Response.OutputStream.Write(buff, 0, buff.Length.ToString())
                    'Response.Flush()
                    'Response.End()
                    St.Close()
                    St.Dispose()
                End If
                fr.Dispose()



        End Select





       
       







    End Sub
End Class
