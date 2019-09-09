Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class Form2

    Private Sub CrystalReportViewer1_Load(sender As Object, e As EventArgs) Handles CrystalReportViewer1.Load
        'MsgBox(Form1.txtmod.Text)
        Try

        
        Dim RptDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument
        'Dim MyVal1 As Integer = Form1.txttemp.Text
        If Form1.txtmod.Text = "Arrest" Then


            RptDocument.Load("\\rms-prod-rpt\Boulder\Cyclops\Arrest.rpt")
            'RptDocument.Load("G:\SunGard\Programs\MFR_query\Arrest.rpt")
            CrystalReportViewer1.ReportSource = RptDocument
            'CrystalReportViewer1.ReportSource = Arrest1
            RptDocument.SetParameterValue("key", Form1.txttemp.Text)
        End If
        If Form1.txtmod.Text = "Field Contact" Then
            RptDocument.Load("\\rms-prod-rpt\Boulder\Cyclops\FI.rpt")
            'RptDocument.Load("G:\SunGard\Programs\MFR_query\FI.rpt")
            CrystalReportViewer1.ReportSource = RptDocument
            RptDocument.SetParameterValue("key", Form1.txttemp.Text)
        End If
        If Form1.txtmod.Text = "Incident" Then
            RptDocument.Load("\\rms-prod-rpt\Boulder\Cyclops\Law.rpt")
            'RptDocument.Load("G:\SunGard\Programs\MFR_query\Law.rpt")
            CrystalReportViewer1.ReportSource = RptDocument
            RptDocument.SetParameterValue("key", Form1.txttemp.Text)
        End If
        If Form1.txtmod.Text = "Supplement" Then
            RptDocument.Load("\\rms-prod-rpt\Boulder\Cyclops\Supp.rpt")
            'RptDocument.Load("G:\SunGard\Programs\MFR_query\Supp.rpt")
            CrystalReportViewer1.ReportSource = RptDocument
            RptDocument.SetParameterValue("key", Form1.txttemp.Text)
        End If
        If Form1.txtmod.Text = "Accident" Then
            RptDocument.Load("\\rms-prod-rpt\Boulder\Cyclops\Traffic.rpt")
            'RptDocument.Load("G:\SunGard\Programs\MFR_query\Traffic.rpt")
            CrystalReportViewer1.ReportSource = RptDocument
            RptDocument.SetParameterValue("key", Form1.txttemp.Text)
        End If
        If Form1.txtmod.Text = "Property" Then
            RptDocument.Load("\\rms-prod-rpt\Boulder\Cyclops\Property.rpt")
            'RptDocument.Load("G:\SunGard\Programs\MFR_query\Property.rpt")
            CrystalReportViewer1.ReportSource = RptDocument
            RptDocument.SetParameterValue("key", Form1.txttemp.Text)
        End If
        'myDataReport.SetParameterValue("MyParameter2", "Hello2");
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
