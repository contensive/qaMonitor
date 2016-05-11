

Imports System
Imports System.Collections.Generic
Imports System.Text
Imports Contensive.BaseClasses

Namespace contensive.addon.contensiveQAMonitor
    '
    ' Sample Vb2005 addon
    '
    Public Class qaHouskeeping
        Inherits AddonBaseClass
        '
        ' - update references to your installed version of cpBase
        ' - Verify project root name space is empty
        ' - Change the namespace to the collection name
        ' - Change this class name to the addon name
        ' - Create a Contensive Addon record with the namespace apCollectionName.ad
        ' - add reference to CPBase.DLL, typically installed in c:\program files\kma\contensive\
        '
        '=====================================================================================
        ' addon api
        '=====================================================================================
        '
        Public Overrides Function Execute(ByVal CP As CPBaseClass) As Object
            Dim returnHtml As String
            Dim orgName As String = ""
            Dim noEmptyOrgs As Integer = 0
            Dim cs As CPCSBaseClass = CP.CSNew
            Dim csSet As CPCSBaseClass = CP.CSNew
            Dim id As Integer = 0
            Dim blockLayoutHtml As String = ""
            Dim blockLayout As CPBlockBaseClass = CP.BlockNew
            Dim layout As CPBlockBaseClass = CP.BlockNew

            layout.OpenLayout("qaHouseKeepingLayout")



            blockLayout.Load(layout.GetInner("#js-housekeep"))



            Try
                If cs.Open("organizations") Then
                    Do
                        orgName = cs.GetText("Name")
                        If orgName = "" Then
                            noEmptyOrgs += 1
                        End If
                        Call cs.GoNext()
                    Loop While cs.OK

                End If
                Call cs.Close()
               


                blockLayout.SetInner(".emptyOrgClass", "<h3> The Number of Companies where the Name is empty is: " & noEmptyOrgs & "</h3>")

                blockLayoutHtml &= blockLayout.GetHtml
                layout.SetInner("#js-housekeep", blockLayoutHtml)

                If csSet.Insert("Site Warnings") Then
                    Call csSet.SetField("Name", "QaHousekeeping Warning")
                    Call csSet.SetField("shortDescription", "The Number of Companies where the Name is empty is: " & noEmptyOrgs)
                    Call csSet.SetField("DateLastReported", Now.ToString)
                End If
                Call csSet.Close()

                returnHtml = layout.GetHtml

            Catch ex As Exception
                errorReport(CP, ex, "execute")
                returnHtml = "Visual Studio 2005 Contensive Addon - Error response"
            End Try
            Return returnHtml
        End Function
        '
        '=====================================================================================
        ' common report for this class
        '=====================================================================================
        '
        Private Sub errorReport(ByVal cp As CPBaseClass, ByVal ex As Exception, ByVal method As String)
            Try
                cp.Site.ErrorReport(ex, "Unexpected error in sampleClass." & method)
            Catch exLost As Exception
                '
                ' stop anything thrown from cp errorReport
                '
            End Try
        End Sub
    End Class
End Namespace
