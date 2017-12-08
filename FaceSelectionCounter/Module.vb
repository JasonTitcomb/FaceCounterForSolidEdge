'Written by Jason Titcomb 12/8/2017
'This code is provided AS IS.
Imports System.Runtime.InteropServices
Imports SolidEdgeConstants, SolidEdgeFramework
Imports SolidEdgeFramework.DocumentTypeConstants

Module Module1
    Private mSolidApp As SolidEdgeFramework.Application = Nothing
    Private mCommand As SolidEdgeFramework.Command = Nothing
    Private mMouse As SolidEdgeFramework.Mouse = Nothing
    Private mHighlightSets As SolidEdgeFramework.HighlightSets = Nothing
    Private mHighlightSet As SolidEdgeFramework.HighlightSet = Nothing

    Sub Main()
        Try
            mSolidApp = Marshal.GetActiveObject("SolidEdge.Application")
            MessageFilter.Register()

            Dim solidDoc As Object = Nothing
            Select Case mSolidApp.ActiveDocumentType
                Case igPartDocument, igSyncPartDocument, igSheetMetalDocument, igSyncSheetMetalDocument
                    solidDoc = mSolidApp.ActiveDocument
                Case Else
                    ReportStatus("Part or Sheetmetal only!")
                    Exit Sub
            End Select

            mHighlightSets = solidDoc.HighlightSets
            mCommand = mSolidApp.CreateCommand(SolidEdgeConstants.seCmdFlag.seNoDeactivate)
            AddHandler mCommand.Terminate, AddressOf Command_Terminate

            mCommand.Start()
            mMouse = mCommand.Mouse
            With mMouse
                .LocateMode = 2
                .WindowTypes = 1
                .EnabledMove = True
                .AddToLocateFilter(seLocateFilterConstants.seLocateFace)
                AddHandler .MouseDown, AddressOf Mouse_MouseDown
            End With
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            ReportStatus(ex.Message)
        Finally
            MessageFilter.Revoke()
        End Try

    End Sub

    Private Sub ReportStatus(msg As String)
        mSolidApp.StatusBar = msg
    End Sub

    Private Sub Mouse_MouseDown(ByVal sButton As Short, ByVal sShift As Short, ByVal dX As Double, ByVal dY As Double, ByVal dZ As Double, ByVal pWindowDispatch As Object, ByVal lKeyPointType As Integer, ByVal pGraphicDispatch As Object)
        If sButton = 2 Or pGraphicDispatch Is Nothing Then
            'Mouse down and nothing selected
            mSolidApp.StartCommand(SolidEdgeFramework.SolidEdgeCommandConstants.sePartSelectCommand)
        Else
            FindFaces(pGraphicDispatch)
        End If
    End Sub

    Private Sub Command_Terminate()
        If mHighlightSet IsNot Nothing Then
            mHighlightSet.RemoveAll()
            ReleaseRCW(mHighlightSet)
            ReleaseRCW(mHighlightSets)
        End If

        ReleaseRCW(mMouse)
        ReleaseRCW(mCommand)
        ReleaseRCW(mSolidApp)
        System.Windows.Forms.Application.Exit()
    End Sub

    Private Sub FindFaces(theFace As SolidEdgeGeometry.Face)
        Dim faces As Object
        Try
            ReportStatus("Searching...")
            If theFace IsNot Nothing Then
                Dim facetype As SolidEdgeGeometry.GNTTypePropertyConstants = Nothing
                facetype = theFace.Geometry.Type
                Dim bdy As SolidEdgeGeometry.Body = Nothing
                If TryGetProp(theFace, "Body", bdy) Then
                    Select Case facetype
                        Case SolidEdgeGeometry.GNTTypePropertyConstants.igCone
                            faces = bdy.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryCone)
                        Case SolidEdgeGeometry.GNTTypePropertyConstants.igCylinder
                            faces = bdy.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryCylinder)
                        Case SolidEdgeGeometry.GNTTypePropertyConstants.igPlane
                            faces = bdy.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryPlane)
                        Case SolidEdgeGeometry.GNTTypePropertyConstants.igBSplineSurface
                            faces = bdy.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQuerySpline)
                        Case SolidEdgeGeometry.GNTTypePropertyConstants.igTorus
                            faces = bdy.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryTorus)
                        Case SolidEdgeGeometry.GNTTypePropertyConstants.igSphere
                            faces = bdy.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQuerySphere)
                        Case Else
                            faces = bdy.Faces(SolidEdgeGeometry.FeatureTopologyQueryTypeConstants.igQueryAll)
                    End Select

                    If mHighlightSet Is Nothing Then
                        mHighlightSet = mHighlightSets.Add
                    Else
                        mHighlightSet.RemoveAll()
                    End If

                    Dim selArea = Math.Round(theFace.Area, 7)

                    For r As Integer = 1 To faces.count
                        Dim face As SolidEdgeGeometry.Face = faces.Item(r)
                        If Math.Round(face.Area, 7) = selArea Then
                            mHighlightSet.AddItem(face)
                        End If
                        ReleaseRCW(face)
                    Next
                    ReportStatus(mHighlightSet.Count & " of " & faces.count & "  " & facetype.ToString)
                    mHighlightSet.Draw()
                    ReleaseRCW(bdy)
                End If
            Else
                ReportStatus("Position mouse over face")
            End If
        Finally

        End Try


    End Sub

    Private Function TryGetProp(o As Object, name As String, ByRef retProp As Object) As Boolean
        retProp = o.GetType.InvokeMember(name, Reflection.BindingFlags.GetProperty, Nothing, o, Nothing)
        Return retProp IsNot Nothing
    End Function

    Private Sub ReleaseRCW(ByRef o As Object)
        If o IsNot Nothing Then
            Dim ret As Integer = Marshal.ReleaseComObject(o)
            'Debug.Assert(0 = ret)
            o = Nothing
        End If
    End Sub

End Module
