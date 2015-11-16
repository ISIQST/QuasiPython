Public Class Sample1
    Implements Quasi97.iHOption

    Private NList As New List(Of Object)
    Private QST As Quasi97.Application
    Private myInst$ = ""
    Private myStatus% = Quasi97.clsHardwareOption.HDWOptStat.Missing
    Private ptrForm As New frmSampleStatus

    Private countStartStop% = 0
    Private lHaltCounting As Boolean
    Public pyEngine As Microsoft.Scripting.Hosting.ScriptEngine
    Public pyScope As Microsoft.Scripting.Hosting.ScriptScope

    Public Sub AddNotifier(ByRef objnot As Object) Implements Quasi97.iHOption.AddNotifier
        If Not NList.Contains(objnot) Then NList.Add(objnot)
    End Sub

    Public Sub RemoveNotifier(ByRef objnot As Object) Implements Quasi97.iHOption.RemoveNotifier
        If NList.Contains(objnot) Then NList.Remove(objnot)
    End Sub

    Private Sub InitializeIronPython()
        'IronPython.Hosting.Options.PrivateBinding = True
        pyEngine = IronPython.Hosting.Python.CreateEngine

        pyScope = pyEngine.CreateScope
        'pyEngine.AddToPath(AppDomain.CurrentDomain.BaseDirectory)
        pyScope.SetVariable("qst", QST)
        'Dim src = pyEngine.CreateScriptSourceFromString("qst.AppVersion")
        'Dim compiled = src.Compile()
        'Debug.Print(compiled.Execute(pyScope))
    End Sub

    Public Sub Initialize3(InstanceName As String, ByRef AppPtr As Object) Implements Quasi97.iHOption.Initialize3
        Try
            QST = AppPtr
            myInst = InstanceName
            myStatus = 0            'not missing

            InitializeIronPython()

            Dim fn As Integer = FreeFile()
            Try
                FileOpen(fn, "log.csv", OpenMode.Input)
                Call Input(fn, countStartStop)
                FileClose(fn)
            Catch ex As Exception

            End Try

            QST.QstStatus.Message = "Hello World!"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub Terminate() Implements Quasi97.iHOption.Terminate
        Dim fn As Integer = FreeFile()
        Try
            FileOpen(fn, "log.csv", OpenMode.Input)
            Call Write(fn, countStartStop)
            FileClose(fn)
        Catch ex As Exception

        End Try

        ptrForm.ownerObj = Nothing
        ptrForm.Close()
        QST = Nothing
        NList.Clear()
    End Sub

    Private Sub NotifyUser(lastUserFeeback As String, prg!)
        Dim s As Object
        For Each s In NList
            s.Message = lastUserFeeback
            s.Progress = prg
        Next
    End Sub

    Public Sub DetectAllNew(ByRef Devs1Based() As String) Implements Quasi97.iHOption.DetectAllNew
        ReDim Devs1Based(1)
        Devs1Based(1) = "main"
    End Sub

    Public ReadOnly Property EventInterests As Integer Implements Quasi97.iHOption.EventInterests
        Get
            Return Quasi97.clsOptionManager.eEventInt.eCheckHealth Or Quasi97.clsOptionManager.eEventInt.eStartStop
        End Get
    End Property

    Public ReadOnly Property Status As Short Implements Quasi97.iHOption.Status
        Get
            Return myStatus
        End Get
    End Property

    Private Function ConnectionLost() As Boolean
        Return False
    End Function

    Public Function CheckHealth(ByRef usrDescr As String, PartLoadedState As Byte) As Short Implements Quasi97.iHOption.CheckHealth
        'checking health here
        If myStatus <> Quasi97.clsHardwareOption.HDWOptStat.Suspended Then
            If ConnectionLost() Then
                myStatus = Quasi97.clsHardwareOption.HDWOptStat.Suspended
                Return 1
            End If
        End If
        Return 0 'no problem
    End Function

    Public Sub ShowDiagnostics() Implements Quasi97.iHOption.ShowDiagnostics
        If ptrForm.IsDisposed Then
            ptrForm = New frmSampleStatus
        End If
        ptrForm.ownerObj = Me
        ptrForm.Show()
    End Sub

    Public Sub ShowUserMenu() Implements Quasi97.iHOption.ShowUserMenu
        ShowDiagnostics()
    End Sub

    Public Sub ConnectHead(ByRef doConnect As Byte) Implements Quasi97.iHOption.ConnectHead

    End Sub

    Public Sub CurrentHeadInitiate(ByRef hdnum As Short) Implements Quasi97.iHOption.CurrentHeadInitiate

    End Sub

    Public Sub CurrentHeadTerminate(hdNum As Short) Implements Quasi97.iHOption.CurrentHeadTerminate

    End Sub

    Public Sub GetChannels(ByRef Channels() As Short) Implements Quasi97.iHOption.GetChannels

    End Sub

    Public Function GetNewProperties2(ByRef propDetails(,) As String) As Object Implements Quasi97.iHOption.GetNewProperties2
        Dim colobjects As New ArrayList
        'propdetails has 1+colobjects.count number of rows and 4 columns.
        Call AddCustomProperty("Halt Counter", "HaltCounter", True, "0,1", False, colobjects, propDetails, Me)
        Return colobjects
    End Function

    Private Sub AddCustomProperty(ByVal UsrName$, ByVal PropName$, ByVal RestrChoice As Boolean, ByVal availlist$, ByVal seqstress As Boolean, ByRef colobjects As ArrayList, ByRef propDetails(,) As String, ByVal objptr As Object)
        'col 0 is the property name
        'col 1 is the user name
        'col 2 is the restricted choice or n ot
        'col 3 is semicolon separated list of possible values for validation
        'col 4 is the flag for stress that must be sequentially rolled backward when restorting parameterse
        Dim i%
        colobjects.Add(objptr)
        i = colobjects.Count
        ReDim Preserve propDetails(4, colobjects.Count)
        propDetails(0, i) = UsrName
        propDetails(1, i) = PropName
        propDetails(2, i) = CStr(RestrChoice)
        propDetails(3, i) = availlist
        propDetails(4, i) = CStr(seqstress)
    End Sub

    Public Property HaltCounter As Boolean
        Get
            Return lHaltCounting
        End Get
        Set(value As Boolean)
            lHaltCounting = value
        End Set
    End Property

    Public WriteOnly Property NetHostCallBack As Object Implements Quasi97.iHOption.NetHostCallBack
        Set(value As Object)

        End Set
    End Property

    Public Sub NotifyOptionsUpdated() Implements Quasi97.iHOption.NotifyOptionsUpdated

    End Sub

    Public Function Recover(ByRef usrDescr As String) As Short Implements Quasi97.iHOption.Recover
        Return 0
    End Function

    Public Sub SetChannels(ByRef Channels() As Short) Implements Quasi97.iHOption.SetChannels

    End Sub

    Public Sub SetupOpenClose(DoOpen As Boolean) Implements Quasi97.iHOption.SetupOpenClose

    End Sub

    Public Sub StartStop(ByRef doStart As Boolean) Implements Quasi97.iHOption.StartStop
        If Not lHaltCounting Then countStartStop += 1
    End Sub

    Public ReadOnly Property UserControl As String Implements Quasi97.iHOption.UserControl
        Get
            Return Nothing
        End Get
    End Property
End Class
