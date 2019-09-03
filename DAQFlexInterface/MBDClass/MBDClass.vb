Imports System.Runtime.InteropServices

<Guid("27A333BA-1E5B-4ecf-B5C8-900422B68FB9")> _
Public Interface MBDTest
    Function FindDevices(ByVal a As Integer) As String
    Function CreateDevice(ByVal DeviceID As String) As String
    Function ReleaseDevice(ByVal DeviceID As String) As String
    Function SendMessage(ByVal Msg As String) As String
    Function ReadInScanMultiChan(ByRef DatArray As Double(,), _
    ByVal Samples As Integer, ByVal timeOut As Integer) As String
    Function WriteOutScan(ByRef DatArray As Double(,), ByVal _
    Samples As Integer, ByVal timeOut As Integer) As String
    Function StartBackgroundScan(ByVal FunctionType As String, _
    ByVal ReadBlockSize As Integer) As String
    Function StopBackgroundScan(ByVal FunctionType As String) As String
    Function EnableEvents(ByVal EventType As Integer, ByVal ReadBlockSize As Integer) As String
    Function DisableEvents(ByVal TypeOfCallback As Integer) As String
    Function GetSupportedMessages(ByVal Component As String) As String
    Function ConvertErrorCode(ByVal ErrorCode As Integer) As String
    Event DataAvailable(ByVal DeviceID As String, _
    ByVal EventType As Integer, ByRef callbackData As Object)
    Event ErrorAvailable(ByVal DeviceID As String, _
    ByVal ErrorCode As Integer)
    ReadOnly Property DeviceID() As String
    ReadOnly Property NumberOfDevices() As Integer
End Interface

<ClassInterface(ClassInterfaceType.None), ComSourceInterfaces(GetType(DataEvents.IDataEvents))> _
Public Class MBDComClass

    Implements MBDTest
    Shared DeviceHandles As System.Collections.Generic.IDictionary(Of String, DaqDevice) _
    = New Dictionary(Of String, DaqDevice)
    Shared UseCount As System.Collections.Generic.IDictionary(Of String, Integer) _
    = New Dictionary(Of String, Integer)
    Shared mnNumDevices As Integer
    Dim MyDevice As DaqDevice
    Dim msDeviceID As String
    Public Event DataAvailable(ByVal DeviceID As String, ByVal EventData As Integer, _
    ByRef DatArray As Object) Implements MBDTest.DataAvailable
    Public Event ErrorAvailable(ByVal DeviceID As String, _
    ByVal ErrorCode As Integer) Implements MBDTest.ErrorAvailable


    Public Sub New()

        ' Class initialization called when MBDInterface is instantiated
        MyBase.New()

    End Sub

    Public Function FindDevices(ByVal a As Integer) As String Implements MBDTest.FindDevices

        Dim DeviceList As String
        Dim Devices As String()
        Dim CurDev As Integer

        Devices = DaqDeviceManager.GetDeviceNames(a)

        DeviceList = ""
        mnNumDevices = Devices.Length
        For CurDev = 1 To mnNumDevices
            DeviceList = DeviceList & Devices(CurDev - 1) & "|"
        Next
        FindDevices = DeviceList

    End Function

    Public Function CreateDevice(ByVal DeviceID As String) As String Implements MBDTest.CreateDevice

        CreateDevice = DeviceID
        For Each ExistingDevice As KeyValuePair(Of String, DaqDevice) In DeviceHandles
            If DeviceID = ExistingDevice.Key Then
                msDeviceID = DeviceID
                MyDevice = ExistingDevice.Value
                UseCount(DeviceID) = UseCount(DeviceID) + 1
                Exit For
            End If
        Next
        If Not (msDeviceID = DeviceID) Then
            Try
                MyDevice = DaqDeviceManager.CreateDevice(DeviceID)
                DeviceHandles.Add(DeviceID, MyDevice)
                UseCount.Add(DeviceID, 1)
                CreateDevice = DeviceID
                msDeviceID = DeviceID
            Catch ex As Exception
                CreateDevice = "Error- " & ex.Message
            End Try
        End If

    End Function

    Public Function ReleaseDevice(ByVal DeviceID As String) As String Implements MBDTest.ReleaseDevice

        ReleaseDevice = ""
        If UseCount(DeviceID) = 1 Then
            MyDevice = DeviceHandles(DeviceID)
            DeviceHandles.Remove(DeviceID)
            Try
                DaqDeviceManager.ReleaseDevice(MyDevice)
                UseCount.Remove(DeviceID)
                msDeviceID = ""
            Catch ex As Exception
                ReleaseDevice = "Error- " & ex.Message
            End Try
        Else
            UseCount(DeviceID) = UseCount(DeviceID) - 1
        End If

    End Function

    Public Function SendMessage(ByVal Msg As String) As String Implements MBDTest.SendMessage

        Dim Response As DaqResponse
        Dim ErrString As String, ErrCategory As String
        Dim MBDWarning As Boolean
        Try
            Response = MyDevice.SendMessage(Msg)
            SendMessage = Response.ToString
        Catch ex As Exception
            ErrString = ex.Message
            MBDWarning = ErrString.Contains("The device does not support the command sent")
            If MBDWarning Then
                ErrCategory = "Warning- "
            Else
                ErrCategory = "Error- "
            End If
            SendMessage = ErrCategory & ErrString
        End Try

    End Function

    Public Function ReadInScanMultiChan(ByRef DatArray As Double(,), ByVal Samples As Integer, _
    ByVal timeOut As Integer) As String Implements MBDTest.ReadInScanMultiChan

        Dim ErrString As String, ErrCategory As String

        Try
            DatArray = MyDevice.ReadScanData(Samples, timeOut)
            ReadInScanMultiChan = "No error occurred."
        Catch ex As Exception
            ErrString = ex.Message
            ErrCategory = "Error- "
            ReadInScanMultiChan = ErrCategory & ErrString
        End Try

    End Function

    Public Function WriteOutScan(ByRef DatArray As Double(,), ByVal Samples As Integer, _
    ByVal timeOut As Integer) As String Implements MBDTest.WriteOutScan

        Dim ErrString As String, ErrCategory As String

        Try
            MyDevice.WriteScanData(DatArray, Samples, timeOut)
            WriteOutScan = "No error occurred."
        Catch ex As Exception
            ErrString = ex.Message
            ErrCategory = "Error- "
            WriteOutScan = ErrCategory & ErrString
        End Try

    End Function

    Public Function EnableEvents(ByVal EventType As Integer, _
    ByVal ReadBlockSize As Integer) As String Implements MBDTest.EnableEvents

        Dim ErrString As String, ErrCategory As String
        Try
            'MyDevice.(AddressOf AInScanCallback, ReadBlockSize, True)
            'EnableEvents = ReadBlockSize.ToString
            'ErrCategory = "Error- Callback not yet implemented. "
            MyDevice.EnableCallback(AddressOf AInScanCallback, EventType, ReadBlockSize)
            EnableEvents = "No error occurred."
        Catch ex As Exception
            ErrString = ex.Message
            ErrCategory = "Error- Could not register callback. "
            EnableEvents = ErrCategory & ErrString
        End Try

    End Function

    Public Function DisableEvents(ByVal TypeOfCallback As _
    Integer) As String Implements MBDTest.DisableEvents

        Dim ErrString, ErrCategory As String
        Dim CallbackType As MeasurementComputing.DAQFlex.CallbackType

        Try
            'MyDevice.RegisterCallback(AddressOf AInScanCallback, 256, False)
            'DisableEvents = "Events disabled."
            CallbackType = TypeOfCallback
            MyDevice.DisableCallback(CallbackType)
            ErrCategory = "No error occurred."
            DisableEvents = ErrCategory
        Catch ex As Exception
            ErrString = ex.Message
            ErrCategory = "Error- Could not unregister callback. "
            DisableEvents = ErrCategory & ErrString
        End Try

    End Function

    Public Function ConvertErrorCode(ByVal ErrorCode As Integer) As String Implements MBDTest.ConvertErrorCode

        Dim ErrString As String, ErrCategory As String

        Try
            ErrString = MyDevice.GetErrorMessage(ErrorCode)
        Catch ex As Exception
            ErrCategory = "Error- " & ex.Message
            ErrString = ErrCategory & " (Error converting code to string.)"
        End Try
        ConvertErrorCode = ErrString

    End Function

    Public Function StartBackgroundScan(ByVal FunctionType As String, ByVal ReadBlockSize As Integer) As String Implements MBDTest.StartBackgroundScan

        Dim Response As DaqResponse
        Dim ErrString As String, ErrCategory As String, Msg As String
        Dim MBDWarning As Boolean

        Try
            Msg = FunctionType & ":SAMPLES=0"
            Response = MyDevice.SendMessage(Msg)
            Msg = FunctionType & ":START"
            Response = MyDevice.SendMessage(Msg)
            StartBackgroundScan = Response.ToString
        Catch ex As Exception
            ErrString = ex.Message
            MBDWarning = ErrString.Contains("The device does not support the command sent")
            If MBDWarning Then
                ErrCategory = "Warning- "
            Else
                ErrCategory = "Error- "
            End If
            StartBackgroundScan = ErrCategory & ErrString & " (Error starting background scan.)"
            Exit Function
        End Try
        Try
            'MyDevice.RegisterCallback(AddressOf AInScanCallback, ReadBlockSize, True)
            ErrCategory = "Error- Callback not yet implemented. "
            StartBackgroundScan = ErrCategory
        Catch ex As Exception
            ErrString = ex.Message
            ErrCategory = "Error- Could not register callback. "
            StartBackgroundScan = ErrCategory & ErrString
        End Try

    End Function

    Public Function StopBackgroundScan(ByVal FunctionType As String) As String Implements MBDTest.StopBackgroundScan

        Dim Response As DaqResponse
        Dim ErrString As String, ErrCategory As String, Msg As String
        Dim MBDWarning As Boolean

        Try
            Msg = FunctionType & ":STOP"
            Response = MyDevice.SendMessage(Msg)
            StopBackgroundScan = Response.ToString
        Catch ex As Exception
            ErrString = ex.Message
            MBDWarning = ErrString.Contains("The device does not support the command sent")
            If MBDWarning Then
                ErrCategory = "Warning- "
            Else
                ErrCategory = "Error- "
            End If
            StopBackgroundScan = ErrCategory & ErrString & " (Error stopping background scan.)"
            Exit Function
        End Try
        Try
            'MyDevice.RegisterCallback(AddressOf AInScanCallback, 256, False)
            ErrCategory = "Error- Callback is not yet implemented."
            StopBackgroundScan = ErrCategory
        Catch ex As Exception
            ErrString = ex.Message
            ErrCategory = "Error- Could not unregister callback. "
            StopBackgroundScan = ErrCategory & ErrString
        End Try

    End Function

    Public Function GetSupportedMessages(ByVal Component As String) As String Implements MBDTest.GetSupportedMessages

        Dim MessageList As New System.Collections.Generic.List(Of String)
        Dim i As Integer, ReturnString As String

        ReturnString = String.Empty
        Try
            MessageList = MyDevice.GetSupportedMessages(Component)
        Catch ex As Exception
            MessageList.Add("Error- " & ex.Message)
        End Try
        If Not IsNothing(MessageList) Then
            For i = 0 To MessageList.Count - 1
                ReturnString = ReturnString & MessageList(i) & "|"
            Next
        End If
        GetSupportedMessages = ReturnString

    End Function

    Public ReadOnly Property DeviceID() As String Implements MBDTest.DeviceID

        Get
            Return msDeviceID
        End Get

    End Property

    Public ReadOnly Property NumberOfDevices() As Integer Implements MBDTest.NumberOfDevices

        Get
            Return mnNumDevices
        End Get

    End Property

    Private Sub AInScanCallback(ByVal ErrorCode As ErrorCodes, _
    ByVal EventType As MeasurementComputing.DAQFlex.CallbackType, ByVal callbackData As Object)

        Dim IntCode, EventCode As Integer

        If ErrorCode = ErrorCodes.NoErrors Then
            EventCode = EventType
            RaiseEvent DataAvailable(msDeviceID, EventCode, callbackData)
        Else
            IntCode = ErrorCode
            RaiseEvent ErrorAvailable(msDeviceID, IntCode)
        End If

    End Sub

End Class
