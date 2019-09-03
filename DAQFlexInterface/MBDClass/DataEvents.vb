Imports System.Runtime.InteropServices

Public Interface DataEvents

    <Guid("752CEC9E-A86D-4295-B1BE-C5CB7EC054BE"), _
    InterfaceType(ComInterfaceType.InterfaceIsIDispatch)> _
    Public Interface IDataEvents

        <DispId(1)> _
        Sub DataAvailable(ByVal DeviceID As String, _
        ByVal EventType As Integer, ByRef callbackData As Object)

        <DispId(2)> _
        Sub ErrorAvailable(ByVal DeviceID As String, ByVal ErrorCode As Integer)

    End Interface

End Interface

