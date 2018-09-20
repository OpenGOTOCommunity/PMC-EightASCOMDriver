Imports ASCOM.DeviceInterface

'
' The Rate class implements IRate, and is used to hold values
' for AxisRates. You do not need to change this class.
'
' The Guid attribute sets the CLSID for ASCOM.ASCOM.ES_PMC8.Telescope.Rate
' The ClassInterface/None addribute prevents an empty interface called
' _Rate from being created and used as the [default] interface
'
<Guid("bf5f946d-3c9c-4963-bf85-c1d066b4224a")> _
<ClassInterface(ClassInterfaceType.None)> _
<ComVisible(True)> _
Public Class Rate
    '		================
    Implements IRate
    '		================

    Private m_dMaximum As Double = 0
    Private m_dMinimum As Double = 0

    '
    ' Default constructor - Internal prevents public creation
    ' of instances. These are values for AxisRates.
    '
    Friend Sub New(ByVal Minimum As Double, ByVal Maximum As Double)
        m_dMaximum = Maximum
        m_dMinimum = Minimum
    End Sub

#Region "IRate Members"

    Public Sub Dispose() Implements IRate.Dispose
        'Throw New System.NotImplementedException()
        'Throw New ASCOM.MethodNotImplementedException()
    End Sub

    Public Property Maximum() As Double Implements IRate.Maximum
        Get
            Return m_dMaximum
        End Get
        Set(ByVal Value As Double)
            m_dMaximum = Value
        End Set
    End Property

    Public Property Minimum() As Double Implements IRate.Minimum
        Get
            Return m_dMinimum
        End Get
        Set(ByVal Value As Double)
            m_dMinimum = Value
        End Set
    End Property

#End Region


End Class

'
' AxisRates is a strongly-typed collection that must be enumerable by
' both COM and .NET. The IAxisRates and IEnumerable interfaces provide
' this polymorphism. 
'
' The Guid attribute sets the CLSID for ASCOM.ASCOM.ES_PMC8.Telescope.AxisRates
' The ClassInterface/None addribute prevents an empty interface called
' _AxisRates from being created and used as the [default] interface
'
<Guid("93e19a3e-c841-46f0-a07a-9606729faaf7")> _
<ClassInterface(ClassInterfaceType.None)> _
<ComVisible(True)> _
Public Class AxisRates
    '		======================
    Implements IAxisRates
    Implements IEnumerable
    '		======================

    Private m_Axis As TelescopeAxes
    Private m_Rates(-1) As Rate         ' Empty array, but an array nonetheless

    '
    ' Constructor - Friend prevents public creation
    ' of instances. Returned by Telescope.AxisRates.
    '
    Friend Sub New(ByVal Axis As TelescopeAxes)
        m_Axis = Axis
        '
        ' This collection must hold zero or more Rate objects describing the 
        ' rates of motion ranges for the Telescope.MoveAxis() method
        ' that are supported by your driver. It is OK to leave this 
        ' array empty, indicating that MoveAxis() is not supported.
        '
        ' Note that we are constructing a rate array for the axis passed
        ' to the constructor. Thus we switch() below, and each case should 
        ' initialize the array for the rate for the selected axis.
        '
        Select Case Axis
            Case TelescopeAxes.axisPrimary
                ' TODO Initialize this array with any Primary axis rates that your driver may provide
                ' Example: ReDim m_Rates(2)
                '          m_Rates(0) = New Rate(10.0, 35.0)
                '          m_Rates(1) = New Rate(60.1, 120.0)
                ReDim m_Rates(4)
                m_Rates(0) = New Rate(2.5#, 2.5#)  'SLEW 2.5 deg/sec 600x Sidereal Rate
                m_Rates(1) = New Rate(0.2083333, 0.2083333)  'MOVE 50x Sidereal Rate
                m_Rates(2) = New Rate(0.0208333, 0.0208333)  'CENTER 5x Sidereal Rate
                m_Rates(3) = New Rate(0.0025, 0.0025)  'GUIDE 0.6x Sidereal Rate
                m_Rates(4) = New Rate(0.0000, 0.0000) 'stop
                Exit Sub
            Case TelescopeAxes.axisSecondary
                ' TODO Initialize this array with any Secondary axis rates that your driver may provide
                ReDim m_Rates(4)
                m_Rates(0) = New Rate(2.5#, 2.5#)  'SLEW 2.5 deg/sec 600x Sidereal Rate
                m_Rates(1) = New Rate(0.2083333, 0.2083333)  'MOVE 50x Sidereal Rate
                m_Rates(2) = New Rate(0.0208333, 0.0208333)  'CENTER 5x Sidereal Rate
                m_Rates(3) = New Rate(0.0025, 0.0025)  'GUIDE 0.6x Sidereal Rate
                m_Rates(4) = New Rate(0.0000, 0.0000) 'stop
                Exit Sub
            Case TelescopeAxes.axisTertiary
                ' TODO Initialize this array with any Tertiary axis rates that your driver may provide
                Exit Sub
        End Select
    End Sub

#Region "IAxisRates Members"

    Public ReadOnly Property Count() As Integer Implements IAxisRates.Count
        Get
            Return m_Rates.Length
        End Get
    End Property

    Public Sub Dispose() Implements IAxisRates.Dispose
        'Throw New System.NotImplementedException()
        'Throw New ASCOM.MethodNotImplementedException()
    End Sub

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator, IAxisRates.GetEnumerator
        Return m_Rates.GetEnumerator()
    End Function

    Default Public ReadOnly Property Item(ByVal Index As Integer) As IRate Implements IAxisRates.Item
        Get
            Return CType(m_Rates(Index - 1), IRate)    ' 1-based
        End Get
    End Property

#End Region

End Class

'
' TrackingRates is a strongly-typed collection that must be enumerable by
' both COM and .NET. The ITrackingRates and IEnumerable interfaces provide
' this polymorphism. 
'
' The Guid attribute sets the CLSID for ASCOM.ASCOM.ES_PMC8.Telescope.TrackingRates
' The ClassInterface/None addribute prevents an empty interface called
' _TrackingRates from being created and used as the [default] interface
'
<Guid("041155dd-d194-416d-aae4-bda96c7c3549")> _
<ClassInterface(ClassInterfaceType.None)> _
<ComVisible(True)> _
Public Class TrackingRates
    '		=========================
    Implements ITrackingRates
    Implements IEnumerable
    Implements IEnumerator
    '		=========================

    Private m_TrackingRates(-1) As DriveRates           ' Empty array, but an array nonetheless
    Private pos As Integer = -1
    '
    ' Default constructor - Friend prevents public creation
    ' of instances. Returned by Telescope.AxisRates.
    '
    Friend Sub New()
        '
        ' This array must hold ONE or more DriveRates values, indicating
        ' the tracking rates supported by your telescope. The one value
        ' (tracking rate) that MUST be supported is driveSidereal!
        '
        ' TODO Initialize this array with any additional tracking rates that your driver may provide
        ' Example: ReDim m_TrackingRates(1)
        '          m_TrackingRates(0) = DriveRates.driveSidereal
        '          m_TrackingRates(1) = DriveRates.driveLunar
        '
        ReDim m_TrackingRates(3)
        m_TrackingRates(0) = DriveRates.driveSidereal
        m_TrackingRates(1) = DriveRates.driveLunar
        m_TrackingRates(2) = DriveRates.driveSolar
        m_TrackingRates(3) = DriveRates.driveKing

    End Sub

#Region "ITrackingRates Members"

    Public ReadOnly Property Count() As Integer Implements ITrackingRates.Count
        Get
            Return m_TrackingRates.Length
        End Get
    End Property

    Public Function GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator, ITrackingRates.GetEnumerator
        Return m_TrackingRates.GetEnumerator()
    End Function

    Public Sub Dispose() Implements ITrackingRates.Dispose
        'Throw New System.NotImplementedException()
        'Throw New ASCOM.MethodNotImplementedException()
    End Sub

    Default Public ReadOnly Property Item(ByVal Index As Integer) As DriveRates Implements ITrackingRates.Item
        Get
            Return m_TrackingRates(Index - 1)  ' 1-based
        End Get
    End Property

#End Region

#Region "IEnumerator members"
    Public Function MoveNext() As Boolean Implements IEnumerator.MoveNext
        pos = pos + 1
        If pos >= m_TrackingRates.Length Then
            Return False
        End If
        Return True
    End Function

    Public ReadOnly Property Current As Object Implements IEnumerator.Current
        Get
            If pos < 0 Or pos >= m_TrackingRates.Length Then
                Throw New InvalidOperationException()
            End If
            Return m_TrackingRates(pos)
        End Get
    End Property

    Public Sub Reset() Implements IEnumerator.Reset
        pos = -1
    End Sub
#End Region

End Class
