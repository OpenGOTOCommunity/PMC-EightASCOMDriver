'tabs=4
' --------------------------------------------------------------------------------
'
' ASCOM Telescope driver for ES_PMC8
'
' Description:	This is the main telescope driver for the Explore Scientific
'				Precision Motion Controller Eight (PMC-Eight). This driver
'               is copyright 2016-2017 Explore Scientific, LLC.
'
' Implements:	ASCOM Telescope interface version: 6.0
' Author:		(GRH) Gerald R. Hubbell <jrh@explorescientific.com>
'               Director Electrical Engineering, Explore Scientific, LLC.
'
' Edit Log:
'
' Date			Who	Vers	Description
' -----------	---	-----	-------------------------------------------------------
' 26-AUG-2015	GRH	1.0.0	Initial edit, from Telescope template
' 30-AUG-2015   GRH 1.0.1   Added code for connect routine
' 19-MAR-2016   GRH 1.0.2   Code added for connecting to serial port
' 19-MAY-2016   GRH 1.0.3   Added calculations for RightAscension and Declination
' 21-MAY-2016   GRH 1.0.4   Added calculations for converting RA,DEC,ALT,AZ to motor
'                           counts.
' 24-MAY-2016   GRH 1.0.5   Added and Tested base functionality for Slewing, Moving,
'                           and Syncing controller. Tested Serial port connection
' 06-JUN-2016   GRH 1.0.6   Added code to implement functionality and Tested Wireless
'                           TCP/IP connection to controller, Further performance testing
' 10-JUN-2016   GRH 1.0.7   Corrected minor issues with driver. Implemented AbortSlew
'                           turned on tracking after slew, other misc fixes.
'                           Added Home functioinality
' 12-JUN-2016   GRH 1.0.8   Added PierSide functionality and calculation and SetPark
'                           functionality. Tested functionality SideOfPier
' 28-JUN-2016   GRH 1.0.9   Continued to test and fix coordinate calculations.
' 29-JUN-2016   GRH 1.0.10  Added Refraction adjustments to coordinate calculations,
'                           and stored the values used:Site Ambient Temperature, and
'                           Apply Refraction Correction in the Profile object.
' 08-JUL-2016   GRH 1.0.11  Further testing SideOfPier and RA/Motor conversion.
' 25-JUL-2015   GRH 1.0.12  Simplified SideOfPier and MotorCount conversion commands
'                           and tested all the routines to make sure they pass all
'                           conformance testing including side of pier tests and 
'                           pier flip from pierWest to pierEast. Corrected problem
'                           with AbortSlew command. 
' 25-JUL-2016   GRH 1.2.0   RELEASE VERSION for BETA TESTING in FIELD
' 04-DEC-2016   GRH 1.2.1   Added Alignment and Ephemeris Routines and dialog box
' ---------------------------------------------------------------------------------
'
' Your driver's ID is ASCOM.ES_PMC8.Telescope
'
' The Guid attribute sets the CLSID for ASCOM.DeviceName.Telescope
' The ClassInterface/None addribute prevents an empty interface called
' _Telescope from being created and used as the [default] interface
'

' This definition is used to select code that's only applicable for one device type
#Const Device = "Telescope"

Imports ASCOM
Imports ASCOM.Astrometry
Imports ASCOM.Astrometry.AstroUtils
Imports ASCOM.DeviceInterface
Imports ASCOM.Utilities
Imports ASCOM.Astrometry.Transform
Imports MathNet.Numerics.LinearAlgebra

Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Net
Imports System.Net.Sockets

<Guid("5276d38b-2048-4b9c-8761-c9a94d6aa372")> _
<ClassInterface(ClassInterfaceType.None)> _
<ComVisible(True)>
Public Class Telescope
    ' The Guid attribute sets the CLSID for ASCOM.ES_PMC8.Telescope
    ' The ClassInterface/None addribute prevents an empty interface called
    ' _ES_PMC8 from being created and used as the [default] interface

    ' TODO Replace the not implemented exceptions with code to implement the function or
    ' throw the appropriate ASCOM exception.
    '
    Implements ITelescopeV3
    '
    ' Driver ID and descriptive string that shows in the Chooser
    '
    Friend Shared driverID As String = "ASCOM.ES_PMC8.Telescope"
    Friend Shared driverDescription As String = "ES_PMC8 Telescope"
    Friend Shared comPortProfileName As String = "COM Port" 'Constants used for Profile persistence
    Friend Shared comSpeedProfileName As String = "COM Speed"
    Friend Shared traceStateProfileName As String = "Trace Level"
    Friend Shared IPAddressProfileName As String = "IP Address"
    Friend Shared IPPortProfileName As String = "IP Port"
    Friend Shared WirelessEnabledProfileName As String = "Wireless Enabled"
    Friend Shared WirelessProtocolProfileName As String = "Wireless Protocol"
    Friend Shared MountProfileName As String = "Mount Type"
    Friend Shared MountRACountsProfileName As String = "Total RA Counts"
    Friend Shared MountDECCountsProfileName As String = "Total DEC Counts"
    Friend Shared SiteLocationProfileName As String = "Site Location"
    Friend Shared SiteElevationProfileName As String = "Site Elevation meters"
    Friend Shared SiteLatitudeProfileName As String = "Site Latitude"
    Friend Shared SiteLongitudeProfileName As String = "Site Longitude"
    Friend Shared SiteAmbientTemperatureProfileName As String = "Site Ambient Temperature"
    Friend Shared ParkRAPositionProfileName As String = "RA Park Position"
    Friend Shared ParkDECPositionProfileName As String = "DEC Park Position"
    Friend Shared ApplyRefractionCorrectionProfileName As String = "RefractionApplied"

    Friend Shared comPortDefault As String = "COM3"
    Friend Shared comSpeedDefault As String = "115200"
    Friend Shared traceStateDefault As String = "False"
    Friend Shared IPAddressDefault As String = "192.168.47.1"
    Friend Shared IPPortDefault As String = "54372"
    Friend Shared WirelessEnabledDefault As String = "True"
    Friend Shared WirelessProtocolDefault As String = "TCP"
    Friend Shared MountDefault As String = "Losmandy G-11"
    Friend Shared MountRACountsDefault As String = "4608000" 'G-11
    Friend Shared MountDECCountsDefault As String = "4608000" 'G-11
    Friend Shared SiteLocationDefault As String = "Explore Scientific HQ"
    Friend Shared SiteElevationDefault As String = "403.0"
    Friend Shared SiteLatitudeDefault As String = "36.18063"
    Friend Shared SiteLongitudeDefault As String = "-94.18838"
    Friend Shared SiteAmbientTemperatureDefault As String = "59.00000"
    Friend Shared ParkRAPositionDefault As String = "0"
    Friend Shared ParkDECPositionDefault As String = "0"
    Friend Shared ApplyRefractionCorrectionDefault As String = "False"

    Friend Shared comPort As String ' Variables to hold the currrent device configuration
    Friend Shared comSpeed As String
    Friend Shared traceState As Boolean
    Friend Shared IPAddress As String
    Friend Shared IPPort As String
    Friend Shared WirelessEnabled As Boolean
    Friend Shared WirelessProtocol As String
    Friend Shared Mount As String
    Friend Shared MountRACCounts As Long
    Friend Shared MountDECCounts As Long
    Friend Shared SiteLocation As String
    Friend Shared SiteElevationValue As Double
    Friend Shared SiteLongitudeValue As Double
    Friend Shared SiteLatitudeValue As Double
    Friend Shared SiteAmbientTemperatureValue As Double
    Friend Shared ParkRAPosition As Int32
    Friend Shared ParkDECPosition As Int32
    Friend Shared ApplyRefractionCorrection As Boolean

    Private connectedState As Boolean ' Private variable to hold the connected state
    Private utilities As Util ' Private variable to hold an ASCOM Utilities object
    Private astroUtilities As AstroUtils ' Private variable to hold an AstroUtils object to provide the Range method
    Private TL As TraceLogger ' Private variable to hold the trace logger object (creates a diagnostic log file with information that you specify)

    Private objSerial As ASCOM.Utilities.Serial 'Serial Port object
    Private objUDPNetwork As System.Net.Sockets.UdpClient 'UDP network object
    'Private objUDPNetwork_S As System.Net.Sockets.UdpClient 'UDP Network Receiver object
    Private objTCPNetwork As System.Net.Sockets.TcpClient 'TCP network object
    Private objTransform As New ASCOM.Astrometry.Transform.Transform 'Transform calculations
    Private SerMutex As New System.Threading.Mutex
    Private PrevRA As Double
    Private PrevDEC As Double
    Private PrevRAMotor As Int32
    Private PrevDECMotor As Int32

    Private j2000 As New DateTime
    Private deltaTime As TimeSpan
    Private LMSTtot As Double
    Private LMST As Double
    Private di As Double
    Private RATracking As Boolean = False
    'Private RATarget As Double = 12.0
    'Private DECTarget As Double = 90.0
    Private RATarget As Double
    Private DECTarget As Double
    Private ParkStatus As Boolean = True
    Private HomeStatus As Boolean = True
    Private RADirection As String = "CCW"
    Private DECDirection As String = "CW"
    Private RATargetSet As Boolean = False
    Private DECTargetSet As Boolean = False
    Private minSlewRate As Int32 = 55
    Private WpE_Normal As Boolean
    Private EpW_Normal As Boolean
    Private WpE_TtP As Boolean
    Private EpW_TtP As Boolean
    Private SoP As Integer
    Private FlipMount As Boolean
    Private MountTrackingRate As DriveRates = DriveRates.driveSidereal
    Private MountPierSide As PierSide
    Private MountRightAscensionRate As Double
    '
    ' Constructor - Must be public for COM registration!
    '
    Public Sub New()

        ReadProfile() ' Read device configuration from the ASCOM Profile store
        TL = New TraceLogger("", "ES_PMC8")
        TL.Enabled = traceState
        TL.LogMessage("Telescope", "Starting initialization")
        connectedState = False ' Initialise connected to false
        utilities = New Util() ' Initialise util object
        astroUtilities = New AstroUtils 'Initialise new astro utiliites object
        objTransform.SiteLatitude = Me.SiteLatitude
        objTransform.SiteLongitude = Me.SiteLongitude
        objTransform.SiteElevation = Me.SiteElevation
        objTransform.SiteTemperature = SiteAmbientTemperatureValue

        'TODO: Implement your additional construction here

        TL.LogMessage("Telescope", "Completed initialisation")
    End Sub

    '
    ' PUBLIC COM INTERFACE ITelescopeV3 IMPLEMENTATION
    '

#Region "Common properties and methods"
    ''' <summary>
    ''' Displays the Setup Dialog form.
    ''' If the user clicks the OK button to dismiss the form, then
    ''' the new settings are saved, otherwise the old values are reloaded.
    ''' THIS IS THE ONLY PLACE WHERE SHOWING USER INTERFACE IS ALLOWED!
    ''' </summary>


    Public Sub SetupDialog() Implements ITelescopeV3.SetupDialog
        ' consider only showing the setup dialog if not connected
        ' or call a different dialog if connected
        If IsConnected Then
            System.Windows.Forms.MessageBox.Show("Already connected, just press OK")
        End If

        Using F As SetupDialogForm = New SetupDialogForm()
            Dim result As System.Windows.Forms.DialogResult = F.ShowDialog()
            If result = DialogResult.OK Then
                WriteProfile() ' Persist device configuration values to the ASCOM Profile store
            End If
        End Using
    End Sub

    Public ReadOnly Property SupportedActions() As ArrayList Implements ITelescopeV3.SupportedActions
        Get
            Dim myActions As New ArrayList
            'List of actions,  may not be totally complete
            myActions.Add("CanPark")
            myActions.Add("CanSetTracking")
            myActions.Add("CanSlew")
            myActions.Add("CanSlewAltAz")
            myActions.Add("CanSlewAsync")
            myActions.Add("CanSlewAltAzAsync")
            myActions.Add("CanSync")
            myActions.Add("CanSyncAltAz")
            myActions.Add("CanUnpark")
            myActions.Add("CanMoveAxis")
            myActions.Add("DoesRefraction")
            myActions.Add("Can...")

            TL.LogMessage("SupportedActions Get", "Returning array list")
            Return myActions
        End Get
    End Property

    Public Function Action(ByVal ActionName As String, ByVal ActionParameters As String) As String Implements ITelescopeV3.Action
        Throw New ASCOM.ActionNotImplementedException("Action " & ActionName & " is not supported by this driver")
    End Function

    Public Sub CommandBlind(ByVal Command As String, Optional ByVal Raw As Boolean = False) Implements ITelescopeV3.CommandBlind
        CheckConnected("CommandBlind")
        ' Call CommandString and return as soon as it finishes
        Me.CommandString(Command, Raw)
        ' or
        Throw New ASCOM.MethodNotImplementedException("CommandBlind")
    End Sub

    Public Function CommandBool(ByVal Command As String, Optional ByVal Raw As Boolean = False) As Boolean _
        Implements ITelescopeV3.CommandBool
        CheckConnected("CommandBool")
        Dim ret As String = CommandString(Command, Raw)
        ' TODO decode the return string and return true or false
        ' or
        Throw New MethodNotImplementedException("CommandBool")
    End Function

    Public Function CommandString(ByVal Command As String, Optional ByVal Raw As Boolean = False) As String _
        Implements ITelescopeV3.CommandString
        'Dim objUDPNetwork_R As New System.Net.Sockets.UdpClient(CInt(UDPPort))
        'Dim RemoteIpEndPoint As New IPEndPoint(System.Net.IPAddress.Any, 0)
        Dim sendBytes As [Byte]()
        Dim receiveBytes As [Byte]()
        Dim receiveString As String
        Dim cmdString As String

        CheckConnected("CommandString")
        cmdString = Command
        ' it's a good idea to put all the low level communication with the device here,
        ' then all communication calls this function
        ' you need something to ensure that only one command is in progress at a time
        If IsConnected Then
            If Not WirelessEnabled Then
                objSerial.Transmit(Command)
                'ES Command Language Terminator String ! (SHRIEK)
                receiveString = objSerial.ReceiveTerminated("!")
                cmdString = receiveString
                objSerial.ClearBuffers() 'clear out waiting for next command

            ElseIf WirelessEnabled Then
                'Dim objUDPNetwork_R As New System.Net.Sockets.UdpClient(CInt(UDPPort))
                'Dim RemoteIpEndPoint As New IPEndPoint(System.Net.IPAddress.Any, 0)
                Try
                    '-----------------------------------------------------------------------
                    ' UDP Connection
                    'Dim objUDPNetwork_R As New UdpClient(CInt(IPPort))
                    'objUDPNetwork_R.Connect(IPAddress, CInt(IPPort))
                    'sendBytes = Encoding.ASCII.GetBytes(Command)
                    'objUDPNetwork_R.Send(sendBytes, sendBytes.Length)
                    'Dim RemoteIpEndPoint As New IPEndPoint(System.Net.IPAddress.Any, 0)
                    'receiveBytes = objUDPNetwork_R.Receive(RemoteIpEndPoint)
                    'cmdString = Encoding.ASCII.GetString(receiveBytes)
                    'objUDPNetwork_R.Close()
                    'objUDPNetwork_R = Nothing
                    '-----------------------------------------------------------------------
                    'TCP Connection
                    Dim objTCPNetwork As New TcpClient(IPAddress, CInt(IPPort))
                    Dim stream As NetworkStream = objTCPNetwork.GetStream()
                    sendBytes = System.Text.Encoding.ASCII.GetBytes(Command)
                    stream.Write(sendBytes, 0, sendBytes.Length)
                    receiveBytes = New [Byte](256) {}
                    Dim responseData As [String] = [String].Empty
                    Dim bytes As Int32 = stream.Read(receiveBytes, 0, receiveBytes.Length)
                    responseData = System.Text.Encoding.ASCII.GetString(receiveBytes, 0, bytes)

                    'stream.Write(sendBytes, 0, sendBytes.Length)
                    'receiveBytes = New [Byte](256) {}
                    responseData = [String].Empty
                    bytes = stream.Read(receiveBytes, 0, receiveBytes.Length)
                    responseData = System.Text.Encoding.ASCII.GetString(receiveBytes, 0, bytes)

                    cmdString = responseData
                    stream.Close()
                    objTCPNetwork.Close()
                    objTCPNetwork = Nothing
                    '-----------------------------------------------------------------------
                Catch e As Sockets.SocketException
                    Console.WriteLine(e.ToString())
                End Try

            End If
        ElseIf Not IsConnected Then
            Throw New ASCOM.MethodNotImplementedException("CommandString")
        End If
        Return cmdString

    End Function

    Public Property Connected() As Boolean Implements ITelescopeV3.Connected

        Get
            TL.LogMessage("Connected Get", IsConnected.ToString())
            Return IsConnected
        End Get
        Set(value As Boolean)
            'Declare local parameters
            Dim receiveBytes As [Byte]()
            Dim sendBytes As [Byte]()

            TL.LogMessage("Connected Set", value.ToString())
            If value = IsConnected Then
                Return
            End If

            If value Then
                'connectedState = True
                ' TODO connect to the device
                'Set serial parameters and Connect to comport ---------------------
                If Not WirelessEnabled Then
                    Try
                        objSerial = New ASCOM.Utilities.Serial
                        TL.LogMessage("Connected Set", "Connecting to port " + comPort)
                        objSerial.Port = Right(comPort, 1)
                        objSerial.Speed = comSpeed
                        objSerial.ReceiveTimeout = 1
                        objSerial.DataBits = 8
                        objSerial.StopBits = SerialStopBits.One
                        objSerial.Parity = SerialParity.None
                        objSerial.Handshake = SerialHandshake.None
                        objSerial.RTSEnable = False
                        objSerial.DTREnable = False
                        objSerial.Connected = True
                        objSerial.ClearBuffers()
                        connectedState = True
                    Catch e As Sockets.SocketException
                        Console.WriteLine(e.ToString())
                    End Try
                    '                    connectedState = True
                Else
                    'Set Network parameters and Connect to TCPPort ---------------------
                    'TODO - put wireless connection here
                    Try
                        TL.LogMessage("Connected Set", "Connecting to IP Address " + IPAddress)
                        Dim objTCPNetwork As New TcpClient(IPAddress, CInt(IPPort))
                        Dim stream As NetworkStream = objTCPNetwork.GetStream()

                        sendBytes = System.Text.Encoding.ASCII.GetBytes("ESGp0!")

                        stream.Write(sendBytes, 0, sendBytes.Length)
                        receiveBytes = New [Byte](256) {}
                        Dim responseData As [String] = [String].Empty
                        Dim bytes As Int32 = stream.Read(receiveBytes, 0, receiveBytes.Length)
                        responseData = System.Text.Encoding.ASCII.GetString(receiveBytes, 0, bytes)

                        If responseData = "*HELLO*" Then
                            connectedState = True
                        End If

                        stream.Close()
                        objTCPNetwork.Close()
                        objTCPNetwork = Nothing
                    Catch e As Sockets.SocketException
                        Console.WriteLine(e.ToString())
                    End Try
                End If
                '------------------------------------------------
            Else
                '                connectedState = False
                TL.LogMessage("Connected Set", "Disconnecting from port " + comPort)
                ' TODO disconnect from the device
                '------------------------------------------------
                If Not WirelessEnabled Then
                    objSerial.ClearBuffers()
                    objSerial.Connected = False
                    objSerial.Dispose()
                    objSerial = Nothing
                    connectedState = False
                Else
                    'objTCPNetwork.Close()
                    'objTCPNetwork = Nothing


                    'objUDPNetwork.Close()
                    'objUDPNetwork_R.Close()
                    'objUDPNetwork = Nothing
                    'objUDPNetwork_R = Nothing
                    connectedState = False
                End If
                '------------------------------------------------
            End If
        End Set
    End Property

    Public ReadOnly Property Description As String Implements ITelescopeV3.Description
        Get
            ' this pattern seems to be needed to allow a public property to return a private field
            Dim d As String = driverDescription
            TL.LogMessage("Description Get", d)
            Return d
        End Get
    End Property

    Public ReadOnly Property DriverInfo As String Implements ITelescopeV3.DriverInfo
        Get
            Dim m_version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            ' TODO customise this driver description
            Dim s_driverInfo As String = "Explore Scientific PMC-Eight Mount Controller ASCOM Driver. Developed by GRHubbell. Contact Explore Scientific at www.explorescientificusa.com . Version: " + m_version.Major.ToString() + "." + m_version.Minor.ToString()
            TL.LogMessage("DriverInfo Get", s_driverInfo)
            Return s_driverInfo
        End Get
    End Property

    Public ReadOnly Property DriverVersion() As String Implements ITelescopeV3.DriverVersion
        Get
            ' Get our own assembly and report its version number
            TL.LogMessage("DriverVersion Get", Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2))
            Return Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString(2)
        End Get
    End Property

    Public ReadOnly Property InterfaceVersion() As Short Implements ITelescopeV3.InterfaceVersion
        Get
            TL.LogMessage("InterfaceVersion Get", "3")
            Return 3
        End Get
    End Property

    Public ReadOnly Property Name As String Implements ITelescopeV3.Name
        Get
            Dim s_name As String = "Explore Scientific PMC-Eight ASCOM Driver"
            TL.LogMessage("Name Get", s_name)
            Return s_name
        End Get
    End Property

    Public Sub Dispose() Implements ITelescopeV3.Dispose
        ' Clean up the tracelogger and util objects
        TL.Enabled = False
        TL.Dispose()
        TL = Nothing
        utilities.Dispose()
        utilities = Nothing
        astroUtilities.Dispose()
        astroUtilities = Nothing
    End Sub

#End Region

#Region "ITelescope Implementation"
    Public Sub AbortSlew() Implements ITelescopeV3.AbortSlew
        Dim abortCommand As String
        Try
            If Not AtPark Then
                'Set right ascension tracking Rate to 0 (zero)
                abortCommand = "ESSr00000!"
                SerMutex.WaitOne()
                CommandString(abortCommand)
                SerMutex.ReleaseMutex()
                'Set declination Rate to 0 (zero)
                abortCommand = "ESSr10000!"
                SerMutex.WaitOne()
                CommandString(abortCommand)
                SerMutex.ReleaseMutex()
                If RATracking = True Then
                    Tracking = True
                Else
                    Tracking = False
                End If
            Else
                TL.LogMessage("AbortSlew", "Parked!")
                Throw New ASCOM.ParkedException("AbortSlew")
            End If
        Catch ex As Exception
            TL.LogMessage("AbortSlew", "Invalid Operation")
            Throw New ASCOM.InvalidOperationException("AbortSlew")
        End Try

    End Sub

    Public ReadOnly Property AlignmentMode() As AlignmentModes Implements ITelescopeV3.AlignmentMode
        Get
            TL.LogMessage("AlignmentMode Get", "Implemented")
            AlignmentMode = AlignmentModes.algGermanPolar
            Return AlignmentMode

            'Throw New ASCOM.PropertyNotImplementedException("AlignmentMode", False)
        End Get
    End Property

    Public ReadOnly Property Altitude() As Double Implements ITelescopeV3.Altitude
        Get
            Dim Altitude__1 As Double
            'Dim AltReceived As String = "0"

            If IsConnected Then
                'objTransform.SetApparent(Me.RightAscension, Me.Declination)
                'objTransform = New ASCOM.Astrometry.Transform.Transform
                Me.objTransform.SetTopocentric(Me.RightAscension, Me.Declination)

                Altitude__1 = objTransform.ElevationTopocentric


                'Altitude__1 = 90.0 - Me.SiteLatitude + Me.Declination
            ElseIf Not IsConnected Then
                Throw New ASCOM.MethodNotImplementedException("Altitude")
            End If
            TL.LogMessage("Altitude", "Get - " & Altitude__1)
            Return Altitude__1
        End Get
    End Property

    Public ReadOnly Property ApertureArea() As Double Implements ITelescopeV3.ApertureArea
        Get
            TL.LogMessage("ApertureArea Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("ApertureArea", False)
        End Get
    End Property

    Public ReadOnly Property ApertureDiameter() As Double Implements ITelescopeV3.ApertureDiameter
        Get
            TL.LogMessage("ApertureDiameter Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("ApertureDiameter", False)
        End Get
    End Property

    Public ReadOnly Property AtHome() As Boolean Implements ITelescopeV3.AtHome
        Get
            TL.LogMessage("AtHome", "Get - " & HomeStatus.ToString())
            Return HomeStatus
        End Get
    End Property

    Public ReadOnly Property AtPark() As Boolean Implements ITelescopeV3.AtPark
        Get
            TL.LogMessage("AtPark", "Get - " & ParkStatus.ToString())
            Return ParkStatus
        End Get
    End Property

    Public Function AxisRates(Axis As TelescopeAxes) As IAxisRates Implements ITelescopeV3.AxisRates
        TL.LogMessage("AxisRates", "Get - " & Axis.ToString())
        Return New AxisRates(Axis)
    End Function

    Public ReadOnly Property Azimuth() As Double Implements ITelescopeV3.Azimuth
        Get
            Dim Azimuth__1 As Double
            Dim AzReceived As String = "0"

            If IsConnected Then
                'Azimuth__1 = Me.RightAscension
                'objTransform.SetApparent(Me.RightAscension, Me.Declination)
                'objTransform = New ASCOM.Astrometry.Transform.Transform
                Me.objTransform.SetTopocentric(Me.RightAscension, Me.Declination)

                Azimuth__1 = objTransform.AzimuthTopocentric

            ElseIf Not IsConnected Then
                Throw New ASCOM.MethodNotImplementedException("Azimuth")
            End If
            TL.LogMessage("Azimuth", "Get - " & Azimuth__1)
            Return Azimuth__1
        End Get
    End Property

    Public ReadOnly Property CanFindHome() As Boolean Implements ITelescopeV3.CanFindHome
        Get
            TL.LogMessage("CanFindHome", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public Function CanMoveAxis(Axis As TelescopeAxes) As Boolean Implements ITelescopeV3.CanMoveAxis
        TL.LogMessage("CanMoveAxis", "Get - " & Axis.ToString())
        Select Case Axis
            Case TelescopeAxes.axisPrimary
                Return True
            Case TelescopeAxes.axisSecondary
                Return True
            Case TelescopeAxes.axisTertiary
                Return False
            Case Else
                'Return False
                Throw New ASCOM.InvalidValueException("CanMoveAxis", Axis.ToString(), "0 to 2")
        End Select
    End Function

    Public ReadOnly Property CanPark() As Boolean Implements ITelescopeV3.CanPark
        Get
            TL.LogMessage("CanPark", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanPulseGuide() As Boolean Implements ITelescopeV3.CanPulseGuide
        Get
            TL.LogMessage("CanPulseGuide", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetDeclinationRate() As Boolean Implements ITelescopeV3.CanSetDeclinationRate
        Get
            TL.LogMessage("CanSetDeclinationRate", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetGuideRates() As Boolean Implements ITelescopeV3.CanSetGuideRates
        Get
            TL.LogMessage("CanSetGuideRates", "Get - " & False.ToString())
            Return False
        End Get
    End Property

    Public ReadOnly Property CanSetPark() As Boolean Implements ITelescopeV3.CanSetPark
        Get
            TL.LogMessage("CanSetPark", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSetPierSide() As Boolean Implements ITelescopeV3.CanSetPierSide
        Get
            TL.LogMessage("CanSetPierSide", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSetRightAscensionRate() As Boolean Implements ITelescopeV3.CanSetRightAscensionRate
        Get
            TL.LogMessage("CanSetRightAscensionRate", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSetTracking() As Boolean Implements ITelescopeV3.CanSetTracking
        Get
            TL.LogMessage("CanSetTracking", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSlew() As Boolean Implements ITelescopeV3.CanSlew
        Get
            TL.LogMessage("CanSlew", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSlewAltAz() As Boolean Implements ITelescopeV3.CanSlewAltAz
        Get
            TL.LogMessage("CanSlewAltAz", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSlewAltAzAsync() As Boolean Implements ITelescopeV3.CanSlewAltAzAsync
        Get
            TL.LogMessage("CanSlewAltAzAsync", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSlewAsync() As Boolean Implements ITelescopeV3.CanSlewAsync
        Get
            TL.LogMessage("CanSlewAsync", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSync() As Boolean Implements ITelescopeV3.CanSync
        Get
            TL.LogMessage("CanSync", "Get - " & True.ToString())
            'Throw New ASCOM.MethodNotImplementedException
            Return True
        End Get
    End Property

    Public ReadOnly Property CanSyncAltAz() As Boolean Implements ITelescopeV3.CanSyncAltAz
        Get
            TL.LogMessage("CanSyncAltAz", "Get - " & CanSyncAltAz.ToString())
            'Throw New ASCOM.MethodNotImplementedException
            Return True
        End Get
    End Property

    Public ReadOnly Property CanUnpark() As Boolean Implements ITelescopeV3.CanUnpark
        Get
            TL.LogMessage("CanUnpark", "Get - " & True.ToString())
            Return True
        End Get
    End Property

    Public ReadOnly Property Declination() As Double Implements ITelescopeV3.Declination
        Get
            Dim declination_1 As Double
            If IsConnected Then
                declination_1 = MotorCounts_to_DEC(GetDECMotorPosition())
            ElseIf Not IsConnected Then
                Throw New ASCOM.NotConnectedException("Declination")
            End If
            TL.LogMessage("Declination", "Get - " & utilities.DegreesToDMS(declination_1))
            Return declination_1
        End Get
    End Property

    Public Property DeclinationRate() As Double Implements ITelescopeV3.DeclinationRate
        Get
            Dim declinationRate__1 As Double = 0.0
            TL.LogMessage("DeclinationRate", "Get - " & DeclinationRate.ToString())
            Return declinationRate__1
        End Get
        Set(value As Double)
            TL.LogMessage("DeclinationRate Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("DeclinationRate", True)
        End Set
    End Property

    Public Function DestinationSideOfPier(RightAscension As Double, Declination As Double) As PierSide Implements ITelescopeV3.DestinationSideOfPier

        Try
            Dim HA As Double
            HA = SiderealTime - RightAscension
            If HA < -12.0# Then
                HA = HA + 24.0#
            ElseIf HA >= 12.0# Then
                HA = HA - 24.0#
            End If

            If HA < 0.0# Then
                Return PierSide.pierWest
            Else
                Return PierSide.pierEast
            End If

            'Return PierSide.pierUnknown
        Catch ex As Exception
            TL.LogMessage("DestinationSideOfPier", "Invalid Operation")
            Throw New ASCOM.InvalidOperationException("DestinationSideOfPier")
        End Try

    End Function

    Public Property DoesRefraction() As Boolean Implements ITelescopeV3.DoesRefraction
        Get
            Try
                DoesRefraction = ApplyRefractionCorrection
            Catch ex As Exception
                TL.LogMessage("DoesRefraction Get", "Error")
                Throw New ASCOM.InvalidOperationException("DoesRefraction")
            End Try
            'TL.LogMessage("DoesRefraction Get", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("DoesRefraction", False)
        End Get
        Set(value As Boolean)
            Try
                ApplyRefractionCorrection = value
                objTransform.Refraction = value
            Catch ex As Exception
                TL.LogMessage("DoesRefraction Set", "Error")
                Throw New ASCOM.InvalidOperationException("DoesRefraction")
            End Try
            'TL.LogMessage("DoesRefraction Set", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("DoesRefraction", True)
        End Set
    End Property

    Public ReadOnly Property EquatorialSystem() As EquatorialCoordinateType Implements ITelescopeV3.EquatorialSystem
        Get
            Dim equatorialSystem__1 As EquatorialCoordinateType = EquatorialCoordinateType.equLocalTopocentric
            TL.LogMessage("DeclinationRate", "Get - " & equatorialSystem__1.ToString())
            Return equatorialSystem__1
        End Get
    End Property

    Public Sub FindHome() Implements ITelescopeV3.FindHome
        Dim RACommand As String
        Dim DECCommand As String
        Dim RAReceived As String
        Dim DECReceived As String

        Try
            If (Not AtPark) And (Not Slewing) Then
                TL.LogMessage("Find Home", "Going Home!")
                Tracking = False
                HomeStatus = True
                RACommand = "ESPt0000000!"
                DECCommand = "ESPt1000000!"
                'send command to slew to HOME position
                SerMutex.WaitOne()
                RAReceived = CommandString(RACommand)
                SerMutex.ReleaseMutex()
                SerMutex.WaitOne()
                DECReceived = CommandString(DECCommand)
                SerMutex.ReleaseMutex()
                While Slewing
                    utilities.WaitForMilliseconds(500)
                    Application.DoEvents()
                End While

                'Tracking = False
                'ParkStatus = True
            Else
                TL.LogMessage("Find Home", "PARKED!")
                Throw New ASCOM.ParkedException("PARKED!")
            End If
        Catch ex As Exception
            TL.LogMessage("FindHome", "Invalid Operation")
            Throw New ASCOM.InvalidOperationException("Find Home")
        End Try
    End Sub

    Public ReadOnly Property FocalLength() As Double Implements ITelescopeV3.FocalLength
        Get
            TL.LogMessage("FocalLength Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("FocalLength", False)
        End Get
    End Property

    Public Property GuideRateDeclination() As Double Implements ITelescopeV3.GuideRateDeclination
        Get
            TL.LogMessage("GuideRateDeclination Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("GuideRateDeclination", False)
        End Get
        Set(value As Double)
            TL.LogMessage("GuideRateDeclination Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("GuideRateDeclination", True)
        End Set
    End Property

    Public Property GuideRateRightAscension() As Double Implements ITelescopeV3.GuideRateRightAscension
        Get
            TL.LogMessage("GuideRateRightAscension Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("GuideRateRightAscension", False)
        End Get
        Set(value As Double)
            TL.LogMessage("GuideRateRightAscension Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("GuideRateRightAscension", True)
        End Set
    End Property

    Public ReadOnly Property IsPulseGuiding() As Boolean Implements ITelescopeV3.IsPulseGuiding
        Get
            TL.LogMessage("IsPulseGuiding Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("IsPulseGuiding", False)
        End Get
    End Property

    Public Sub MoveAxis(Axis As TelescopeAxes, Rate As Double) Implements ITelescopeV3.MoveAxis

        If IsConnected Then
            Try 'to move the axis specified
                Dim Command As String
                Dim myRate As Integer
                Dim recvString As String
                'check to see if parked first and throw error if so...
                If AtPark Then
                    TL.LogMessage("MoveAxis", "At Park!")
                    Throw New ASCOM.ParkedException
                    Return
                End If
                If Axis = TelescopeAxes.axisPrimary Then
                    'Get and set direction as needed
                    If Rate < 0 Then
                        'Check for minimum rate
                        If Rate >= -0.00125# Then
                            Throw New ASCOM.InvalidValueException("MoveAxis")
                            TL.LogMessage("MoveAxis", "Invalid Rate Value")
                        End If
                        'Get current direction
                        Command = "ESGd0!"
                        SerMutex.WaitOne()
                        recvString = CommandString(Command)
                        SerMutex.ReleaseMutex()
                        If Mid(recvString, 6, 1) = "0" Then 'CCW'
                            Command = "ESSd01!"
                            SerMutex.WaitOne()
                            recvString = CommandString(Command)
                            SerMutex.ReleaseMutex()
                        End If
                    ElseIf Rate > 0 Then
                        'Check for minimum rate
                        If Rate <= 0.00125# Then
                            Throw New ASCOM.InvalidValueException("MoveAxis")
                            TL.LogMessage("MoveAxis", "Invalid Rate Value")
                        End If
                        'Get current direction
                        Command = "ESGd0!"
                        SerMutex.WaitOne()
                        recvString = CommandString(Command)
                        SerMutex.ReleaseMutex()
                        If Mid(recvString, 6, 1) = "1" Then 'CW'
                            Command = "ESSd00!"
                            SerMutex.WaitOne()
                            recvString = CommandString(Command)
                            SerMutex.ReleaseMutex()
                        End If
                    End If
                    'Set rate
                    myRate = Convert.ToInt16((Math.Abs(Rate) * MountRACCounts) / 360.0#)
                    Command = "ESSr0" + Format(myRate, "X4") + "!"
                    SerMutex.WaitOne()
                    recvString = CommandString(Command)
                    SerMutex.ReleaseMutex()
                ElseIf Axis = TelescopeAxes.axisSecondary Then
                    'Get and set direction as needed
                    If Rate < 0 Then
                        'Check for minimum rate
                        If Rate >= -0.00125# Then
                            Throw New ASCOM.InvalidValueException("MoveAxis")
                            TL.LogMessage("MoveAxis", "Invalid Rate Value")
                        End If
                        'Get current direction
                        Command = "ESGd1!"
                        SerMutex.WaitOne()
                        recvString = CommandString(Command)
                        SerMutex.ReleaseMutex()
                        If Mid(recvString, 6, 1) = "1" Then 'CW'
                            Command = "ESSd10!"
                            SerMutex.WaitOne()
                            recvString = CommandString(Command)
                            SerMutex.ReleaseMutex()
                        End If
                    ElseIf Rate > 0 Then
                        'Check for minimum rate
                        If Rate <= 0.00125# Then
                            Throw New ASCOM.InvalidValueException("MoveAxis")
                            TL.LogMessage("MoveAxis", "Invalid Rate Value")
                        End If
                        'Get current direction
                        Command = "ESGd1!"
                        SerMutex.WaitOne()
                        recvString = CommandString(Command)
                        SerMutex.ReleaseMutex()
                        If Mid(recvString, 6, 1) = "0" Then 'CCW'
                            Command = "ESSd11!"
                            SerMutex.WaitOne()
                            recvString = CommandString(Command)
                            SerMutex.ReleaseMutex()
                        End If
                    End If
                    'Set rate
                    myRate = Convert.ToInt16((Math.Abs(Rate) * MountDECCounts) / 360.0#)
                    Command = "ESSr1" + Format(myRate, "X4") + "!"
                    SerMutex.WaitOne()
                    recvString = CommandString(Command)
                    SerMutex.ReleaseMutex()
                ElseIf Axis = TelescopeAxes.axisTertiary Then
                    TL.LogMessage("MoveAxis", "Invalid Operation")
                    Throw New ASCOM.InvalidOperationException("MoveAxis")
                End If
            Catch ex As Exception
                TL.LogMessage("MoveAxis", "Invalid Value")
                Throw New ASCOM.InvalidValueException("MoveAxis")
            End Try
        Else
            TL.LogMessage("MoveAxis", "Not Connected")
            Throw New ASCOM.NotConnectedException("Not Connected")
        End If
    End Sub

    Public Sub Park() Implements ITelescopeV3.Park
        Dim RACommand As String
        Dim DECCommand As String
        Dim RAReceived As String
        Dim DECReceived As String

        Try
            If (Not ParkStatus) And (Not Slewing) Then
                TL.LogMessage("Park", "Parking Mount")
                Tracking = False
                ParkStatus = True
                'Set Park to stored values
                RACommand = "ESPt0" & Format(ParkRAPosition, "X6") & "!"
                DECCommand = "ESPt1" & Format(ParkDECPosition, "X6") & "!"
                'RACommand = "ESPt0000000!"
                'DECCommand = "ESPt1000000!"
                'send command to slew to PARK position
                SerMutex.WaitOne()
                RAReceived = CommandString(RACommand)
                SerMutex.ReleaseMutex()
                SerMutex.WaitOne()
                DECReceived = CommandString(DECCommand)
                SerMutex.ReleaseMutex()
                While Slewing
                    utilities.WaitForMilliseconds(500)
                    Application.DoEvents()
                End While
                'Tracking = False
                'ParkStatus = True
                TL.LogMessage("Park", "Mount Parked")
            Else
                TL.LogMessage("Park", "PARKED!")
                Throw New ASCOM.ParkedException("Park")
            End If
        Catch ex As Exception
            TL.LogMessage("Park", "Invalid Operation")
            'Throw New ASCOM.InvalidOperationException("Park")
        End Try
    End Sub

    Public Sub PulseGuide(Direction As GuideDirections, Duration As Integer) Implements ITelescopeV3.PulseGuide
        TL.LogMessage("PulseGuide", "Not implemented")
        Throw New ASCOM.MethodNotImplementedException("PulseGuide")
    End Sub

    Public ReadOnly Property RightAscension() As Double Implements ITelescopeV3.RightAscension
        Get
            Dim rightAscension_1 As Double

            If IsConnected Then
                rightAscension_1 = MotorCounts_to_RA(GetRAMotorPosition())
            ElseIf Not IsConnected Then
                Throw New ASCOM.NotConnectedException("RightAscension")
            End If
            TL.LogMessage("RightAscension", "Get - " & utilities.HoursToHMS(rightAscension_1))
            Return rightAscension_1
        End Get
    End Property

    Public Property RightAscensionRate() As Double Implements ITelescopeV3.RightAscensionRate
        Get
            Dim rightAscensionRate__1 As Double = 0.0
            TL.LogMessage("RightAscensionRate", "Get - " & RightAscensionRate.ToString())
            Return MountRightAscensionRate
        End Get
        Set(value As Double)
            Try
                Dim cmdString As String
                'Dim rcvString As String
                Dim arcSecPerCount As Double
                Dim ratevalue As Double
                Dim intratevalue As Int32

                arcSecPerCount = 1296000.0 / Telescope.MountRACCounts
                ' Set Tracking Rate for desired Rate (uses ESTr0000! command)
                ratevalue = (value / arcSecPerCount) * 25.0
                intratevalue = Convert.ToInt32(Math.Round(ratevalue))
                cmdString = "ESTr" & intratevalue.ToString("X4") & "!"
                SerMutex.WaitOne()
                CommandString(cmdString)
                SerMutex.ReleaseMutex()
                MountRightAscensionRate = value

            Catch ex As Exception
                TL.LogMessage("RightAscensionRate Set", "Invalid Operation")
                Throw New ASCOM.InvalidOperationException("RightAscensionRate")
            End Try
        End Set
    End Property

    Public Sub SetPark() Implements ITelescopeV3.SetPark
        Try
            ParkRAPosition = GetRAMotorPosition()
            ParkDECPosition = GetDECMotorPosition()
            WriteProfile()
            ParkStatus = True

        Catch ex As Exception
            TL.LogMessage("SetPark", "Invalid Operation")
            Throw New ASCOM.InvalidOperationException("SetPark")

        End Try

    End Sub

    Public Property SideOfPier() As PierSide Implements ITelescopeV3.SideOfPier

        Get
            Dim DECMP As Int32
            Dim SOP As PierSide

            DECMP = GetDECMotorPosition()
            If DECMP >= 0 Then
                SOP = PierSide.pierWest
                MountPierSide = PierSide.pierWest
                'SOP = PierSide.pierEast
                'MountPierSide = PierSide.pierEast
            ElseIf DECMP < 0 Then
                SOP = PierSide.pierEast
                MountPierSide = PierSide.pierEast
                'SOP = PierSide.pierWest
                'MountPierSide = PierSide.pierWest
            End If
            Return SOP
        End Get
        Set(value As PierSide)
            Try
                If SideOfPier <> value Then
                    FlipMount = True
                    SlewToCoordinates(RightAscension, Declination)
                End If
                MountPierSide = value

            Catch ex As Exception
                TL.LogMessage("SideOfPier", "Set SideOfPier Fail")
                Throw New ASCOM.InvalidValueException("Set SideOfPier Fail")
            End Try
        End Set
    End Property

    Public ReadOnly Property SiderealTime() As Double Implements ITelescopeV3.SiderealTime
        Get
            Dim my_LMST As Double
            my_LMST = Sidereal_Time()
            TL.LogMessage("SiderealTime", "Get - " & my_LMST.ToString())
            Return my_LMST
        End Get
    End Property

    Public Property SiteElevation() As Double Implements ITelescopeV3.SiteElevation
        Get
            'TL.LogMessage("SiteElevation Get", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("SiteElevation", False)
            TL.LogMessage("SiteElevation", "Get - " & SiteElevationValue.ToString)
            Return Convert.ToDouble(SiteElevationValue)
        End Get
        Set(value As Double)
            'TL.LogMessage("SiteElevation Set", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("SiteElevation", True)
            If (value > 10000.0 Or value < -300) Then
                Throw New ASCOM.InvalidValueException("Invalid Site Elevation Value, -300 to 10000")
                Exit Property
            End If
            SiteElevationValue = Convert.ToString(value)
            WriteProfile()
        End Set
    End Property

    Public Property SiteLatitude() As Double Implements ITelescopeV3.SiteLatitude
        Get
            'TL.LogMessage("SiteLatitude Get", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("SiteLatitude", False)
            TL.LogMessage("SiteLatitude", "Get - " & SiteLatitudeValue.ToString)
            Return Convert.ToDouble(SiteLatitudeValue)
        End Get
        Set(value As Double)
            'TL.LogMessage("SiteLatitude Set", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("SiteLatitude", True)
            If (value > 90.0 Or value < -90.0) Then
                Throw New ASCOM.InvalidValueException("Invalid Site Latitude Value, -90 to +90")
                Exit Property
            End If
            SiteLatitudeValue = Convert.ToString(value)
            WriteProfile()
        End Set
    End Property

    Public Property SiteLongitude() As Double Implements ITelescopeV3.SiteLongitude
        Get
            'TL.LogMessage("SiteLongitude Get", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("SiteLongitude", False)
            TL.LogMessage("SiteLongitude", "Get - " & SiteLongitudeValue.ToString)
            Return Convert.ToDouble(SiteLongitudeValue)
        End Get
        Set(value As Double)
            'TL.LogMessage("SiteLongitude Set", "Not implemented")
            'Throw New ASCOM.PropertyNotImplementedException("SiteLongitude", True)
            If (value > 180.0 Or value < -180.0) Then
                Throw New ASCOM.InvalidValueException("Invalid Site Longitude Value, -180.0 to +180.0")
                Exit Property
            End If
            SiteLongitudeValue = Convert.ToString(value)
            WriteProfile()
        End Set
    End Property

    Public Property SlewSettleTime() As Short Implements ITelescopeV3.SlewSettleTime
        Get
            TL.LogMessage("SlewSettleTime Get", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SlewSettleTime", False)
        End Get
        Set(value As Short)
            TL.LogMessage("SlewSettleTime Set", "Not implemented")
            Throw New ASCOM.PropertyNotImplementedException("SlewSettleTime", True)
        End Set
    End Property

    Public Sub SlewToAltAz(Azimuth As Double, Altitude As Double) Implements ITelescopeV3.SlewToAltAz
        Try
            If Azimuth < 0.0# Or Azimuth >= 360.0 Or Altitude < -90.0# Or Altitude > 90.0# Then
                Throw New ASCOM.InvalidValueException("SlewToAltAz")
            End If
            objTransform.SetAzimuthElevation(Azimuth, Altitude)
            TargetRightAscension = objTransform.RATopocentric
            TargetDeclination = objTransform.DECTopocentric

            SlewToTarget()
        Catch ex As Exception
            TL.LogMessage("SlewToAltAz", "Invalid Operation")
            Throw New ASCOM.InvalidValueException("SlewToAltAz")
        End Try

    End Sub

    Public Sub SlewToAltAzAsync(Azimuth As Double, Altitude As Double) Implements ITelescopeV3.SlewToAltAzAsync
        Try
            If Azimuth < 0.0# Or Azimuth >= 360.0 Or Altitude < -90.0# Or Altitude > 90.0# Then
                Throw New ASCOM.InvalidValueException("SlewToAltAz")
            End If
            objTransform.SetAzimuthElevation(Azimuth, Altitude)
            TargetRightAscension = objTransform.RATopocentric
            TargetDeclination = objTransform.DECTopocentric
            SlewToTargetAsync()
        Catch ex As Exception
            TL.LogMessage("SlewToAltAzAsync", "Invalid Operation")
            Throw New ASCOM.InvalidValueException("SlewToAltAzSync")
        End Try
    End Sub

    Public Sub SlewToCoordinates(RightAscension As Double, Declination As Double) Implements ITelescopeV3.SlewToCoordinates

        Try
            'check to see if parked first and throw error if so...
            If (Not AtPark) And Tracking Then
                TargetRightAscension = RightAscension
                TargetDeclination = Declination
                SlewToTarget()
            Else
                TL.LogMessage("SlewToCoordinates", "@Park OR NOT Tracking!")
                Throw New ASCOM.ParkedException("SlewToCoordinate")
            End If
        Catch ex As Exception
            TL.LogMessage("SlewToCoordinates", "Invalid Operation")
            Throw New ASCOM.InvalidValueException("SlewToCoordinates")
        End Try

    End Sub

    Public Sub SlewToCoordinatesAsync(RightAscension As Double, Declination As Double) Implements ITelescopeV3.SlewToCoordinatesAsync
        Try
            'check to see if parked first and throw error if so...
            If (Not AtPark) And Tracking Then
                TargetRightAscension = RightAscension
                TargetDeclination = Declination
                SlewToTargetAsync()
            Else
                TL.LogMessage("SlewToCoordinatesAsync", "@Park OR NOT Tracking!")
                Throw New ASCOM.ParkedException("SlewToCoordinates")
            End If
        Catch ex As Exception
            TL.LogMessage("SlewToCoordinatesAsync", "Invalid Operation")
            Throw New ASCOM.InvalidValueException("SlewToCoordinateAsync")
        End Try

    End Sub

    Public Sub SlewToTarget() Implements ITelescopeV3.SlewToTarget
        Dim RAReceived As String
        Dim DECReceived As String
        Dim RACounts As Int32
        Dim DECCounts As Int32
        Dim RACounts_Current As Int32
        Dim DECCounts_Current As Int32
        Dim RA_offset As Int32
        Dim RACommand As String
        Dim DECCommand As String
        'Dim SoP As PierSide

        Try
            If IsConnected Then
                'check to see if parked first and throw error if so...
                If Not AtPark Then
                    'Calculate counts from coordinates
                    'SideOfPier = PierSide.pierWest
                    'SoP = SideOfPier

                    RACounts = RA_to_MotorCounts(TargetRightAscension, DestinationSideOfPier(TargetRightAscension, TargetDeclination))
                    DECCounts = DEC_to_MotorCounts(TargetDeclination, DestinationSideOfPier(TargetRightAscension, TargetDeclination))

                    If FlipMount Then
                        'If RACounts < 0 Then
                        'RACounts = RACounts + (MountRACCounts / 2)
                        'ElseIf RACounts >= 0 Then
                        'RACounts = RACounts - (MountRACCounts / 2)
                        'End If
                        'DECCounts = -DECCounts
                        FlipMount = False
                    End If

                    If Tracking Then
                        'Get current counts to calculate slew time for offset adjustment for tracking
                        RACounts_Current = RA_to_MotorCounts(RightAscension, SideOfPier)
                        DECCounts_Current = DEC_to_MotorCounts(Declination, SideOfPier)

                        'Calculate new target value
                        If RACounts > RACounts_Current Then
                            'RA_offset = 53.33 * ((RACounts - RACounts_Current) / 64000) + 250
                            RA_offset = (MountRACCounts / 86400) * (((RACounts - RACounts_Current) / 40000) + 4)
                            RACounts = RACounts + RA_offset
                        ElseIf RACounts_Current > RACounts Then
                            'RA_offset = 53.33 * ((RACounts_Current - RACounts) / 64000) + 250
                            RA_offset = (MountRACCounts / 86400) * (((RACounts_Current - RACounts) / 40000) - 12)
                            RACounts = RACounts - RA_offset
                        End If
                    End If

                    RACommand = "ESPt0" & Mid(RACounts.ToString("X8"), 3, 6) & "!"
                    DECCommand = "ESPt1" & Mid(DECCounts.ToString("X8"), 3, 6) & "!"
                    'send command to slew to position
                    SerMutex.WaitOne()
                    RAReceived = CommandString(RACommand)
                    SerMutex.ReleaseMutex()
                    SerMutex.WaitOne()
                    DECReceived = CommandString(DECCommand)
                    SerMutex.ReleaseMutex()
                    While Slewing
                        utilities.WaitForMilliseconds(500)
                        Application.DoEvents()
                    End While
                    utilities.WaitForMilliseconds(1000)
                    'Reassert tracking
                    If Tracking Then
                        Tracking = True
                    Else
                        Tracking = False
                    End If
                Else
                    TL.LogMessage("SlewTo...", "At Park!")
                    Throw New ASCOM.ParkedException("SlewTo...")
                End If
            End If

        Catch ex As Exception
            TL.LogMessage("SlewToTarget", "Invalid Operation")
            Throw New ASCOM.InvalidValueException("SlewTo...")
        End Try
    End Sub

    Public Sub SlewToTargetAsync() Implements ITelescopeV3.SlewToTargetAsync
        Dim RAReceived As String
        Dim DECReceived As String
        Dim RACounts As Int32
        Dim DECCounts As Int32
        Dim RACounts_Current As Int32
        Dim DECCounts_Current As Int32
        Dim RA_offset As Int32
        Dim RACommand As String
        Dim DECCommand As String

        Try 'to slew the mount as commanded
            If IsConnected Then
                'check to see if parked first and throw error if so...
                If Not AtPark Then
                    'Calculate counts from coordinates
                    'SideOfPier = PierSide.pierWest
                    RACounts = RA_to_MotorCounts(TargetRightAscension, DestinationSideOfPier(TargetRightAscension, TargetDeclination))
                    DECCounts = DEC_to_MotorCounts(TargetDeclination, DestinationSideOfPier(TargetRightAscension, TargetDeclination))

                    If Tracking Then
                        'Get current counts to calculate slew time for offset adjustment for tracking
                        RACounts_Current = RA_to_MotorCounts(RightAscension, SideOfPier)
                        DECCounts_Current = DEC_to_MotorCounts(Declination, SideOfPier)

                        'Calculate new target value
                        If RACounts > RACounts_Current Then
                            'RA_offset = 53.33 * ((RACounts - RACounts_Current) / 64000) + 250
                            RA_offset = (MountRACCounts / 86400) * (((RACounts - RACounts_Current) / 40000) + 4)
                            RACounts = RACounts + RA_offset
                        ElseIf RACounts_Current > RACounts Then
                            'RA_offset = 53.33 * ((RACounts_Current - RACounts) / 64000) + 250
                            RA_offset = (MountRACCounts / 86400) * (((RACounts - RACounts_Current) / 40000) - 12)
                            RACounts = RACounts - RA_offset
                        End If
                    End If

                    RACommand = "ESPt0" & Mid(RACounts.ToString("X8"), 3, 6) & "!"
                    DECCommand = "ESPt1" & Mid(DECCounts.ToString("X8"), 3, 6) & "!"
                    'send command to slew to position
                    SerMutex.WaitOne()
                    RAReceived = CommandString(RACommand)
                    SerMutex.ReleaseMutex()
                    SerMutex.WaitOne()
                    DECReceived = CommandString(DECCommand)
                    SerMutex.ReleaseMutex()
                Else
                    TL.LogMessage("SlewTo...Async", "At Park!")
                    Throw New ASCOM.ParkedException("SlewTo...Async")
                End If
            End If

        Catch ex As Exception
            TL.LogMessage("SlewToTargetAsync", "Invalid Operation")
            Throw New ASCOM.InvalidValueException("SlewTo...Async")
        End Try
    End Sub

    Public ReadOnly Property Slewing() As Boolean Implements ITelescopeV3.Slewing
        Get
            Dim tempRARATE As String
            Dim tempDECRate As String
            Dim tempCOMMAND As String
            Dim RARate As Int32
            Dim DECRate As Int32

            Try
                tempCOMMAND = "ESGr0!"
                SerMutex.WaitOne()
                tempRARATE = CommandString(tempCOMMAND)
                SerMutex.ReleaseMutex()
                If Not EvalCommand(tempCOMMAND, tempRARATE) Then
                    Exit Property
                End If

                tempCOMMAND = "ESGr1!"
                SerMutex.WaitOne()
                tempDECRate = CommandString(tempCOMMAND)
                SerMutex.ReleaseMutex()
                If Not EvalCommand(tempCOMMAND, tempDECRate) Then
                    Exit Property
                End If

                RARate = Convert.ToInt32("0000" + Mid(tempRARATE, 6, 4), 16)
                DECRate = Convert.ToInt32("0000" + Mid(tempDECRate, 6, 4), 16)
                If (RARate >= minSlewRate) Or (DECRate >= minSlewRate) Then 'the value of 100 is a rate less than any slew rate
                    Slewing = True
                Else
                    Slewing = False
                    If RATracking = True Then
                        're-assert tracking
                        Tracking = True
                        'Return
                    End If

                End If

            Catch ex As Exception
                TL.LogMessage("Slewing Get", "Invalid Operation")
                'Throw New ASCOM.InvalidValueException

            End Try
        End Get
    End Property

    Public Sub SyncToAltAz(Azimuth As Double, Altitude As Double) Implements ITelescopeV3.SyncToAltAz
        Dim tempRA As Double
        Dim tempDEC As Double
        Try
            If Azimuth < 0.0# Or Azimuth >= 360.0 Or Altitude < -90.0# Or Altitude > 90.0# Then
                Throw New ASCOM.InvalidValueException("SyncToAltAz")
            End If
            objTransform.SetAzimuthElevation(Azimuth, Altitude)
            tempRA = objTransform.RATopocentric
            tempDEC = objTransform.DECTopocentric
            TargetRightAscension = tempRA
            TargetDeclination = tempDEC
            SyncToTarget()
        Catch ex As Exception
            TL.LogMessage("SyncToAltAz", "Invalid Value")
            Throw New ASCOM.InvalidValueException("SyncToAltAz")
        End Try
    End Sub

    Public Sub SyncToCoordinates(RightAscension As Double, Declination As Double) Implements ITelescopeV3.SyncToCoordinates
        Try
            'check to see if parked first and throw error if so...
            If Not AtPark Then
                TargetRightAscension = RightAscension
                TargetDeclination = Declination
                SyncToTarget()
            Else
                TL.LogMessage("SyncToCoordinates", "At Park!")
                Throw New ASCOM.ParkedException("SyncToCoordinates")
            End If
        Catch ex As Exception
            TL.LogMessage("SyncToCoordinatesAsync", "Invalid Value")
            Throw New ASCOM.InvalidValueException("SyncToCoordinates")
        End Try
    End Sub

    Public Sub SyncToTarget() Implements ITelescopeV3.SyncToTarget
        Dim RACounts As Int32
        Dim DECCounts As Int32
        Dim RACommand As String
        Dim DECCommand As String
        Dim RAReceived As String
        Dim DECReceived As String

        Try
            If IsConnected Then
                'check to see if parked first and throw error if so...ALSO IF NOT TRACKING!!!!
                If Not AtPark Then
                    RACounts = RA_to_MotorCounts(TargetRightAscension, DestinationSideOfPier(TargetRightAscension, TargetDeclination))
                    DECCounts = DEC_to_MotorCounts(TargetDeclination, DestinationSideOfPier(TargetRightAscension, TargetDeclination))
                    RACommand = "ESSp0" & Mid(RACounts.ToString("X8"), 3, 6) & "!"
                    DECCommand = "ESSp1" & Mid(DECCounts.ToString("X8"), 3, 6) & "!"
                    'send commands to set new motor position
                    SerMutex.WaitOne()
                    RAReceived = CommandString(RACommand)
                    SerMutex.ReleaseMutex()
                    SerMutex.WaitOne()
                    DECReceived = CommandString(DECCommand)
                    SerMutex.ReleaseMutex()
                    If Tracking Then
                        'reassert tracking
                        Tracking() = True
                    Else
                        TL.LogMessage("SyncToTarget", "Not Tracking!")
                        Throw New ASCOM.InvalidOperationException("SyncToTarget")
                    End If
                Else
                    TL.LogMessage("SyncToTarget", "At Park!")
                    Throw New ASCOM.ParkedException("SyncToTarget")
                End If
            End If
        Catch ex As Exception
            TL.LogMessage("SyncToTarget", "Invalid Operation")
            Throw New ASCOM.InvalidOperationException("SyncToTarget")
        End Try
    End Sub

    Public Property TargetDeclination() As Double Implements ITelescopeV3.TargetDeclination
        Get
            Try
                If Not DECTargetSet Then
                    Throw New ASCOM.ValueNotSetException
                End If
                If Not (DECTarget < -90.0 Or DECTarget > 90.0) Then
                    Return DECTarget
                Else
                    Throw New ASCOM.InvalidValueException
                End If
            Catch ex As Exception
                Throw New ASCOM.ValueNotSetException
            End Try
        End Get
        Set(value As Double)
            Try
                If Not (value > 90.0 Or value < -90.0) Then
                    DECTarget = value
                    DECTargetSet = True
                Else
                    Throw New ASCOM.InvalidValueException
                    'Exit Property
                End If
            Catch ex As Exception
                Throw New ASCOM.InvalidValueException
            End Try
        End Set
    End Property

    Public Property TargetRightAscension() As Double Implements ITelescopeV3.TargetRightAscension

        Get
            Try
                If Not RATargetSet Then
                    Throw New ASCOM.ValueNotSetException
                End If
                If Not (RATarget < 0.0 Or RATarget >= 24.0) Then
                    Return RATarget
                Else
                    Throw New ASCOM.InvalidValueException
                End If
            Catch ex As Exception
                Throw New ASCOM.ValueNotSetException
            End Try
        End Get
        Set(value As Double)
            Try
                If Not (value >= 24.0 Or value < 0.0) Then
                    RATarget = value
                    RATargetSet = True
                Else
                    Throw New ASCOM.InvalidValueException
                    'Exit Property
                End If
            Catch ex As Exception
                Throw New ASCOM.InvalidValueException
            End Try
        End Set
    End Property

    Public Property Tracking() As Boolean Implements ITelescopeV3.Tracking

        Get
            'Dim tracking__1 As Boolean = True
            TL.LogMessage("Tracking", "Get - " & RATracking.ToString())
            Tracking = RATracking
        End Get
        Set(value As Boolean)
            Dim TrackCommand As String
            'Dim TrackRate As Int32

            Try
                If value = True Then
                    'check to see if already tracking
                    'Set correct direction for hemisphere, 1 for Northern, 0 for Southern
                    SerMutex.WaitOne()
                    CommandString("ESSd01!")
                    SerMutex.ReleaseMutex()

                    'Set Sidereal Tracking Rate (tracking on PMC-Eight is 25 x slower than normal rate)
                    TrackingRate = DriveRates.driveSidereal
                    MountRightAscensionRate = 15.0225
                    MountTrackingRate = DriveRates.driveSidereal
                    'TrackRate = Int((15.0225 / (1296000.0 / Telescope.MountRACCounts)) * 25.0)
                    'TrackCommand = "ESTr" & TrackRate.ToString("X4") & "!"
                    'SerMutex.WaitOne()
                    'CommandString(TrackCommand)
                    'SerMutex.ReleaseMutex()

                    RATracking = True
                ElseIf value = False Then
                    'Set tracking Rate to 0 (zero)
                    TrackCommand = "ESTr0000!"
                    SerMutex.WaitOne()
                    CommandString(TrackCommand)
                    SerMutex.ReleaseMutex()
                    RATracking = False
                End If
            Catch ex As Exception
                TL.LogMessage("Tracking Set", "Invalid Operation")
                Throw New ASCOM.InvalidOperationException("Tracking")
            End Try
        End Set
    End Property

    Public Property TrackingRate() As DriveRates Implements ITelescopeV3.TrackingRate
        Get
            TL.LogMessage("TrackingRate Get", "")
            'Throw New ASCOM.PropertyNotImplementedException("TrackingRate", False)
            'For Each myTrackingRate In TrackingRates
            Return MountTrackingRate
            'Next
        End Get
        Set(value As DriveRates)
            Dim cmdString As String
            'Dim rcvString As String
            Dim arcSecPerCount As Double
            Dim ratevalue As Double
            Dim intratevalue As Int32

            Try
                arcSecPerCount = 1296000.0 / Telescope.MountRACCounts
                Select Case value
                    Case DriveRates.driveSidereal 'Sidereal Rate
                        ' Set Tracking Rate for Sidereal Rate (uses ESTr0000! command)
                        ratevalue = (15.0225 / arcSecPerCount) * 25.0
                        intratevalue = Convert.ToInt32(Math.Round(ratevalue))
                        cmdString = "ESTr" & intratevalue.ToString("X4") & "!"
                        SerMutex.WaitOne()
                        CommandString(cmdString)
                        SerMutex.ReleaseMutex()
                        MountTrackingRate = DriveRates.driveSidereal
                        MountRightAscensionRate = 15.0225
                    Case DriveRates.driveLunar
                        ' Set Tracking Rate for Sidereal Rate (uses ESTr0000! command)
                        ratevalue = (15.0225 / arcSecPerCount) * 25.0
                        intratevalue = Convert.ToInt32(Math.Round(ratevalue))
                        cmdString = "ESTr" & intratevalue.ToString("X4") & "!"
                        SerMutex.WaitOne()
                        CommandString(cmdString)
                        SerMutex.ReleaseMutex()
                        MountTrackingRate = DriveRates.driveLunar
                        MountRightAscensionRate = 15.0225
                    Case DriveRates.driveSolar
                        ' Set Tracking Rate for Sidereal Rate (uses ESTr0000! command)
                        ratevalue = (15.0225 / arcSecPerCount) * 25.0
                        intratevalue = Convert.ToInt32(Math.Round(ratevalue))
                        cmdString = "ESTr" & intratevalue.ToString("X4") & "!"
                        SerMutex.WaitOne()
                        CommandString(cmdString)
                        SerMutex.ReleaseMutex()
                        MountTrackingRate = DriveRates.driveSolar
                        MountRightAscensionRate = 15.0225
                    Case DriveRates.driveKing
                        ' Set Tracking Rate for Sidereal Rate (uses ESTr0000! command)
                        ratevalue = (15.0225 / arcSecPerCount) * 25.0
                        intratevalue = Convert.ToInt32(Math.Round(ratevalue))
                        cmdString = "ESTr" & intratevalue.ToString("X4") & "!"
                        SerMutex.WaitOne()
                        CommandString(cmdString)
                        SerMutex.ReleaseMutex()
                        MountTrackingRate = DriveRates.driveKing
                        MountRightAscensionRate = 15.0225
                End Select
            Catch ex As Exception
                TL.LogMessage("TrackingRate Set", "Invalid Operation")
                Throw New ASCOM.InvalidOperationException("TrackingRate")

            End Try
        End Set
    End Property

    Public ReadOnly Property TrackingRates() As ITrackingRates Implements ITelescopeV3.TrackingRates
        Get
            Dim trackingRates__1 As ITrackingRates = New TrackingRates()
            'Dim trackingRates__1 As New TrackingRates
            TL.LogMessage("TrackingRates", "Get - ")
            For Each driveRate As DriveRates In trackingRates__1
                TL.LogMessage("TrackingRates", "Get - " & driveRate.ToString())
            Next
            Return trackingRates__1
            'Return New TrackingRates()
        End Get
    End Property

    Public Property UTCDate() As DateTime Implements ITelescopeV3.UTCDate
        Get
            Dim utcDate__1 As DateTime = DateTime.UtcNow
            TL.LogMessage("UTC DateTime", "Get - " & [String].Format("MM/dd/yy HH:mm:ss", utcDate__1))
            Return utcDate__1
        End Get
        Set(value As DateTime)
            Throw New ASCOM.PropertyNotImplementedException("UTCDate", True)
        End Set
    End Property

    Public Sub Unpark() Implements ITelescopeV3.Unpark
        Try
            ParkStatus = False
            'TestMotorCalcs()
        Catch ex As Exception
            TL.LogMessage("Unpark", "Unparked")
            'Throw New ASCOM.MethodNotImplementedException("Unpark")

        End Try
    End Sub

#End Region

#Region "Private properties and methods"
    ' here are some useful properties and methods that can be used as required
    ' to help with

#Region "ASCOM Registration"

    Private Shared Sub RegUnregASCOM(ByVal bRegister As Boolean)

        Using P As New Profile() With {.DeviceType = "Telescope"}
            If bRegister Then
                P.Register(driverID, driverDescription)
            Else
                P.Unregister(driverID)
            End If
        End Using

    End Sub

    <ComRegisterFunction()> _
    Public Shared Sub RegisterASCOM(ByVal T As Type)

        RegUnregASCOM(True)

    End Sub

    <ComUnregisterFunction()> _
    Public Shared Sub UnregisterASCOM(ByVal T As Type)

        RegUnregASCOM(False)

    End Sub

#End Region

    ''' <summary>
    ''' Returns true if there is a valid connection to the driver hardware
    ''' </summary>
    Private ReadOnly Property IsConnected As Boolean
        Get
            ' TODO check that the driver hardware connection exists and is connected to the hardware
            Return connectedState
        End Get
    End Property

    ''' <summary>
    ''' Use this function to throw an exception if we aren't connected to the hardware
    ''' </summary>
    ''' <param name="message"></param>
    Private Sub CheckConnected(ByVal message As String)
        If Not IsConnected Then
            Throw New NotConnectedException(message)
        End If
    End Sub

    ''' <summary>
    ''' Read the device configuration from the ASCOM Profile store
    ''' </summary>
    Friend Sub ReadProfile()
        Using driverProfile As New Profile()
            driverProfile.DeviceType = "Telescope"
            traceState = Convert.ToBoolean(driverProfile.GetValue(driverID, traceStateProfileName, String.Empty, traceStateDefault))
            comPort = driverProfile.GetValue(driverID, comPortProfileName, String.Empty, comPortDefault)
            comSpeed = driverProfile.GetValue(driverID, comSpeedProfileName, String.Empty, comSpeedDefault)
            IPAddress = driverProfile.GetValue(driverID, IPAddressProfileName, String.Empty, IPAddressDefault)
            IPPort = driverProfile.GetValue(driverID, IPPortProfileName, String.Empty, IPPortDefault)
            WirelessEnabled = Convert.ToBoolean(driverProfile.GetValue(driverID, WirelessEnabledProfileName, String.Empty, WirelessEnabledDefault))
            WirelessProtocol = driverProfile.GetValue(driverID, WirelessProtocolProfileName, String.Empty, WirelessProtocolDefault)
            Mount = driverProfile.GetValue(driverID, MountProfileName, String.Empty, MountDefault)
            MountRACCounts = Convert.ToInt32(driverProfile.GetValue(driverID, MountRACountsProfileName, String.Empty, MountRACountsDefault))
            MountDECCounts = Convert.ToInt32(driverProfile.GetValue(driverID, MountDECCountsProfileName, String.Empty, MountDECCountsDefault))
            SiteLocation = driverProfile.GetValue(driverID, SiteLocationProfileName, String.Empty, SiteLocationDefault)
            SiteElevationValue = Convert.ToDouble(driverProfile.GetValue(driverID, SiteElevationProfileName, String.Empty, SiteElevationDefault))
            SiteLatitudeValue = Convert.ToDouble(driverProfile.GetValue(driverID, SiteLatitudeProfileName, String.Empty, SiteLatitudeDefault))
            SiteLongitudeValue = Convert.ToDouble(driverProfile.GetValue(driverID, SiteLongitudeProfileName, String.Empty, SiteLongitudeDefault))
            SiteAmbientTemperatureValue = Convert.ToDouble(driverProfile.GetValue(driverID, SiteAmbientTemperatureProfileName, String.Empty, SiteAmbientTemperatureDefault))
            ApplyRefractionCorrection = Convert.ToBoolean(driverProfile.GetValue(driverID, ApplyRefractionCorrectionProfileName, String.Empty, ApplyRefractionCorrectionDefault))
            ParkRAPosition = Convert.ToInt32(driverProfile.GetValue(driverID, ParkRAPositionProfileName, String.Empty, ParkRAPositionDefault))
            ParkDECPosition = Convert.ToInt32(driverProfile.GetValue(driverID, ParkDECPositionProfileName, String.Empty, ParkDECPositionDefault))
        End Using
    End Sub

    ''' <summary>
    ''' Write the device configuration to the  ASCOM  Profile store
    ''' </summary>
    Friend Sub WriteProfile()
        Using driverProfile As New Profile()
            driverProfile.DeviceType = "Telescope"
            driverProfile.WriteValue(driverID, traceStateProfileName, traceState.ToString())
            driverProfile.WriteValue(driverID, comPortProfileName, comPort.ToString())
            driverProfile.WriteValue(driverID, comSpeedProfileName, comSpeed.ToString())
            driverProfile.WriteValue(driverID, IPAddressProfileName, IPAddress.ToString())
            driverProfile.WriteValue(driverID, IPPortProfileName, IPPort.ToString())
            driverProfile.WriteValue(driverID, WirelessEnabledProfileName, WirelessEnabled.ToString())
            driverProfile.WriteValue(driverID, WirelessProtocolProfileName, WirelessProtocol.ToString())
            driverProfile.WriteValue(driverID, MountProfileName, Mount.ToString())
            driverProfile.WriteValue(driverID, MountRACountsProfileName, MountRACCounts.ToString())
            driverProfile.WriteValue(driverID, MountDECCountsProfileName, MountDECCounts.ToString())
            driverProfile.WriteValue(driverID, SiteLocationProfileName, SiteLocation.ToString())
            driverProfile.WriteValue(driverID, SiteElevationProfileName, SiteElevationValue.ToString())
            driverProfile.WriteValue(driverID, SiteLatitudeProfileName, SiteLatitudeValue.ToString())
            driverProfile.WriteValue(driverID, SiteLongitudeProfileName, SiteLongitudeValue.ToString())
            driverProfile.WriteValue(driverID, SiteAmbientTemperatureProfileName, SiteAmbientTemperatureValue.ToString())
            driverProfile.WriteValue(driverID, ApplyRefractionCorrectionProfileName, ApplyRefractionCorrection.ToString())
            driverProfile.WriteValue(driverID, ParkRAPositionProfileName, ParkRAPosition.ToString())
            driverProfile.WriteValue(driverID, ParkDECPositionProfileName, ParkDECPosition.ToString())
        End Using
    End Sub

    Private Function EvalCommand(CmdString As String, ReturnString As String) As Boolean
        Dim sent As String
        Dim received As String
        'only check that the sub command type and axis is the same.
        sent = Mid(CmdString, 1, 5)
        received = Mid(ReturnString, 1, 5)
        If sent <> received Then
            EvalCommand = False
            'objSerial.ClearBuffers()
        Else
            EvalCommand = True
        End If
    End Function

    Private Function Sidereal_Time() As Double
        j2000 = "1/1/2000 12:00:00"
        deltaTime = DateTime.UtcNow() - j2000
        LMSTtot = 0.77905727325# + (1.00273790935079# * deltaTime.TotalDays)
        di = Math.Floor(LMSTtot)
        LMST = ((LMSTtot - di) * 360.0#) + Telescope.SiteLongitudeValue
        If (LMST < 0) Then
            LMST = LMST + 360.0#
        ElseIf (LMST > 360.0#) Then
            LMST = LMST - 360.0#
        End If
        LMST = 24.0# * (LMST / 360.0#)
        Return LMST
    End Function

    Private Function GetRAMotorPosition() As Int32
        Dim RAReceived As String
        Dim HourAnglePlusSix As Int32
        Dim cmdString As String
        Dim recString As String

        If IsConnected Then
            cmdString = "ESGp0!"
            SerMutex.WaitOne()
            recString = CommandString(cmdString)
            SerMutex.ReleaseMutex()
            RAReceived = "00" + Mid(recString, 6, 6)
            If EvalCommand(cmdString, recString) Then
                HourAnglePlusSix = Convert.ToInt32(RAReceived, 16)
                If HourAnglePlusSix >= 8388608 Then ' calculate negative value
                    HourAnglePlusSix = 0 - (16777216 - HourAnglePlusSix)
                End If
                PrevRAMotor = HourAnglePlusSix
            Else
                HourAnglePlusSix = PrevRAMotor
            End If
        ElseIf Not IsConnected Then
            Throw New ASCOM.NotConnectedException("GetRAMotorPosition")
        End If
        Return HourAnglePlusSix

    End Function

    Private Function GetDECMotorPosition() As Int32
        Dim DECReceived As String
        Dim Degrees As Int32
        Dim cmdString As String
        Dim RecString As String

        If IsConnected Then
            cmdString = "ESGp1!"
            SerMutex.WaitOne()
            RecString = CommandString(cmdString)
            SerMutex.ReleaseMutex()
            DECReceived = "00" + Mid(RecString, 6, 6)
            If EvalCommand(cmdString, RecString) Then
                Degrees = Convert.ToInt32(DECReceived, 16)
                If Degrees >= 8388608 Then ' calculate negative value
                    Degrees = 0 - (16777216 - Degrees)
                End If
                PrevDECMotor = Degrees
            Else
                Degrees = PrevDECMotor
            End If
        ElseIf Not IsConnected Then
            Throw New ASCOM.NotConnectedException("GetDECMotorPosition")
        End If
        Return Degrees

    End Function

    Public Function RA_to_MotorCounts(RA_value As Double, SOP As PierSide) As Int32
        Dim MotorAngle As Double
        Dim MotorCounts As Int32
        Dim HourAngle As Double

        HourAngle = SiderealTime - RA_value

        'limit values to +/- 12 hours
        If HourAngle > 12 Then
            HourAngle = HourAngle - 24
        ElseIf HourAngle <= -12 Then
            HourAngle = HourAngle + 24
        End If

        If SOP = PierSide.pierEast Then
            MotorAngle = HourAngle - 6
        ElseIf SOP = PierSide.pierWest Then
            MotorAngle = HourAngle + 6
        End If

        MotorCounts = MotorAngle * MountRACCounts / 24

        Return MotorCounts
    End Function

    Public Function MotorCounts_to_RA(MC_value As Int32) As Double
        Dim MotorAngle As Double
        Dim RA_value As Double
        Dim HourAngle As Double
        Dim DECCounts As Int32

        DECCounts = GetDECMotorPosition()

        MotorAngle = (24.0# * MC_value) / Telescope.MountRACCounts

        If DECCounts < 0 Then
            HourAngle = MotorAngle + 6
        ElseIf DECCounts >= 0 Then
            HourAngle = MotorAngle - 6
        End If

        RA_value = SiderealTime - HourAngle

        If RA_value >= 24.0# Then
            RA_value = RA_value - 24.0#
        ElseIf RA_value < 0.0# Then
            RA_value = RA_value + 24.0#
        End If

        Return RA_value
    End Function

    Public Function DEC_to_MotorCounts(DEC_value As Double, SOP As PierSide) As Int32
        Dim MotorAngle As Double
        Dim MotorCounts As Int32

        If SOP = PierSide.pierEast Then
            MotorAngle = (DEC_value - 90.0#)
        ElseIf SOP = PierSide.pierWest Then
            MotorAngle = -(DEC_value - 90.0#)
        End If

        MotorCounts = (MotorAngle / 360.0) * Telescope.MountDECCounts

        Return MotorCounts
    End Function

    Public Function MotorCounts_to_DEC(MC_value As Int32) As Double
        Dim MotorAngle As Double
        Dim DEC_value As Double

        MotorAngle = (360.0# * MC_value) / Telescope.MountDECCounts

        If MotorAngle >= 0 Then
            DEC_value = 90 - MotorAngle
        ElseIf MotorAngle < 0 Then
            DEC_value = 90 + MotorAngle
        End If

        Return DEC_value
    End Function

    Private Sub TestMotorCalcs()
        'put test code here to test private routines
        Dim RA_test As Double
        Dim DEC_test As Double
        Dim RAMC_test As Int32
        Dim DECMC_test As Int32
        Dim RA_test_res As Double
        Dim DEC_test_res As Double
        'Dim RAMC_test_res As Int32
        Dim DECMC_test_res As Int32
        Dim ST As Double
        Dim ALT_test As Double
        Dim AZ_test As Double

        'normal PierEast
        ALT_test = 45.0
        AZ_test = 225.0

        objTransform.SetAzimuthElevation(AZ_test, ALT_test)
        RA_test = objTransform.RATopocentric
        DEC_test = objTransform.DECTopocentric

        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierEast)
        'RAMC_test = RA_to_MotorCounts(RA_test)
        'RAMC_test = RA_to_MotorCounts(RA_test)
        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierWest)
        'RAMC_test = RA_to_MotorCounts(RA_test)
        'RAMC_test = RA_to_MotorCounts(RA_test)

        'normal PierWest
        ALT_test = 45.0
        AZ_test = 135.0

        objTransform.SetAzimuthElevation(AZ_test, ALT_test)
        RA_test = objTransform.RATopocentric
        DEC_test = objTransform.DECTopocentric

        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierWest)
        'RAMC_test = RA_to_MotorCounts(RA_test)
        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierEast)
        'RAMC_test = RA_to_MotorCounts(RA_test)

        ST = SiderealTime
        ' -------------------------------------------------------------------------------------
        'RA_test = ST + 3
        If RA_test >= 24.0 Then
            RA_test = RA_test - 24.0
        ElseIf RA_test < 0.0 Then
            RA_test = RA_test + 24.0
        End If

        'DEC_test = 45.0
        'RAMC_test = RA_to_MotorCounts(RA_test)
        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierEast)
        RA_test_res = MotorCounts_to_RA(RAMC_test)
        DEC_test_res = MotorCounts_to_DEC(DECMC_test)
        'RAMC_test_res = RA_to_MotorCounts(RA_test_res)
        DECMC_test_res = DEC_to_MotorCounts(DEC_test_res, PierSide.pierEast)

        RA_test = ST + 3
        If RA_test >= 24.0 Then
            RA_test = RA_test - 24.0
        ElseIf RA_test < 0.0 Then
            RA_test = RA_test + 24.0
        End If
        DEC_test = 45.0
        'RAMC_test = RA_to_MotorCounts(RA_test)
        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierWest)
        RA_test_res = MotorCounts_to_RA(RAMC_test)
        DEC_test_res = MotorCounts_to_DEC(DECMC_test)
        'RAMC_test_res = RA_to_MotorCounts(RA_test_res)
        DECMC_test_res = DEC_to_MotorCounts(DEC_test_res, PierSide.pierWest)

        ' -------------------------------------------------------------------------------------
        RA_test = ST + 9
        If RA_test >= 24.0 Then
            RA_test = RA_test - 24.0
        ElseIf RA_test < 0.0 Then
            RA_test = RA_test + 24.0
        End If
        DEC_test = 45.0
        'RAMC_test = RA_to_MotorCounts(RA_test)
        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierEast)
        RA_test_res = MotorCounts_to_RA(RAMC_test)
        DEC_test_res = MotorCounts_to_DEC(DECMC_test)
        'RAMC_test_res = RA_to_MotorCounts(RA_test_res)
        DECMC_test_res = DEC_to_MotorCounts(DEC_test_res, PierSide.pierEast)

        RA_test = ST + 9
        If RA_test >= 24.0 Then
            RA_test = RA_test - 24.0
        ElseIf RA_test < 0.0 Then
            RA_test = RA_test + 24.0
        End If
        DEC_test = 45.0
        'RAMC_test = RA_to_MotorCounts(RA_test)
        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierWest)
        RA_test_res = MotorCounts_to_RA(RAMC_test)
        DEC_test_res = MotorCounts_to_DEC(DECMC_test)
        'RAMC_test_res = RA_to_MotorCounts(RA_test_res)
        DECMC_test_res = DEC_to_MotorCounts(DEC_test_res, PierSide.pierWest)

        ' -------------------------------------------------------------------------------------
        RA_test = ST - 3
        If RA_test >= 24.0 Then
            RA_test = RA_test - 24.0
        ElseIf RA_test < 0.0 Then
            RA_test = RA_test + 24.0
        End If
        DEC_test = 45.0
        'RAMC_test = RA_to_MotorCounts(RA_test)
        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierEast)
        RA_test_res = MotorCounts_to_RA(RAMC_test)
        DEC_test_res = MotorCounts_to_DEC(DECMC_test)
        'RAMC_test_res = RA_to_MotorCounts(RA_test_res)
        DECMC_test_res = DEC_to_MotorCounts(DEC_test_res, PierSide.pierEast)

        RA_test = ST - 3
        If RA_test >= 24.0 Then
            RA_test = RA_test - 24.0
        ElseIf RA_test < 0.0 Then
            RA_test = RA_test + 24.0
        End If
        DEC_test = 45.0
        'RAMC_test = RA_to_MotorCounts(RA_test)
        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierWest)
        RA_test_res = MotorCounts_to_RA(RAMC_test)
        DEC_test_res = MotorCounts_to_DEC(DECMC_test)
        'RAMC_test_res = RA_to_MotorCounts(RA_test_res)
        DECMC_test_res = DEC_to_MotorCounts(DEC_test_res, PierSide.pierWest)

        ' -------------------------------------------------------------------------------------
        RA_test = ST - 9
        If RA_test >= 24.0 Then
            RA_test = RA_test - 24.0
        ElseIf RA_test < 0.0 Then
            RA_test = RA_test + 24.0
        End If
        DEC_test = 45.0
        'RAMC_test = RA_to_MotorCounts(RA_test)
        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierEast)
        RA_test_res = MotorCounts_to_RA(RAMC_test)
        DEC_test_res = MotorCounts_to_DEC(DECMC_test)
        'RAMC_test_res = RA_to_MotorCounts(RA_test_res)
        DECMC_test_res = DEC_to_MotorCounts(DEC_test_res, PierSide.pierEast)

        RA_test = ST - 9
        If RA_test >= 24.0 Then
            RA_test = RA_test - 24.0
        ElseIf RA_test < 0.0 Then
            RA_test = RA_test + 24.0
        End If
        DEC_test = 45.0
        'RAMC_test = RA_to_MotorCounts(RA_test)
        DECMC_test = DEC_to_MotorCounts(DEC_test, PierSide.pierWest)
        RA_test_res = MotorCounts_to_RA(RAMC_test)
        DEC_test_res = MotorCounts_to_DEC(DECMC_test)
        'RAMC_test_res = RA_to_MotorCounts(RA_test_res)
        DECMC_test_res = DEC_to_MotorCounts(DEC_test_res, PierSide.pierWest)

        ST = SiderealTime


    End Sub
    Private Function Test(value As Double) As Double

        Dim CelestialCoordinates = Matrix(Of Double).Build.Dense(3, 3)
        Dim TelescopeCoordinates = Matrix(Of Double).Build.Dense(3, 3)

        Dim CL_star1 As Double
        Dim CL_star2 As Double
        Dim CL_star3 As Double
        Dim CM_star1 As Double
        Dim CM_star2 As Double
        Dim CM_star3 As Double
        Dim CN_star1 As Double
        Dim CN_star2 As Double
        Dim CN_star3 As Double

        Dim TL_star1 As Double
        Dim TL_star2 As Double
        Dim TL_star3 As Double
        Dim TM_star1 As Double
        Dim TM_star2 As Double
        Dim TM_star3 As Double
        Dim TN_star1 As Double
        Dim TN_star2 As Double
        Dim TN_star3 As Double

        CelestialCoordinates(1, 1) = CL_star1
        CelestialCoordinates(1, 2) = CL_star2
        CelestialCoordinates(1, 3) = CL_star3
        CelestialCoordinates(2, 1) = CM_star1
        CelestialCoordinates(2, 2) = CM_star2
        CelestialCoordinates(2, 2) = CM_star3
        CelestialCoordinates(3, 1) = CN_star1
        CelestialCoordinates(3, 2) = CN_star2
        CelestialCoordinates(3, 3) = CN_star3

        TelescopeCoordinates(1, 1) = TL_star1
        TelescopeCoordinates(1, 2) = TL_star2
        TelescopeCoordinates(1, 3) = TL_star3
        TelescopeCoordinates(2, 1) = TM_star1
        TelescopeCoordinates(2, 2) = TM_star2
        TelescopeCoordinates(2, 3) = TM_star3
        TelescopeCoordinates(3, 1) = TN_star1
        TelescopeCoordinates(3, 2) = TN_star2
        TelescopeCoordinates(3, 3) = TN_star3



    End Function

    Private Function ALT_to_MotorCounts() As Long

    End Function

    Private Function AZ_to_MotorCounts() As Long

    End Function

    Private Function CalcPointingState(RACounts As Int32, DECCounts As Int32) As PierSide
        Dim HACounts As Double
        Dim STCounts As Int32
        Dim DECRange As Int32

        STCounts = MountRACCounts / 4
        DECRange = MountDECCounts / 2
        HACounts = STCounts - RACounts
        'WpE Normal - POINT A
        If ((HACounts >= 0) And (HACounts <= STCounts)) And ((DECCounts >= 0) And (DECCounts <= DECRange)) Then
            WpE_Normal = True
            EpW_Normal = False
            WpE_TtP = False
            EpW_TtP = False
            Return PierSide.pierWest

            'WpE TtP - POINT B
        ElseIf ((HACounts >= 0) And (HACounts <= STCounts)) And ((DECCounts < 0) And (DECCounts >= -DECRange)) Then
            WpE_Normal = False
            EpW_Normal = False
            WpE_TtP = True
            EpW_TtP = False
            Return PierSide.pierEast

            'EpW Normal - POINT D
        ElseIf ((HACounts < 0) And (HACounts >= -STCounts)) And ((DECCounts < 0) And (DECCounts >= -DECRange)) Then
            WpE_Normal = False
            EpW_Normal = True
            WpE_TtP = False
            EpW_TtP = False
            Return PierSide.pierEast

            'EpW TtP - POINT C
        ElseIf ((HACounts < 0) And (HACounts >= -STCounts)) And ((DECCounts >= 0) And (DECCounts <= DECRange)) Then
            WpE_Normal = False
            EpW_Normal = False
            WpE_TtP = False
            EpW_TtP = True
            Return PierSide.pierWest

        End If


    End Function

#End Region

End Class
