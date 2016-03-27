Attribute VB_Name = "iic"
Option Explicit
Dim MSComm1 As Object

Public Function open_iic_bus(serial_port As Object) As Integer
    Set MSComm1 = serial_port
    If Not MSComm1.PortOpen Then
        MSComm1.PortOpen = True
        MSComm1.RTSEnable = True
        MSComm1.DTREnable = True
    End If
    open_iic_bus = NO_ERROR
    Exit Function
End Function

Public Sub close_iic_bus()
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
    End If
End Sub

Public Sub IIC_start()
        SDA_output
        SDA_high
        SCL_high
        SDA_low
End Sub

Public Sub IIC_stop()
        SDA_output
        SDA_low
        SCL_high
        SDA_high
End Sub

Public Function IIC_tx_byte(ByVal B As Long) As Integer
    Dim i As Integer
    For i = 0 To 7
        If (B And &H80) Then
            IIC_tx_bit_1
        Else
            IIC_tx_bit_0
        End If
        B = B * 2
    Next
    IIC_tx_byte = IIC_rx_bit
End Function

Public Function IIC_rx_byte(ByVal acknowledge As Integer) As Integer
    Dim i As Integer
    Dim retval As Integer
    For i = 0 To 7
        retval = retval * 2
        If IIC_rx_bit() Then
            retval = retval + 1
        End If
    Next
    If acknowledge Then
        IIC_tx_bit_0
    Else
        IIC_tx_bit_1
    End If
    IIC_rx_byte = retval
End Function
 
Private Sub IIC_tx_bit_1()
        SCL_low
        SDA_output
        SDA_high
        SCL_high
        SCL_low
End Sub
         
Private Sub IIC_tx_bit_0()
        SCL_low
        SDA_output
        SDA_low
        SCL_high
        SCL_low
End Sub
                         
Private Function IIC_rx_bit() As Integer
        Dim retval As Integer
        SDA_input
        SCL_low
        SCL_high
        retval = SDA_value()
        SCL_low
        IIC_rx_bit = retval
End Function
                          
Private Sub SDA_output()
    'there is no real difference between SDA_output
    'and SDA_input, because we have two unidirectional
    'lines instead of a single bidirectional one
    'This call is left here for compatibility with
    'other hardware implementations
    MSComm1.DTREnable = True
End Sub

Private Sub SDA_input()
    'since DATA is open-collector, putting DTR high sets it as PC input
    MSComm1.DTREnable = True
End Sub

Private Sub SCL_high()
    MSComm1.RTSEnable = True
    iic_wait
End Sub

Private Sub SCL_low()
    MSComm1.RTSEnable = False
    iic_wait
End Sub

Private Sub SDA_high()
    MSComm1.DTREnable = True
    iic_wait
End Sub
Private Sub SDA_low()
    MSComm1.DTREnable = False
    iic_wait
End Sub

Private Function SDA_value() As Integer
    SDA_value = MSComm1.CTSHolding
End Function


Private Sub iic_wait()
'void , insert here wait code for very fast systems
'actual measured speed is 1,5 kbps on a P90...
Randomize
frm_thermometer.Label3 = Int(100 * Rnd())
End Sub
