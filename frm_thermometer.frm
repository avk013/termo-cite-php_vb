VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frm_thermometer 
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   Icon            =   "frm_thermometer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   9825
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   2760
      Width           =   9255
   End
   Begin VB.Timer Timer3 
      Interval        =   6000
      Left            =   4080
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   4080
      Top             =   240
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1095
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   5895
      ExtentX         =   10398
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Frame Frame4 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "outside"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "indoor"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label label_temperature_outside 
         Alignment       =   2  'Center
         Caption         =   "--.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label label_temperature_indoor 
         Alignment       =   2  'Center
         Caption         =   "--.-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1200
      Top             =   4200
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2760
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   3
      DTREnable       =   -1  'True
      RTSEnable       =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   4680
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frm_thermometer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const minut = 30 ' ïîñûëêà íà ñåðâåð
Const timer1interval = 6000 '1 ðàç â ìèíóòó íà ôîðìå
Const comport = 1
Const site = "http://fei.idgu.edu.ua/infr/temp/add.php?"
Dim k As Integer
Dim temp0, temp1 As Double
Dim url As String

Private Sub Form_Load()
    update_com_port
    update_sample_rate
Timer2.Interval = 40000
Timer3.Interval = 60000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    close_iic_bus
End Sub


Private Sub Timer1_Timer()
    Dim t As Double
    Dim unit As String
            unit = "°C "
    t = temperature(&H48)
  
    If t = ERROR_TEMPERATURE_NOT_READ Then
        MsgBox "Unable to read internal temperature", vbOKOnly, "Error"
        Timer1.Enabled = False
        label_temperature_indoor.Caption = "--.-" + unit
    Else
        label_temperature_indoor.Caption = Format(t, "#0.0" + unit)
   End If
   
   
        t = temperature(&H49)
        If t = ERROR_TEMPERATURE_NOT_READ Then
            MsgBox "Unable to read external temperature", vbOKOnly, "Error"
            Timer1.Enabled = False
            label_temperature_outside.Caption = "--.-" + unit
        Else
            label_temperature_outside.Caption = Format(t, "#0.0" + unit)
        End If
   

   
        frm_thermometer.Caption = label_temperature_indoor.Caption + "    ( out " + label_temperature_outside.Caption + " )"
  
   
End Sub

Private Function temperature(ByVal address As Integer) As Double

    Dim temperature_int As Long
    Dim temperature_frac As Long
    
    'For I2C bus communication, addresses are shifted one place to left,
    'as the least significant bit is used for the R/W flag.
    'In binary, shifting to left is equivalent to  multiplying by two
    address = address * 2
    
    On Error GoTo errors
    open_iic_bus MSComm1.object
    
    'an extra stop doesn't hurt...and ensures we start from a clean bus condition
    IIC_stop
    
    'read sequence, as per DS1621 datasheet
    IIC_start                           'Bus Master initiates a START condition.
    IIC_tx_byte address                 'Bus Master sends DS1621 address; R/ W= 0 (DS1621 generates acknowledge bit).
    IIC_tx_byte &HAC                    'Bus Master sends Access Config command protocol.DS1621 generates acknowledge bit.
    IIC_tx_byte &H1                     'Bus Master sets up DS1621 for output polarity active low, one-shot conversion.
                                        'DS1621 generates acknowledge bit.
   
    IIC_start                           'Bus Master generates a repeated START condition.
    IIC_tx_byte address                 'Bus Master sends DS1621 address; R/ W= 0.DS1621 generates acknowledge bit.
    IIC_tx_byte &HEE                    'Bus Master sends Start Convert T command protocol.DS1621 generates acknowledge bit.
    IIC_stop                            'Bus Master initiates STOP condition.

    
    IIC_start                           'Bus Master initiates a START condition.
    IIC_tx_byte address                 'Bus Master sends DS1621 address; R/ W= 0 (DS1621 generates acknowledge bit).
    IIC_tx_byte &HAA                    'Bus Master sends Read Temperature command protocol.DS1621 generates acknowledge bit.
    IIC_start                           'Bus Master generates a repeated START condition.
    IIC_tx_byte address + 1             'Bus Master sends DS1621 address; R/ W= 1 = READING (DS1621 generates acknowledge bit).
    temperature_int = IIC_rx_byte(1)    'Bus Master receives first byte of data and generates acknowledge.
    
    temperature_frac = IIC_rx_byte(0)   'Bus Master receives second byte of data from DS162 and does not generate acknowledge to signal end of reception.
    IIC_stop                            'Bus Master initiates STOP condition.
    
    'some bynary math to convert to a data format Visual Basic can understand
    temperature = (temperature_int * 256 + temperature_frac) / 128 * 5 / 10
    If temperature_int >= 128 Then
        temperature = temperature - 256
    End If
    Exit Function

errors:
    temperature = ERROR_TEMPERATURE_NOT_READ
End Function

Private Function update_sample_rate()
            Timer1.Interval = timer1interval
End Function
Private Function update_com_port()
    If MSComm1.PortOpen Then
        MSComm1.PortOpen = False
    End If
        
        MSComm1.CommPort = comport
        Timer1.Enabled = True
        
End Function







Private Sub Timer2_Timer()
tempr
End Sub

Private Sub Timer3_Timer()
k = k + 1
Label2 = k
If k >= 1 Then If Timer2.Enabled = True Then Timer2.Enabled = False
If k >= minut Then
If Timer2.Enabled = True Then Timer2.Enabled = False Else Timer2.Enabled = True
k = 0
End If
End Sub

Sub tempr()
temp0 = temperature(&H48)
temp0 = temperature(&H48)
temp1 = temperature(&H49)
'temp0 = temp0 + 1
'temp1 = temp1 + 1
url = site + "usr=a1s1d5f9ss&temp0=" + Str(temp0) + "&temp1=" + Str(temp1) + "&dt=" + Str(Time)
Text1 = url
WebBrowser1.Navigate url
End Sub
