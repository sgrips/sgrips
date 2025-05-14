VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"message"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarMessage As String 'local copy
Private mvarCheckSum As Long 'local copy
Private mvarSTART_CHAR As Boolean 'local copy
Private mvarEND_CHAR As Boolean
Private mvarSTX As Boolean 'local copy
Private mvarETX As Boolean 'local copy
Private mvarCR As Boolean 'local copy
Private mvarLF As Boolean 'local copy
Private mvarchk1 As Long 'local copy
Private mvarchk2 As Long 'local copy
Private mvarCheckSumON As Boolean 'local copy
'local variable(s) to hold property value(s)
Private mvarSTXETX_ON As Boolean 'local copy
Private mvarSTARCHAR_ON As Boolean
Private mvarENDCHAR_ON As Boolean
'local variable(s) to hold property value(s)
Private mvarLineFeedON As Boolean 'local copy
Private mvarCarriageReturnON As Boolean

Public Property Let CarriageReturnON(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LineFeedON = 5
    mvarCarriageReturnON = vData
End Property


Public Property Get CarriageReturnON() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LineFeedON
    CarriageReturnON = mvarCarriageReturnON
End Property


Public Property Let LineFeedON(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LineFeedON = 5
    mvarLineFeedON = vData
End Property





Public Property Get LineFeedON() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LineFeedON
    LineFeedON = mvarLineFeedON
End Property


Public Property Let STXETX_ON(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.STXETX_ON = 5
    mvarSTXETX_ON = vData
End Property


Public Property Get STXETX_ON() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.STXETX_ON
    STXETX_ON = mvarSTXETX_ON
End Property

Public Property Get STARCHAR_ON() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.STXETX_ON
    STARCHAR_ON = mvarSTARCHAR_ON
End Property

Public Property Let STARCHAR_ON(v As Boolean)
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.STXETX_ON
    mvarSTARCHAR_ON = v
End Property

Friend Property Let CheckSumON(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CheckSumON = 5
    mvarCheckSumON = vData
End Property


Friend Property Get CheckSumON() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CheckSumON
    CheckSumON = mvarCheckSumON
End Property



Friend Property Let chk2(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.chk2 = 5
    mvarchk2 = vData
End Property


Friend Property Get chk2() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.chk2
    chk2 = mvarchk2
End Property



Friend Property Let chk1(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.chk1 = 5
    mvarchk1 = vData
End Property


Friend Property Get chk1() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.chk1
    chk1 = mvarchk1
End Property



Friend Property Let LF(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LF = 5
    mvarLF = vData
End Property


Friend Property Get LF() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LF
    LF = mvarLF
End Property



Friend Property Let CR(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CR = 5
    mvarCR = vData
End Property


Friend Property Get CR() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CR
    CR = mvarCR
End Property



Friend Property Let ETX(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ETX = 5
    mvarETX = vData
End Property


Friend Property Get START_CHAR() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ETX
    START_CHAR = IIf(STARCHAR_ON = True, mvarSTART_CHAR, True)
End Property

Friend Property Let START_CHAR(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ETX = 5
    mvarSTART_CHAR = vData
End Property


Friend Property Get ETX() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ETX
    ETX = IIf(STXETX_ON, mvarETX, True)
End Property


Friend Property Let STX(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.STX = 5
    mvarSTX = vData
End Property


Friend Property Get STX() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.STX
    STX = IIf(STXETX_ON, mvarSTX, True)
End Property



Friend Property Let CheckSum(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CheckSum = 5
    mvarCheckSum = vData
End Property


Friend Property Get CheckSum() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CheckSum
    CheckSum = mvarCheckSum
End Property



Public Property Let message(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Message = 5
    mvarMessage = vData
End Property


Public Property Get message() As String
Attribute message.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Message
    message = mvarMessage
End Property



