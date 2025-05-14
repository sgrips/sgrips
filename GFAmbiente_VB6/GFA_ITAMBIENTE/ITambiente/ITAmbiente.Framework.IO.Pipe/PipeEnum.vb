Imports System.ComponentModel

Public Enum PipeStatusEnum
    PipeWaitConnecting
    PipeConnected
    PipeWaitDisconnecting
    PipeDisconnected
End Enum

Public Enum MessagePipeIDEnum
    MessagePipeClientInfo
    MessagePipeClientName
    MessagePipeClientRefresh
    MessagePipeClientStatus
    MessagePipeClientEnd
    MessagePipeServerEnd
End Enum