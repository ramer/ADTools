Imports TeleBotDotNet
Imports TeleBotDotNet.Requests.Methods
Imports TeleBotDotNet.Responses.Methods

Module mdlTelegram

    Public Bot As TeleBot
    Public lastupdateid As Integer

    Public Sub SendTelegramMessage(ChatId As Integer, Message As String, Optional keyboard As List(Of List(Of String)) = Nothing, Optional hidekeyboard As Boolean = False)
        If Bot Is Nothing Then Exit Sub

        Dim sendMsgRequest As New SendMessageRequest
        sendMsgRequest.Text = Message
        sendMsgRequest.ChatId = ChatId

        If keyboard IsNot Nothing Then
            sendMsgRequest.ReplyMarkup = New Requests.Types.ReplyMarkupRequest With {
                .ReplyMarkupReplyKeyboardMarkup = New Requests.Types.ReplyKeyboardMarkupRequest With {.Keyboard = keyboard, .ResizeKeyboard = True}}
        ElseIf hidekeyboard = True Then
            sendMsgRequest.ReplyMarkup = New Requests.Types.ReplyMarkupRequest With {
                .ReplyMarkupReplyKeyboardHide = New Requests.Types.ReplyKeyboardHideRequest}
        End If

        Dim sendMsgResponce As SendMessageResponse
        sendMsgResponce = Bot.SendMessage(sendMsgRequest)

        If sendMsgResponce.Ok = True Then
            Log("Answer: " & Message)
        Else
            Log("Error while sending message: " & Message)
        End If
    End Sub

    Public Sub GetTelegramMessages()
        If Bot Is Nothing Then Exit Sub

        Dim getUpdRequest As New GetUpdatesRequest
        Dim getUpdResponse As GetUpdatesResponse
        Dim responce As Responses.Types.UpdateResponse

        getUpdRequest.Offset = lastupdateid + 1
        getUpdResponse = Bot.GetUpdates(getUpdRequest)

        For Each responce In getUpdResponse.Result
            If lastupdateid < responce.UpdateId Then lastupdateid = responce.UpdateId 'записываем номер последнего сообщения чтобы избежать повторного чтения

            If responce.Message Is Nothing Then Continue For
            Log("[" & responce.UpdateId & "] " & responce.Message.From.UserName & ": " & responce.Message.Text)

            If TelegramUsername = responce.Message.From.UserName Then
                ' Если это свои тогда обработать запрос и ответить
                ProcessDialogue(responce)

            Else
                ' Если это не понятно кто тогда приподпослать
                SendTelegramMessage(responce.Message.From.Id, "Sorry, I don't know you...")
            End If
        Next
    End Sub

End Module
