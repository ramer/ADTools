
Imports Telegram
Imports Telegram.Bot.Types

Module mdlTelegram

    Public Bot As Bot.TelegramBotClient
    Public lastupdateid As Integer

    Public Sub SendTelegramMessage(ChatId As Integer, Message As String, Optional keyboard As IEnumerable(Of IEnumerable(Of ReplyMarkups.KeyboardButton)) = Nothing)
        If Bot Is Nothing Then Exit Sub

        Dim response As Message = Nothing

        If keyboard IsNot Nothing Then
            response = Bot.SendTextMessageAsync(
                ChatId, Message, Enums.ParseMode.Default, True, False, 0,
                New ReplyMarkups.ReplyKeyboardMarkup With {.Keyboard = keyboard, .ResizeKeyboard = True},
                Nothing).Result
        Else
            response = Bot.SendTextMessageAsync(
                ChatId, Message, Enums.ParseMode.Default, True, False, 0,
                New ReplyMarkups.ReplyKeyboardRemove,
                Nothing).Result
        End If

        Log("[To   " & ChatId & " / " & response.From.Username & "] : " & Message)
    End Sub

    Public Sub GetTelegramMessages()
        If Bot Is Nothing Then Exit Sub

        Dim updates As Update() = Bot.GetUpdatesAsync(lastupdateid + 1,,, {Enums.UpdateType.Message}).Result
        Dim update As Update

        For Each update In updates

            If lastupdateid < update.Id Then lastupdateid = update.Id 'записываем номер последнего сообщения чтобы избежать повторного чтения

            If update.Message Is Nothing Then Continue For
            Log("[From " & update.Message.From.Id & " / " & update.Message.From.Username & "] : " & update.Message.Text)

            If TelegramUsername = update.Message.From.Username Then
                ' Если это свои тогда обработать запрос и ответить
                If Bot IsNot Nothing Then Bot.SendChatActionAsync(update.Message.From.Id, Enums.ChatAction.Typing)
                ProcessDialogue(update)
            Else
                ' Если это не понятно кто тогда приподпослать
                SendTelegramMessage(update.Message.From.Id, "Sorry, I don't know you...")
            End If
        Next
    End Sub

End Module
