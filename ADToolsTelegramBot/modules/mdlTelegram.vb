
Imports Telegram
Imports Telegram.Bot.Args
Imports Telegram.Bot.Types

Module mdlTelegram

    Public WithEvents Bot As Bot.TelegramBotClient

    Private Sub Bot_OnMessage(sender As Object, e As MessageEventArgs) Handles Bot.OnMessage
        If e Is Nothing Then Exit Sub
        Log("Message From " & e.Message.Chat.Id & " / " & e.Message.Chat.Username & " : " & e.Message.Text)

        If e.Message.Chat.Username = TelegramUsername Then
            ' Если это свои тогда обработать запрос и ответить
            Bot.SendChatActionAsync(e.Message.Chat.Id, Enums.ChatAction.Typing)
            OnMessage(e.Message)
        Else
            ' Если это не понятно кто тогда приподпослать
            Bot.SendTextMessageAsync(e.Message.Chat.Id, "Sorry, I don't know you...")
        End If
    End Sub

    Private Sub Bot_OnCallbackQuery(sender As Object, e As CallbackQueryEventArgs) Handles Bot.OnCallbackQuery
        If e Is Nothing Then Exit Sub
        Log("CallbackQuery From " & e.CallbackQuery.Message.Chat.Id & " / " & e.CallbackQuery.Message.From.Username & " : " & e.CallbackQuery.Message.Text)

        If e.CallbackQuery.Message.Chat.Username = TelegramUsername Then
            OnCallbackQuery(e.CallbackQuery)
        Else
            ' Если это не понятно кто тогда игнорить
        End If
    End Sub

    Private Sub Bot_OnInlineQuery(sender As Object, e As InlineQueryEventArgs) Handles Bot.OnInlineQuery
        If e Is Nothing Then Exit Sub
        Log("InlineQuery From " & e.InlineQuery.From.Id & " / " & e.InlineQuery.From.Username & " : " & e.InlineQuery.Query)

        If e.InlineQuery.From.Username = TelegramUsername Then
            ' Если это свои тогда обработать запрос и ответить
            OnInlineQuery(e.InlineQuery)
        Else
            ' Если это не понятно кто тогда приподпослать
            Bot.AnswerInlineQueryAsync(e.InlineQuery.Id, New List(Of InlineQueryResults.InlineQueryResultBase))
        End If
    End Sub

    Private Sub Bot_OnReceiveError(sender As Object, e As ReceiveErrorEventArgs) Handles Bot.OnReceiveError
        Log("ReceiveError: " & e.ApiRequestException.Message)
    End Sub

End Module
