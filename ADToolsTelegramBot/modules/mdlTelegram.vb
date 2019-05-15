
Imports Telegram
Imports Telegram.Bot.Args
Imports Telegram.Bot.Types

Module mdlTelegram

    Public WithEvents Bot As Bot.TelegramBotClient

    Private Async Sub Bot_OnMessage(sender As Object, e As MessageEventArgs) Handles Bot.OnMessage
        If e Is Nothing Then Exit Sub
        Log("Message From " & e.Message.Chat.Id & " / " & e.Message.Chat.Username & " : " & e.Message.Text)

        If e.Message.Chat.Username = TelegramUsername Then
            ' Если это свои тогда обработать запрос и ответить
            Await Bot.SendChatActionAsync(e.Message.Chat.Id, Enums.ChatAction.Typing)
            OnMessage(e.Message)
        Else
            ' Если это не понятно кто тогда приподпослать
            Await Bot.SendTextMessageAsync(e.Message.Chat.Id, "Sorry, I don't know you...")
        End If
    End Sub

    Private Sub Bot_OnCallbackQuery(sender As Object, e As CallbackQueryEventArgs) Handles Bot.OnCallbackQuery
        If e Is Nothing Then Exit Sub
        Log("CallbackQuery From " & e.CallbackQuery.Message.Chat.Id & " / " & e.CallbackQuery.Message.From.Username & " : " & e.CallbackQuery.Message.Text)

        'If e.CallbackQuery.Message.Chat.Username = TelegramUsername Then
        '    dialogues(e.From.Username).OnCallbackQuery(e)
        'Else
        '    ' Если это не понятно кто тогда приподпослать
        '    Await Bot.SendTextMessageAsync(e.From.Id, "Sorry, I don't know you...")
        'End If
    End Sub



End Module
