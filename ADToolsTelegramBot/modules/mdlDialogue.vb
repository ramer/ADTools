Imports Telegram.Bot.Types
Imports Telegram.Bot.Types.ReplyMarkups

Module mdlDialogue

    Private currentuser As clsDirectoryObject
    Private currentgroup As clsDirectoryObject
    Private searcher As New clsSearcher

    Public Enum DialogueStage
        SearchUser
        UserConfirmResetPassword
        UserConfirmEnableDisable
        SearchGroup
        GroupConfirmMemberOf
    End Enum


    Const DIALOGUE_BUTTON_YES = "✅ канеш"
    Const DIALOGUE_BUTTON_NO = "❌ передумал"

    'Const DIALOGUE_WARNING = "❗️"

    Const DIALOGUE_INLINEBUTTON_DETAILS = "ℹ️"
    Const DIALOGUE_INLINEBUTTON_RESETPASSWORD = "🔑"
    Const DIALOGUE_INLINEBUTTON_ENABLE = "🍏"
    Const DIALOGUE_INLINEBUTTON_DISABLE = "🍎"
    Const DIALOGUE_INLINEBUTTON_GROUPS = "👥"

    Private confimkeyboard As New List(Of KeyboardButton) From {New KeyboardButton(DIALOGUE_BUTTON_YES), New KeyboardButton(DIALOGUE_BUTTON_NO)}

    Public Property Stage As DialogueStage = DialogueStage.SearchUser

    Public Sub OnMessage(message As Message)

        If message.Text = "/start" Then

            Stage = DialogueStage.SearchUser
            SendRequestStageGreeting(message)

        ElseIf message.Text = DIALOGUE_BUTTON_YES Then
            Select Case Stage
                Case DialogueStage.UserConfirmResetPassword

                    Stage = DialogueStage.SearchUser
                    SendRequestResetPasswordCompleted(message)

                Case DialogueStage.UserConfirmEnableDisable

                    Stage = DialogueStage.SearchUser
                    SendRequestEnableDisableCompleted(message)

                Case DialogueStage.GroupConfirmMemberOf

                    Stage = DialogueStage.SearchUser
                    SendRequestStageGroupMemberOfCompleted(message)

                Case Else

                    Stage = DialogueStage.SearchUser
                    SendRequestStageSearchUser(message)

            End Select

        ElseIf message.Text = DIALOGUE_BUTTON_NO Then

            Stage = DialogueStage.SearchUser
            Bot.SendTextMessageAsync(message.Chat.Id, "Ну и ладно...", Enums.ParseMode.Markdown,,,, New ReplyKeyboardRemove)

        Else ' нажали не кнопку

            Select Case Stage
                Case DialogueStage.SearchUser

                    Dim results As New List(Of clsDirectoryObject)
                    For Each dmn In domains
                        results.AddRange(searcher.SearchSync(New clsDirectoryObject(dmn.DefaultNamingContext, dmn), New clsFilter(message.Text, attributesForSearchDefault, New clsSearchObjectClasses(True, False, False, False, False))))
                    Next

                    SendRequestUserList(message, results)

                Case DialogueStage.SearchGroup

                    If currentuser Is Nothing Then
                        Stage = DialogueStage.SearchUser
                        SendRequestStageSearchUser(message)
                        Exit Sub
                    End If

                    Dim results As New List(Of clsDirectoryObject)
                    For Each dmn In domains
                        results.AddRange(searcher.SearchSync(New clsDirectoryObject(dmn.DefaultNamingContext, dmn), New clsFilter("*" & message.Text & "*", attributesForSearchDefault, New clsSearchObjectClasses(False, False, False, True, False))))
                    Next

                    SendRequestGroupList(message, results)

                Case Else

                    SendRequestUnexpectedRequest(message)

            End Select

        End If

    End Sub

    Public Sub OnCallbackQuery(query As CallbackQuery)
        Dim data = query.Data.Split({";"}, StringSplitOptions.RemoveEmptyEntries)
        If data.Length <> 2 Then Exit Sub
        Dim action = data(0)

        Dim responceguid As Guid = Nothing
        Try
            responceguid = New Guid(Decode58(data(1)))
        Catch
            Exit Sub
        End Try
        If responceguid = Nothing Then Exit Sub

        Dim guidresults As New List(Of clsDirectoryObject)
        For Each dmn In domains
            guidresults.AddRange(searcher.SearchSync(New clsDirectoryObject(dmn.DefaultNamingContext, dmn), New clsFilter(responceguid.ToString, attributesForSearchDefault, New clsSearchObjectClasses(True, False, False, False, False))))
        Next
        Dim obj As clsDirectoryObject = If(guidresults.Count = 1, guidresults(0), Nothing)
        If obj Is Nothing Then Exit Sub

        If obj.SchemaClass = enmDirectoryObjectSchemaClass.User Then

            currentuser = obj

            If action = "Details" Then

                SendRequestUserDetails(query.Message, obj)

            ElseIf action = "Pass" Then

                Stage = DialogueStage.UserConfirmResetPassword
                SendRequestResetPasswordConfirm(query.Message)

            ElseIf action = "Enable" Then

                Stage = DialogueStage.UserConfirmEnableDisable
                SendRequestEnableDisableConfirm(query.Message)

            ElseIf action = "Disable" Then

                Stage = DialogueStage.UserConfirmEnableDisable
                SendRequestEnableDisableConfirm(query.Message)

            ElseIf action = "Groups" Then

                Stage = DialogueStage.SearchGroup
                SendRequestUserGroups(query.Message)

            Else

                SendRequestUnexpectedInlineButton(query.Message)

            End If

        ElseIf obj.SchemaClass = enmDirectoryObjectSchemaClass.Group Then

            If currentuser Is Nothing Then
                Stage = DialogueStage.SearchUser
                SendRequestUnexpectedInlineButton(query.Message)
                Exit Sub
            End If

            currentgroup = obj

            If action = "MemberOf" Then

                Stage = DialogueStage.GroupConfirmMemberOf
                SendRequestConfirmMemberOf(query.Message)

            Else

                SendRequestUnexpectedInlineButton(query.Message)

            End If

        Else

            SendRequestUnexpectedInlineButton(query.Message)

        End If

    End Sub

    Public Sub OnInlineQuery(inlinequery As InlineQuery)
        Dim queryresults As New List(Of InlineQueryResults.InlineQueryResultBase)

        If inlinequery.Query.Length >= 3 Then
            Dim results As New List(Of clsDirectoryObject)
            For Each dmn In domains
                results.AddRange(searcher.SearchSync(New clsDirectoryObject(dmn.DefaultNamingContext, dmn), New clsFilter(inlinequery.Query, attributesForSearchDefault, New clsSearchObjectClasses(True, False, False, False, False))))
            Next

            Dim count = 0
            For Each obj In results
                count += 1
                If count = 10 Then Exit For

                Dim msg As String = ""
                InsertUser(msg, obj)
                msg &= If(String.IsNullOrEmpty(obj.userPrincipalNameName), "", "🎫 " & obj.userPrincipalNameName & nl)
                msg &= If(String.IsNullOrEmpty(obj.physicalDeliveryOfficeName), "", "🏠 " & obj.physicalDeliveryOfficeName & nl)
                msg &= If(String.IsNullOrEmpty(obj.telephoneNumber), "", "📞 " & obj.telephoneNumber & nl)
                msg &= If(String.IsNullOrEmpty(obj.mail), "", "✉️ " & obj.mail & nl)
                msg &= If(String.IsNullOrEmpty(obj.title), "", "🗄 " & obj.title & nl)
                msg &= If(String.IsNullOrEmpty(obj.passwordExpiresFormated), "", "🔑 " & obj.passwordExpiresFormated & nl)

                Dim queryresultmessage = New InlineQueryResults.InputTextMessageContent(msg)
                queryresultmessage.DisableWebPagePreview = True
                queryresultmessage.ParseMode = Enums.ParseMode.Markdown

                Dim queryresultinlinekeyboard = New List(Of InlineKeyboardButton)

                Dim queryobj = New InlineQueryResults.InlineQueryResultArticle(Encode58(obj.objectGUID.ToByteArray), obj.name, queryresultmessage)
                queryobj.ReplyMarkup = New InlineKeyboardMarkup(queryresultinlinekeyboard)
                queryobj.Description = obj.userPrincipalName & nl & obj.title

                If obj.Status = enmDirectoryObjectStatus.Normal Then
                    queryobj.ThumbUrl = "http://icons.iconarchive.com/icons/custom-icon-design/flatastic-11/64/User-green-icon.png"
                ElseIf obj.Status = enmDirectoryObjectStatus.Expired Then
                    queryobj.ThumbUrl = "http://icons.iconarchive.com/icons/custom-icon-design/flatastic-11/64/User-yellow-icon.png"
                ElseIf obj.Status = enmDirectoryObjectStatus.Blocked Then
                    queryobj.ThumbUrl = "http://icons.iconarchive.com/icons/custom-icon-design/flatastic-11/64/User-red-icon.png"
                End If

                queryresults.Add(queryobj)
            Next

        End If

        Bot.AnswerInlineQueryAsync(inlinequery.Id, queryresults)
    End Sub

    Private Sub SendRequestStageGreeting(message As Message)
        Bot.SendTextMessageAsync(message.Chat.Id, String.Format(
        "Привет, {0}!" & nl &
        "Я бот ADTools." & nl &
        "Кого ищем?", message.Chat.FirstName), Enums.ParseMode.Markdown,,,, New ReplyKeyboardRemove)
    End Sub

    Private Sub SendRequestUnexpectedInlineButton(message As Message)
        Bot.SendTextMessageAsync(message.Chat.Id, "Это чо за кнопка еще?!", Enums.ParseMode.Markdown)
    End Sub

    Private Sub SendRequestUnexpectedRequest(message As Message)
        Bot.SendTextMessageAsync(message.Chat.Id, "Чойта??? Нажми кнопку!", Enums.ParseMode.Markdown)
    End Sub

    Private Sub SendRequestStageSearchUser(message As Message)
        Bot.SendTextMessageAsync(message.Chat.Id, "Кого ищем?", Enums.ParseMode.Markdown,,,, New ReplyKeyboardRemove)
    End Sub

    Private Sub SendRequestUserList(message As Message, objects As List(Of clsDirectoryObject))

        If objects.Count = 0 Then

            Bot.SendTextMessageAsync(message.Chat.Id, "Никого не найдено", Enums.ParseMode.Markdown,,,, New ReplyKeyboardRemove)

        ElseIf objects.Count > 10 Then

            Bot.SendTextMessageAsync(message.Chat.Id, "Найдено больше 10 объектов, пешы ищё", Enums.ParseMode.Markdown,,,, New ReplyKeyboardRemove)

        Else

            For Each obj In objects
                SendRequestUser(message, obj)
            Next

        End If
    End Sub

    Private Sub SendRequestGroupList(message As Message, objects As List(Of clsDirectoryObject))

        If objects.Count = 0 Then

            Bot.SendTextMessageAsync(message.Chat.Id, "Группа не найдена", Enums.ParseMode.Markdown)

        ElseIf objects.Count > 10 Then

            Bot.SendTextMessageAsync(message.Chat.Id, "Найдено больше 10 объектов, пешы ищё", Enums.ParseMode.Markdown)

        Else

            Dim groupinlinekeyboard As New List(Of List(Of InlineKeyboardButton))
            For Each group In objects
                groupinlinekeyboard.Add(New List(Of InlineKeyboardButton) From {New InlineKeyboardButton With {.Text = "👥 " & group.name, .CallbackData = "MemberOf;" & Encode58(group.objectGUID.ToByteArray)}})
            Next

            Dim msg As String = "В какую группу добавить?"

            Bot.SendTextMessageAsync(message.Chat.Id, msg, Enums.ParseMode.Markdown,,,, New InlineKeyboardMarkup(groupinlinekeyboard))

        End If
    End Sub

    Private Sub SendRequestUser(message As Message, obj As clsDirectoryObject)
        If obj Is Nothing Then Exit Sub

        Dim userinlinekeyboard As New List(Of InlineKeyboardButton) From {
            New InlineKeyboardButton With {.Text = DIALOGUE_INLINEBUTTON_DETAILS, .CallbackData = "Details;" & Encode58(obj.objectGUID.ToByteArray)},
            New InlineKeyboardButton With {.Text = DIALOGUE_INLINEBUTTON_RESETPASSWORD, .CallbackData = "Pass;" & Encode58(obj.objectGUID.ToByteArray)},
            New InlineKeyboardButton With {.Text = If(obj.disabled, DIALOGUE_INLINEBUTTON_DISABLE, DIALOGUE_INLINEBUTTON_ENABLE), .CallbackData = If(obj.disabled, "Enable;" & Encode58(obj.objectGUID.ToByteArray), "Disable;" & Encode58(obj.objectGUID.ToByteArray))},
            New InlineKeyboardButton With {.Text = DIALOGUE_INLINEBUTTON_GROUPS, .CallbackData = "Groups;" & Encode58(obj.objectGUID.ToByteArray)}}

        Dim msg As String = "*" & obj.name & "*" & nl
        msg &= If(String.IsNullOrEmpty(obj.userPrincipalNameName), "", "🎫 " & obj.userPrincipalNameName & nl)

        Bot.SendTextMessageAsync(message.Chat.Id, msg, Enums.ParseMode.Markdown,,,, New InlineKeyboardMarkup(userinlinekeyboard))
    End Sub

    Private Sub SendRequestUserDetails(message As Message, obj As clsDirectoryObject)
        If obj Is Nothing Then Exit Sub

        Dim userinlinekeyboard As New List(Of InlineKeyboardButton) From {
            New InlineKeyboardButton With {.Text = DIALOGUE_INLINEBUTTON_DETAILS, .CallbackData = "Details;" & Encode58(obj.objectGUID.ToByteArray)},
            New InlineKeyboardButton With {.Text = DIALOGUE_INLINEBUTTON_RESETPASSWORD, .CallbackData = "Pass;" & Encode58(obj.objectGUID.ToByteArray)},
            New InlineKeyboardButton With {.Text = If(obj.disabled, DIALOGUE_INLINEBUTTON_DISABLE, DIALOGUE_INLINEBUTTON_ENABLE), .CallbackData = If(obj.disabled, "Enable;" & Encode58(obj.objectGUID.ToByteArray), "Disable;" & Encode58(obj.objectGUID.ToByteArray))},
            New InlineKeyboardButton With {.Text = DIALOGUE_INLINEBUTTON_GROUPS, .CallbackData = "Groups;" & Encode58(obj.objectGUID.ToByteArray)}}

        Dim msg As String = ""
        InsertUser(msg, obj)
        msg &= If(String.IsNullOrEmpty(obj.userPrincipalNameName), "", "🎫 " & obj.userPrincipalNameName & nl)
        msg &= If(String.IsNullOrEmpty(obj.physicalDeliveryOfficeName), "", "🏠 " & obj.physicalDeliveryOfficeName & nl)
        msg &= If(String.IsNullOrEmpty(obj.telephoneNumber), "", "📞 " & obj.telephoneNumber & nl)
        msg &= If(String.IsNullOrEmpty(obj.mail), "", "✉️ " & obj.mail & nl)
        msg &= If(String.IsNullOrEmpty(obj.title), "", "🗄 " & obj.title & nl)
        msg &= If(String.IsNullOrEmpty(obj.passwordExpiresFormated), "", "🔑 " & obj.passwordExpiresFormated & nl)

        Bot.SendTextMessageAsync(message.Chat.Id, msg, Enums.ParseMode.Markdown,,,, New InlineKeyboardMarkup(userinlinekeyboard))
    End Sub

    Private Sub SendRequestResetPasswordConfirm(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = "Прям сбросить пароль??" & nl
        InsertUser(msg, currentuser)

        Bot.SendTextMessageAsync(message.Chat.Id, msg, Enums.ParseMode.Markdown,,,, New ReplyKeyboardMarkup(confimkeyboard, True, True))
    End Sub

    Private Sub SendRequestEnableDisableConfirm(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""

        If currentuser.disabled = True Then
            msg &= String.Format("Разблочить?" & nl)
        Else
            msg &= String.Format("Заблочить?" & nl)
        End If
        InsertUser(msg, currentuser)

        Bot.SendTextMessageAsync(message.Chat.Id, msg, Enums.ParseMode.Markdown,,,, New ReplyKeyboardMarkup(confimkeyboard, True, True))
    End Sub

    Private Sub SendRequestResetPasswordCompleted(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""
        Try

            currentuser.ResetPassword()
            currentuser.passwordNeverExpires = False
            msg = "Пароль сброшен." & nl
            InsertUser(msg, currentuser)

        Catch ex As Exception

            msg = "Не получилось сбросить пароль." & nl
            InsertUser(msg, currentuser)
            msg &= ex.Message

        End Try

        Bot.SendTextMessageAsync(message.Chat.Id, msg, Enums.ParseMode.Markdown,,,, New ReplyKeyboardRemove)
    End Sub

    Private Sub SendRequestEnableDisableCompleted(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""
        Try

            If currentuser.disabled Then
                currentuser.disabled = False
                msg &= "Разблокирован" & nl
            Else
                currentuser.disabled = True
                msg &= "Заблокирован" & nl
            End If
            InsertUser(msg, currentuser)

        Catch ex As Exception

            msg = "Не получилось заблочить/разблочить." & nl
            InsertUser(msg, currentuser)
            msg &= ex.Message

        End Try

        Bot.SendTextMessageAsync(message.Chat.Id, msg, Enums.ParseMode.Markdown,,,, New ReplyKeyboardRemove)
    End Sub

    Private Sub SendRequestUserGroups(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim groupinlinekeyboard As New List(Of List(Of InlineKeyboardButton))

        Dim msg As String = ""
        msg &= "Введи кусок названия группы для добавления" & nl

        If currentuser.memberOf.Count > 0 Then
            msg &= "или выбери какую удалить:"
            For Each group As clsDirectoryObject In currentuser.memberOf
                groupinlinekeyboard.Add(New List(Of InlineKeyboardButton) From {New InlineKeyboardButton With {.Text = "👥 " & group.name, .CallbackData = "MemberOf;" & Encode58(group.objectGUID.ToByteArray)}})
            Next
        End If

        Bot.SendTextMessageAsync(message.Chat.Id, msg, Enums.ParseMode.Markdown,,,, New InlineKeyboardMarkup(groupinlinekeyboard))
    End Sub

    Private Sub SendRequestConfirmMemberOf(message As Message)
        If currentuser Is Nothing Or currentgroup Is Nothing Then Exit Sub

        Dim msg As String = ""

        Dim newgroup As Boolean = True
        For Each group As clsDirectoryObject In currentuser.memberOf
            If group.name = currentgroup.name Then newgroup = False
        Next

        If newgroup Then

            msg &= "Добавить" & nl
            InsertUser(msg, currentuser)
            msg &= "в группу" & nl
            msg &= "👥 *" & currentgroup.name & "*" & nl
            msg &= "ммм?"

        Else

            msg &= "Удалить" & nl
            InsertUser(msg, currentuser)
            msg &= "из группы" & nl
            msg &= "👥 *" & currentgroup.name & "*" & nl
            msg &= "ммм?"

        End If

        Bot.SendTextMessageAsync(message.Chat.Id, msg, Enums.ParseMode.Markdown,,,, New ReplyKeyboardMarkup(confimkeyboard, True, True))
    End Sub

    Private Sub SendRequestStageGroupMemberOfCompleted(message As Message)
        If currentuser Is Nothing Or currentgroup Is Nothing Then Exit Sub

        Dim msg As String = ""

        Dim newgroup As Boolean = True
        For Each group As clsDirectoryObject In currentuser.memberOf
            If group.name = currentgroup.name Then newgroup = False
        Next

        Try

            If newgroup Then

                currentgroup.UpdateAttribute(DirectoryServices.Protocols.DirectoryAttributeOperation.Add, "member", currentuser.distinguishedName)
                currentuser.memberOf.Add(currentgroup)

                InsertUser(msg, currentuser)
                msg &= "добавлен в группу" & nl
                msg &= "👥 *" & currentgroup.name & "*" & nl

            Else

                currentgroup.UpdateAttribute(DirectoryServices.Protocols.DirectoryAttributeOperation.Delete, "member", currentuser.distinguishedName)
                currentuser.memberOf.Remove(currentgroup)
                For Each group As clsDirectoryObject In currentuser.memberOf
                    If group.name = currentgroup.name Then currentuser.memberOf.Remove(group) : Exit For
                Next

                InsertUser(msg, currentuser)
                msg &= "удален из группы" & nl
                msg &= "👥 *" & currentgroup.name & "*" & nl

            End If

        Catch ex As Exception

            msg = "Не получилось добавить/удалить" & nl
            InsertUser(msg, currentuser)
            msg &= "в/из группы" & nl
            msg &= "👥 *" & currentgroup.name & "*" & nl
            msg &= ex.Message

        End Try

        Bot.SendTextMessageAsync(message.Chat.Id, msg, Enums.ParseMode.Markdown,,,, New ReplyKeyboardRemove)
    End Sub

    Private Sub InsertUser(ByRef msg As String, obj As clsDirectoryObject)
        If obj Is Nothing Then Return
        If obj.Status = enmDirectoryObjectStatus.Normal Then
            msg &= "🍏 *" & obj.name & "*" & nl
        ElseIf obj.Status = enmDirectoryObjectStatus.Expired Then
            msg &= "🍋 *" & obj.name & "*" & nl
        ElseIf obj.Status = enmDirectoryObjectStatus.Blocked Then
            msg &= "🍎 *" & obj.name & "*" & nl
        End If
    End Sub

End Module
