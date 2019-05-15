Imports Telegram
Imports Telegram.Bot.Types
Imports Telegram.Bot.Types.ReplyMarkups

Module mdlDialogue

    Private currentuser As clsDirectoryObject
    Private currentgroup As clsDirectoryObject
    Private searcher As New clsSearcher

    Public Enum DialogueStage
        SearchUser
        UserMenu
        UserConfirmResetPassword
        UserConfirmEnableDisable
        SearchGroup
        GroupConfirmMemberOf
    End Enum

    'Const DIALOGUE_BUTTON_CURRENTUSER = "👤 выбранный юзер"
    Const DIALOGUE_BUTTON_USERRESETPASSWORD = "🔑 сбросить пароль"
    Const DIALOGUE_BUTTON_USERENABLEDISABLE = "⛔️  заблочить / разблочить"
    Const DIALOGUE_BUTTON_USERMEMBEROF = "👥 член групп"
    Const DIALOGUE_BUTTON_USERDETAILS = "ℹ️ подробнее"
    'Const DIALOGUE_BUTTON_SHOWCONFIG = "ℹ️ показать настройки"
    'Const DIALOGUE_BUTTON_UPDATE = "⏺ начать перепись"
    Const DIALOGUE_BUTTON_BACK = "↪️ назад"
    Const DIALOGUE_BUTTON_YES = "✅ канеш"
    Const DIALOGUE_BUTTON_NO = "❌ передумал"
    'Const DIALOGUE_BUTTON_HOME = "🏠 главная"
    'Const DIALOGUE_BUTTON_UPDATELASTDEVICES = "Переписать последние"
    'Const DIALOGUE_BUTTON_CLEARHISTORY = "🚫 удалить историю"
    'Const DIALOGUE_BUTTON_PRINT = "🖨 печать акта"
    'Const DIALOGUE_BUTTON_COPY = "🆕 копировать"
    'Const DIALOGUE_BUTTON_SHOWPRINTERLIST = "ℹ️ список принтеров"

    'Const DIALOGUE_WARNING = "❗️"

    Private confimkeyboard As New List(Of List(Of ReplyMarkups.KeyboardButton)) From {
                        {{New ReplyMarkups.KeyboardButton(DIALOGUE_BUTTON_YES), New ReplyMarkups.KeyboardButton(DIALOGUE_BUTTON_NO)}.ToList}}

    Private userkeyboard As New List(Of List(Of ReplyMarkups.KeyboardButton)) From {
                        {{New ReplyMarkups.KeyboardButton(DIALOGUE_BUTTON_USERDETAILS), New ReplyMarkups.KeyboardButton(DIALOGUE_BUTTON_BACK)}.ToList},
                        {{New ReplyMarkups.KeyboardButton(DIALOGUE_BUTTON_USERRESETPASSWORD), New ReplyMarkups.KeyboardButton(DIALOGUE_BUTTON_USERENABLEDISABLE), New ReplyMarkups.KeyboardButton(DIALOGUE_BUTTON_USERMEMBEROF)}.ToList}}

    Private searchgroupkeyboard As New List(Of List(Of ReplyMarkups.KeyboardButton)) From {
                        {{New ReplyMarkups.KeyboardButton(DIALOGUE_BUTTON_BACK)}.ToList}}

    Private _stage As DialogueStage = DialogueStage.SearchUser

    Public Property Stage As DialogueStage
        Get
            Return _stage
        End Get
        Set(value As DialogueStage)
            _stage = value
        End Set
    End Property

    Public Sub OnMessage(message As Message)

        If message.Text = "/start" Then

            Stage = DialogueStage.SearchUser
            SendRequestStageGreeting(message)

        ElseIf message.Text = DIALOGUE_BUTTON_USERDETAILS Then

            Stage = DialogueStage.UserMenu
            SendRequestStageUserDetails(message)

        ElseIf message.Text = DIALOGUE_BUTTON_BACK Then

            Select Case Stage
                Case DialogueStage.UserMenu
                    Stage = DialogueStage.SearchUser
                    SendRequestStageSearchUser(message)
                Case DialogueStage.SearchGroup
                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserMenu(message)
            End Select

        ElseIf message.Text = DIALOGUE_BUTTON_USERRESETPASSWORD Then

            If currentuser Is Nothing Then
                Stage = DialogueStage.SearchUser
                SendRequestStageSearchUser(message)
            End If

            Stage = DialogueStage.UserConfirmResetPassword
            SendRequestStageUserConfirmResetPassword(message)

        ElseIf message.Text = DIALOGUE_BUTTON_USERENABLEDISABLE Then

            If currentuser Is Nothing Then
                Stage = DialogueStage.SearchUser
                SendRequestStageSearchUser(message)
            End If

            Stage = DialogueStage.UserConfirmEnableDisable
            SendRequestStageUserConfirmEnableDisable(message)

        ElseIf message.Text = DIALOGUE_BUTTON_USERMEMBEROF Then

            If currentuser Is Nothing Then
                Stage = DialogueStage.SearchUser
                SendRequestStageSearchUser(message)
            End If

            Stage = DialogueStage.SearchGroup
            SendRequestStageSearchGroup(message)

        ElseIf message.Text = DIALOGUE_BUTTON_YES Then
            Select Case Stage
                Case DialogueStage.UserConfirmResetPassword

                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserResetPasswordCompleted(message)

                Case DialogueStage.UserConfirmEnableDisable

                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserEnableDisableCompleted(message)

                Case DialogueStage.GroupConfirmMemberOf

                    Stage = DialogueStage.UserMenu
                    SendRequestStageGroupMemberOfCompleted(message)

                Case Else

                    Stage = DialogueStage.SearchUser
                    SendRequestStageSearchUser(message)

            End Select

        ElseIf message.Text = DIALOGUE_BUTTON_NO Then

            Select Case Stage
                Case DialogueStage.UserConfirmResetPassword

                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserMenu(message)

                Case DialogueStage.UserConfirmEnableDisable

                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserMenu(message)

                Case DialogueStage.GroupConfirmMemberOf

                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserMenu(message)

                Case Else

                    Stage = DialogueStage.SearchUser
                    SendRequestStageSearchUser(message)

            End Select

        Else ' нажали не кнопку

            Select Case Stage
                Case DialogueStage.SearchUser

                    Dim responceguid As Guid = Nothing
                    Try
                        responceguid = New Guid(Decode58(message.Text.Replace("/", "")))
                    Catch
                    End Try
                    If Not responceguid = Nothing Then
                        Dim guidresults As New List(Of clsDirectoryObject)
                        For Each dmn In domains
                            guidresults.AddRange(searcher.SearchSync(New clsDirectoryObject(dmn.DefaultNamingContext, dmn), New clsFilter(responceguid.ToString, attributesForSearchDefault, New clsSearchObjectClasses(True, False, False, False, False))))
                        Next
                        Dim obj As clsDirectoryObject = If(guidresults.Count = 1, guidresults(0), Nothing)
                        If obj IsNot Nothing Then
                            If Not obj.SchemaClass = enmDirectoryObjectSchemaClass.User Then
                                SendRequestUnexpectedUser(message)
                                Exit Sub
                            End If

                            currentuser = obj

                            Stage = DialogueStage.UserMenu
                            SendRequestStageUserMenu(message)
                            Exit Sub
                        End If
                    End If

                    Dim results As New List(Of clsDirectoryObject)
                    For Each dmn In domains
                        results.AddRange(searcher.SearchSync(New clsDirectoryObject(dmn.DefaultNamingContext, dmn), New clsFilter(message.Text, attributesForSearchDefault, New clsSearchObjectClasses(True, False, False, False, False))))
                    Next

                    SendRequestStageSearchListObjects(message, results)

                Case DialogueStage.SearchGroup

                    If currentuser Is Nothing Then
                        Stage = DialogueStage.SearchUser
                        SendRequestStageSearchUser(message)
                    End If

                    Dim responceguid As Guid = Nothing
                    Try
                        responceguid = New Guid(Decode58(message.Text.Replace("/", "")))
                    Catch
                    End Try
                    If Not responceguid = Nothing Then
                        Dim guidresults As New List(Of clsDirectoryObject)
                        For Each dmn In domains
                            guidresults.AddRange(searcher.SearchSync(New clsDirectoryObject(dmn.DefaultNamingContext, dmn), New clsFilter(responceguid.ToString, attributesForSearchDefault, New clsSearchObjectClasses(False, False, False, True, False))))
                        Next
                        Dim obj As clsDirectoryObject = If(guidresults.Count = 1, guidresults(0), Nothing)
                        If obj IsNot Nothing Then
                            If Not obj.SchemaClass = enmDirectoryObjectSchemaClass.Group Then
                                SendRequestUnexpectedGroup(message)
                                Exit Sub
                            End If

                            currentgroup = obj

                            Stage = DialogueStage.GroupConfirmMemberOf
                            SendRequestStageGroupConfirmMemberOf(message)
                            Exit Sub
                        End If
                    End If

                    Dim results As New List(Of clsDirectoryObject)
                    For Each dmn In domains
                        results.AddRange(searcher.SearchSync(New clsDirectoryObject(dmn.DefaultNamingContext, dmn), New clsFilter("*" & message.Text & "*", attributesForSearchDefault, New clsSearchObjectClasses(False, False, False, True, False))))
                    Next

                    SendRequestStageSearchListObjects(message, results)

                Case Else

                    SendRequestUnexpected(message)

            End Select

        End If

    End Sub

    Private Async Sub SendRequestUnexpected(message As Message)
        Await Bot.SendTextMessageAsync(message.Chat.Id, "ВТФ??? нажми на кнопку!",)
    End Sub

    Private Async Sub SendRequestUnexpectedUser(message As Message)
        Await Bot.SendTextMessageAsync(message.Chat.Id, "Чойта??? Надо юзера выбрать!",)
    End Sub

    Private Async Sub SendRequestUnexpectedGroup(message As Message)
        Await Bot.SendTextMessageAsync(message.Chat.Id, "Нутычо??? Тут надо группу выбрать!",)
    End Sub

    Private Async Sub SendRequestStageGreeting(message As Message)
        Await Bot.SendTextMessageAsync(message.Chat.Id, String.Format(
        "Привет, {0}!" & nl &
        "Я бот ADTools." & nl &
        "Кого ищем?", message.Chat.Username))
    End Sub

    Private Async Sub SendRequestStageSearchUser(message As Message)
        Await Bot.SendTextMessageAsync(message.Chat.Id, "Кого ищем?")
    End Sub

    Private Async Sub SendRequestStageSearchListObjects(message As Message, objects As List(Of clsDirectoryObject))

        If objects.Count = 0 Then
            If Stage = DialogueStage.SearchUser Then
                Await Bot.SendTextMessageAsync(message.Chat.Id, "Никого не найдено")
            ElseIf Stage = DialogueStage.SearchGroup Then
                Await Bot.SendTextMessageAsync(message.Chat.Id, "Группа не найдена",,,,, New ReplyKeyboardMarkup(searchgroupkeyboard))
            End If

        ElseIf objects.Count > 50 Then

            Await Bot.SendTextMessageAsync(message.Chat.Id, "Найдено больше 50 объектов")

        Else

            Dim msg As String = ""
            For Each obj In objects
                msg &= If(obj.disabled = True, "⛔️ ", If(Stage = DialogueStage.SearchUser, "👤 ", "👥 ")) & obj.name & nl
                msg &= If(String.IsNullOrEmpty(obj.userPrincipalNameName), "", "📲 " & obj.userPrincipalNameName & nl)
                msg &= If(String.IsNullOrEmpty(obj.title), "", "📃 " & obj.title & nl)
                msg &= "/" & Encode58(obj.objectGUID.ToByteArray) & dnl
            Next
            Await Bot.SendTextMessageAsync(message.Chat.Id, msg)

        End If
    End Sub

    Private Async Sub SendRequestStageUserDetails(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = "Выбранный юзер:" & dnl
        msg &= If(currentuser.disabled = True, "⛔️ ", "👤 ") & currentuser.name & nl
        msg &= If(String.IsNullOrEmpty(currentuser.userPrincipalNameName), "", "📲 " & currentuser.userPrincipalNameName & nl)
        msg &= If(String.IsNullOrEmpty(currentuser.physicalDeliveryOfficeName), "", "🏢 " & currentuser.physicalDeliveryOfficeName & nl)
        msg &= If(String.IsNullOrEmpty(currentuser.telephoneNumber), "", "📞 " & currentuser.telephoneNumber & nl)
        msg &= If(String.IsNullOrEmpty(currentuser.mail), "", "✉️ " & currentuser.mail & nl)
        msg &= If(String.IsNullOrEmpty(currentuser.title), "", "📃 " & currentuser.title & ", " & currentuser.department & nl)
        msg &= If(String.IsNullOrEmpty(currentuser.passwordExpiresFormated), "", "🔑 " & currentuser.passwordExpiresFormated & nl)

        Await Bot.SendTextMessageAsync(message.Chat.Id, msg,,,,, New ReplyKeyboardMarkup(userkeyboard))
    End Sub


    Private Async Sub SendRequestStageUserMenu(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = "Юзер выбран:" & dnl
        msg &= If(currentuser.disabled = True, "⛔️ ", "👤 ") & currentuser.name & nl
        msg &= If(String.IsNullOrEmpty(currentuser.title), "", "📃 " & currentuser.title & nl)
        msg &= If(String.IsNullOrEmpty(currentuser.userPrincipalName), "", "📲 " & currentuser.userPrincipalName & nl)

        Await Bot.SendTextMessageAsync(message.Chat.Id, msg,,,,, New ReplyKeyboardMarkup(userkeyboard))
    End Sub

    Private Async Sub SendRequestStageUserConfirmResetPassword(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = String.Format("Сброс пароля:" & dnl & "👤 {0}" & dnl & "Чо серьезно?", currentuser.name)

        Await Bot.SendTextMessageAsync(message.Chat.Id, msg,,,,, New ReplyKeyboardMarkup(confimkeyboard))
    End Sub

    Private Async Sub SendRequestStageUserConfirmEnableDisable(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""

        If currentuser.disabled = True Then
            msg &= String.Format("Разблочить:" & dnl & "👤 {0}" & dnl & "А надо?", currentuser.name)
        Else
            msg &= String.Format("Заблочить:" & dnl & "👤 {0}" & dnl & "А надо?", currentuser.name)
        End If

        Await Bot.SendTextMessageAsync(message.Chat.Id, msg,,,,, New ReplyKeyboardMarkup(confimkeyboard))
    End Sub

    Private Async Sub SendRequestStageUserResetPasswordCompleted(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""

        Try
            currentuser.ResetPassword()
            currentuser.passwordNeverExpires = False
            msg &= String.Format("👤 {0}" & dnl & "Пароль сброшен", currentuser.name)
        Catch ex As Exception
            msg = String.Format("Не получилось сбросить пароль:" & dnl & "👤 {0}" & dnl & "{1}", currentuser.name, ex.Message)
        End Try

        Await Bot.SendTextMessageAsync(message.Chat.Id, msg,,,,, New ReplyKeyboardMarkup(userkeyboard))
    End Sub

    Private Async Sub SendRequestStageUserEnableDisableCompleted(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""

        Try
            If currentuser.disabled Then
                currentuser.disabled = False
                msg &= String.Format("👤 {0}" & dnl & "разблокирован", currentuser.name)
            Else
                currentuser.disabled = True
                msg &= String.Format("👤 {0}" & dnl & "заблокирован", currentuser.name)
            End If

        Catch ex As Exception
            msg = String.Format("Не получилось заблочить/разблочить:" & dnl & "👤 {0}" & dnl & "{1}", currentuser.name, ex.Message)
        End Try

        Await Bot.SendTextMessageAsync(message.Chat.Id, msg,,,,, New ReplyKeyboardMarkup(userkeyboard))
    End Sub

    Private Async Sub SendRequestStageSearchGroup(message As Message)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""

        If currentuser.memberOf.Count > 0 Then
            msg &= "Текущие группы:" & dnl
            For Each group As clsDirectoryObject In currentuser.memberOf
                msg &= "👥 " & group.name & nl
                msg &= "/" & Encode58(group.objectGUID.ToByteArray) & dnl
            Next
        End If

        msg &= "Введи кусок названия группы"

        Await Bot.SendTextMessageAsync(message.Chat.Id, msg,,,,, New ReplyKeyboardMarkup(searchgroupkeyboard))
    End Sub

    Private Async Sub SendRequestStageGroupConfirmMemberOf(message As Message)
        If currentuser Is Nothing Or currentgroup Is Nothing Then Exit Sub

        Dim msg As String = ""

        Dim newgroup As Boolean = True
        For Each group As clsDirectoryObject In currentuser.memberOf
            If group.name = currentgroup.name Then newgroup = False
        Next

        If newgroup Then
            msg &= String.Format("Добавить:" & dnl & "👤 {0}" & dnl & "в группу" & dnl & "👥 {1}" & dnl & "ммм?", currentuser.name, currentgroup.name)
        Else
            msg &= String.Format("Удалить:" & dnl & "👤 {0}" & dnl & "из группы" & dnl & "👥 {1}" & dnl & "ммм?", currentuser.name, currentgroup.name)
        End If

        Await Bot.SendTextMessageAsync(message.Chat.Id, msg,,,,, New ReplyKeyboardMarkup(confimkeyboard))
    End Sub

    Private Async Sub SendRequestStageGroupMemberOfCompleted(message As Message)
        If currentuser Is Nothing Or currentgroup Is Nothing Then Exit Sub

        Dim msg As String = ""

        Try
            Dim newgroup As Boolean = True
            For Each group As clsDirectoryObject In currentuser.memberOf
                If group.name = currentgroup.name Then newgroup = False
            Next

            If newgroup Then
                currentgroup.UpdateAttribute(DirectoryServices.Protocols.DirectoryAttributeOperation.Add, "member", currentuser.distinguishedName)
                currentuser.memberOf.Add(currentgroup)
                msg &= String.Format("👤 {0}" & dnl & "добавлен в группу" & dnl & "👥 {1}", currentuser.name, currentgroup.name)
            Else
                currentgroup.UpdateAttribute(DirectoryServices.Protocols.DirectoryAttributeOperation.Delete, "member", currentuser.distinguishedName)
                currentuser.memberOf.Remove(currentgroup)
                For Each group As clsDirectoryObject In currentuser.memberOf
                    If group.name = currentgroup.name Then currentuser.memberOf.Remove(group) : Exit For
                Next
                msg &= String.Format("👤 {0}" & dnl & "удален из группы" & dnl & "👥 {1}", currentuser.name, currentgroup.name)
            End If

        Catch ex As Exception
            msg = String.Format("Не получилось добавить/удалить:" & dnl & "👤 {0}" & dnl & "в/из группы" & dnl & "👥 {1}" & dnl & "{2}", currentuser.name, currentgroup.name, ex.Message)
        End Try

        Await Bot.SendTextMessageAsync(message.Chat.Id, msg,,,,, New ReplyKeyboardMarkup(userkeyboard))
    End Sub

End Module
