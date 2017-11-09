Imports TeleBotDotNet.Requests.Methods
Imports TeleBotDotNet.Responses.Methods

Module mdlDialogue

    Private currentuser As clsDirectoryObject
    Private currentgroup As clsDirectoryObject

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

    Private confimkeyboard As New List(Of List(Of String)) From {
                        {{DIALOGUE_BUTTON_YES, DIALOGUE_BUTTON_NO}.ToList}}

    Private userkeyboard As New List(Of List(Of String)) From {
                        {{DIALOGUE_BUTTON_USERDETAILS, DIALOGUE_BUTTON_BACK}.ToList},
                        {{DIALOGUE_BUTTON_USERRESETPASSWORD, DIALOGUE_BUTTON_USERENABLEDISABLE, DIALOGUE_BUTTON_USERMEMBEROF}.ToList}}

    Private searchgroupkeyboard As New List(Of List(Of String)) From {
                        {{DIALOGUE_BUTTON_BACK}.ToList}}

    Private _stage As DialogueStage = DialogueStage.SearchUser

    Public Property Stage As DialogueStage
        Get
            Return _stage
        End Get
        Set(value As DialogueStage)
            _stage = value
        End Set
    End Property

    Public Sub ProcessDialogue(responce As TeleBotDotNet.Responses.Types.UpdateResponse)

        If responce.Message.Text = "/start" Then

            Stage = DialogueStage.SearchUser
            SendRequestStageGreeting(responce)

        ElseIf responce.Message.Text = DIALOGUE_BUTTON_USERDETAILS Then

            Stage = DialogueStage.UserMenu
            SendRequestStageUserDetails(responce)

        ElseIf responce.Message.Text = DIALOGUE_BUTTON_BACK Then

            Select Case Stage
                Case DialogueStage.UserMenu
                    Stage = DialogueStage.SearchUser
                    SendRequestStageSearchUser(responce)
                Case DialogueStage.SearchGroup
                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserMenu(responce)
            End Select

        ElseIf responce.Message.Text = DIALOGUE_BUTTON_USERRESETPASSWORD Then

            If currentuser Is Nothing Then
                Stage = DialogueStage.SearchUser
                SendRequestStageSearchUser(responce)
            End If

            Stage = DialogueStage.UserConfirmResetPassword
            SendRequestStageUserConfirmResetPassword(responce)

        ElseIf responce.Message.Text = DIALOGUE_BUTTON_USERENABLEDISABLE Then

            If currentuser Is Nothing Then
                Stage = DialogueStage.SearchUser
                SendRequestStageSearchUser(responce)
            End If

            Stage = DialogueStage.UserConfirmEnableDisable
            SendRequestStageUserConfirmEnableDisable(responce)

        ElseIf responce.Message.Text = DIALOGUE_BUTTON_USERMEMBEROF Then

            If currentuser Is Nothing Then
                Stage = DialogueStage.SearchUser
                SendRequestStageSearchUser(responce)
            End If

            Stage = DialogueStage.SearchGroup
            SendRequestStageSearchGroup(responce)

        ElseIf responce.Message.Text = DIALOGUE_BUTTON_YES Then
            Select Case Stage
                Case DialogueStage.UserConfirmResetPassword

                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserResetPasswordCompleted(responce)

                Case DialogueStage.UserConfirmEnableDisable

                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserEnableDisableCompleted(responce)

                Case DialogueStage.GroupConfirmMemberOf

                    Stage = DialogueStage.UserMenu
                    SendRequestStageGroupMemberOfCompleted(responce)

                Case Else

                    Stage = DialogueStage.SearchUser
                    SendRequestStageSearchUser(responce)

            End Select

        ElseIf responce.Message.Text = DIALOGUE_BUTTON_NO Then

            Select Case Stage
                Case DialogueStage.UserConfirmResetPassword

                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserMenu(responce)

                Case DialogueStage.UserConfirmEnableDisable

                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserMenu(responce)

                Case DialogueStage.GroupConfirmMemberOf

                    Stage = DialogueStage.UserMenu
                    SendRequestStageUserMenu(responce)

                Case Else

                    Stage = DialogueStage.SearchUser
                    SendRequestStageSearchUser(responce)

            End Select

        Else ' нажали не кнопку

            Select Case Stage
                Case DialogueStage.SearchUser

                    Dim responceguid As Guid = Nothing
                    Try
                        responceguid = New Guid(Decode58(responce.Message.Text.Replace("/", "")))
                    Catch
                    End Try
                    If Not responceguid = Nothing Then
                        Dim obj As clsDirectoryObject = SearchGUID(responceguid)
                        If obj IsNot Nothing Then
                            If Not obj.SchemaClass = clsDirectoryObject.enmSchemaClass.User Then
                                SendRequestUnexpectedUser(responce)
                                Exit Sub
                            End If

                            currentuser = obj

                            Stage = DialogueStage.UserMenu
                            SendRequestStageUserMenu(responce)
                            Exit Sub
                        End If
                    End If

                    SendRequestStageSearchListObjects(responce, Search(New clsFilter(responce.Message.Text, attributesForSearchDefault, New clsSearchObjectClasses(True, False, False, False, False), False), Nothing))

                Case DialogueStage.SearchGroup

                    If currentuser Is Nothing Then
                        Stage = DialogueStage.SearchUser
                        SendRequestStageSearchUser(responce)
                    End If

                    Dim responceguid As Guid = Nothing
                    Try
                        responceguid = New Guid(Decode58(responce.Message.Text.Replace("/", "")))
                    Catch
                    End Try
                    If Not responceguid = Nothing Then
                        Dim obj As clsDirectoryObject = SearchGUID(responceguid)
                        If obj IsNot Nothing Then
                            If Not obj.SchemaClass = clsDirectoryObject.enmSchemaClass.Group Then
                                SendRequestUnexpectedGroup(responce)
                                Exit Sub
                            End If

                            currentgroup = obj

                            Stage = DialogueStage.GroupConfirmMemberOf
                            SendRequestStageGroupConfirmMemberOf(responce)
                            Exit Sub
                        End If
                    End If

                    SendRequestStageSearchListObjects(responce, Search(New clsFilter(responce.Message.Text, attributesForSearchDefault, New clsSearchObjectClasses(False, False, False, True, False), True), currentuser.Domain))

                Case Else

                    SendRequestUnexpected(responce)

            End Select

        End If

    End Sub

    Private Sub SendRequestUnexpected(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        SendTelegramMessage(responce.Message.From.Id, "ВТФ??? нажми на кнопку!",)
    End Sub

    Private Sub SendRequestUnexpectedUser(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        SendTelegramMessage(responce.Message.From.Id, "Чойта??? Надо юзера выбрать!",)
    End Sub

    Private Sub SendRequestUnexpectedGroup(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        SendTelegramMessage(responce.Message.From.Id, "Нутычо??? Тут надо группу выбрать!",)
    End Sub

    Private Sub SendRequestStageGreeting(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        SendTelegramMessage(responce.Message.From.Id, String.Format(
        "Привет, {0}!" & vbCrLf &
        "Я бот ADTools." & vbCrLf &
        "Кого ищем?", responce.Message.From.UserName), , True)
    End Sub

    Private Sub SendRequestStageSearchUser(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        SendTelegramMessage(responce.Message.From.Id, "Кого ищем?", , True)
    End Sub

    Private Sub SendRequestStageSearchListObjects(responce As TeleBotDotNet.Responses.Types.UpdateResponse, objects As List(Of clsDirectoryObject))
        Dim msg As String = ""

        For Each obj In objects
            msg &= If(obj.disabled = True, "⛔️ ", If(Stage = DialogueStage.SearchUser, "👤 ", "👥 ")) & obj.name & vbCrLf
            msg &= If(String.IsNullOrEmpty(obj.userPrincipalNameName), "", "📲 " & obj.userPrincipalNameName & vbCrLf)
            msg &= If(String.IsNullOrEmpty(obj.title), "", "📃 " & obj.title & vbCrLf)
            msg &= "/" & Encode58(obj.objectGUID.ToByteArray) & vbCrLf & vbCrLf
        Next

        If Stage = DialogueStage.SearchUser Then
            If objects.Count = 0 Then msg = "Никого не найдено"
            SendTelegramMessage(responce.Message.From.Id, msg)
        ElseIf Stage = DialogueStage.SearchGroup Then
            If objects.Count = 0 Then msg = "Группа не найдена"
            SendTelegramMessage(responce.Message.From.Id, msg, searchgroupkeyboard)
        Else

        End If
    End Sub

    Private Sub SendRequestStageUserDetails(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = "Выбранный юзер:" & vbCrLf & vbCrLf
        msg &= If(currentuser.disabled = True, "⛔️ ", "👤 ") & currentuser.name & vbCrLf
        msg &= If(String.IsNullOrEmpty(currentuser.userPrincipalNameName), "", "📲 " & currentuser.userPrincipalNameName & vbCrLf)
        msg &= If(String.IsNullOrEmpty(currentuser.physicalDeliveryOfficeName), "", "🏢 " & currentuser.physicalDeliveryOfficeName & vbCrLf)
        msg &= If(String.IsNullOrEmpty(currentuser.telephoneNumber), "", "📞 " & currentuser.telephoneNumber & vbCrLf)
        msg &= If(String.IsNullOrEmpty(currentuser.mail), "", "✉️ " & currentuser.mail & vbCrLf)
        msg &= If(String.IsNullOrEmpty(currentuser.title), "", "📃 " & currentuser.title & vbCrLf)
        msg &= If(String.IsNullOrEmpty(currentuser.passwordExpiresFormated), "", "🔑 " & currentuser.passwordExpiresFormated & vbCrLf)

        SendTelegramMessage(responce.Message.From.Id, msg, userkeyboard)
    End Sub


    Private Sub SendRequestStageUserMenu(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = "Юзер выбран:" & vbCrLf & vbCrLf
        msg &= If(currentuser.disabled = True, "⛔️ ", "👤 ") & currentuser.name & vbCrLf
        msg &= If(String.IsNullOrEmpty(currentuser.title), "", "📃 " & currentuser.title & vbCrLf)
        msg &= If(String.IsNullOrEmpty(currentuser.userPrincipalName), "", "🔑 " & currentuser.userPrincipalName & vbCrLf)

        SendTelegramMessage(responce.Message.From.Id, msg, userkeyboard)
    End Sub

    Private Sub SendRequestStageUserConfirmResetPassword(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = String.Format("Сброс пароля:" & vbCrLf & vbCrLf & "👤 {0}" & vbCrLf & vbCrLf & "Чо серьезно?", currentuser.name)

        SendTelegramMessage(responce.Message.From.Id, msg, confimkeyboard)
    End Sub

    Private Sub SendRequestStageUserConfirmEnableDisable(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""

        If currentuser.disabled = True Then
            msg &= String.Format("Разблочить:" & vbCrLf & vbCrLf & "👤 {0}" & vbCrLf & vbCrLf & "А надо?", currentuser.name)
        Else
            msg &= String.Format("Заблочить:" & vbCrLf & vbCrLf & "👤 {0}" & vbCrLf & vbCrLf & "А надо?", currentuser.name)
        End If

        SendTelegramMessage(responce.Message.From.Id, msg, confimkeyboard)
    End Sub

    Private Sub SendRequestStageUserResetPasswordCompleted(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""

        Try
            currentuser.ResetPassword()
            currentuser.passwordNeverExpires = False
            msg &= String.Format("👤 {0}" & vbCrLf & vbCrLf & "Пароль сброшен", currentuser.name)
        Catch ex As Exception
            msg = String.Format("Не получилось сбросить пароль:" & vbCrLf & vbCrLf & "👤 {0}" & vbCrLf & vbCrLf & "{1}", currentuser.name, ex.Message)
        End Try

        SendTelegramMessage(responce.Message.From.Id, msg, userkeyboard)
    End Sub

    Private Sub SendRequestStageUserEnableDisableCompleted(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""

        Try
            If currentuser.disabled Then
                currentuser.disabled = False
                msg &= String.Format("👤 {0}" & vbCrLf & vbCrLf & "разблокирован", currentuser.name)
            Else
                currentuser.disabled = True
                msg &= String.Format("👤 {0}" & vbCrLf & vbCrLf & "заблокирован", currentuser.name)
            End If

        Catch ex As Exception
            msg = String.Format("Не получилось заблочить/разблочить:" & vbCrLf & vbCrLf & "👤 {0}" & vbCrLf & vbCrLf & "{1}", currentuser.name, ex.Message)
        End Try

        SendTelegramMessage(responce.Message.From.Id, msg, userkeyboard)
    End Sub

    Private Sub SendRequestStageSearchGroup(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        If currentuser Is Nothing Then Exit Sub

        Dim msg As String = ""

        If currentuser.memberOf.Count > 0 Then
            msg &= "Текущие группы:" & vbCrLf & vbCrLf
            For Each group As clsDirectoryObject In currentuser.memberOf
                msg &= "👥 " & group.name & vbCrLf
                msg &= "/" & Encode58(group.objectGUID.ToByteArray) & vbCrLf & vbCrLf
            Next
        End If

        msg &= "Введи кусок названия группы"

        SendTelegramMessage(responce.Message.From.Id, msg, searchgroupkeyboard)
    End Sub

    Private Sub SendRequestStageGroupConfirmMemberOf(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        If currentuser Is Nothing Or currentgroup Is Nothing Then Exit Sub

        Dim msg As String = ""

        Dim newgroup As Boolean = True
        For Each group As clsDirectoryObject In currentuser.memberOf
            If group.name = currentgroup.name Then newgroup = False
        Next

        If newgroup Then
            msg &= String.Format("Добавить:" & vbCrLf & vbCrLf & "👤 {0}" & vbCrLf & vbCrLf & "в группу" & vbCrLf & vbCrLf & "👥 {1}" & vbCrLf & vbCrLf & "ммм?", currentuser.name, currentgroup.name)
        Else
            msg &= String.Format("Удалить:" & vbCrLf & vbCrLf & "👤 {0}" & vbCrLf & vbCrLf & "из группы" & vbCrLf & vbCrLf & "👥 {1}" & vbCrLf & vbCrLf & "ммм?", currentuser.name, currentgroup.name)
        End If

        SendTelegramMessage(responce.Message.From.Id, msg, confimkeyboard)
    End Sub

    Private Sub SendRequestStageGroupMemberOfCompleted(responce As TeleBotDotNet.Responses.Types.UpdateResponse)
        If currentuser Is Nothing Or currentgroup Is Nothing Then Exit Sub

        Dim msg As String = ""

        Try
            Dim newgroup As Boolean = True
            For Each group As clsDirectoryObject In currentuser.memberOf
                If group.name = currentgroup.name Then newgroup = False
            Next

            If newgroup Then
                currentgroup.Entry.Invoke("Add", currentuser.distinguishedNameFull)
                currentgroup.Entry.CommitChanges()
                currentuser.memberOf.Add(currentgroup)
                msg &= String.Format("👤 {0}" & vbCrLf & vbCrLf & "добавлен в группу" & vbCrLf & vbCrLf & "👥 {1}", currentuser.name, currentgroup.name)
            Else
                currentgroup.Entry.Invoke("Remove", currentuser.distinguishedNameFull)
                currentgroup.Entry.CommitChanges()

                For Each group As clsDirectoryObject In currentuser.memberOf
                    If group.name = currentgroup.name Then currentuser.memberOf.Remove(group) : Exit For
                Next
                msg &= String.Format("👤 {0}" & vbCrLf & vbCrLf & "удален из группы" & vbCrLf & vbCrLf & "👥 {1}", currentuser.name, currentgroup.name)
            End If

        Catch ex As Exception
            msg = String.Format("Не получилось добавить/удалить:" & vbCrLf & vbCrLf & "👤 {0}" & vbCrLf & vbCrLf & "в/из группы" & vbCrLf & vbCrLf & "👥 {1}" & vbCrLf & vbCrLf & "{2}", currentuser.name, currentgroup.name, ex.Message)
        End Try

        SendTelegramMessage(responce.Message.From.Id, msg, userkeyboard)
    End Sub

End Module
