
Namespace My

    'このクラスでは設定クラスでの特定のイベントを処理することができます:
    ' SettingChanging イベントは、設定値が変更される前に発生します。
    ' PropertyChanged イベントは、設定値が変更された後に発生します。
    ' SettingsLoaded イベントは、設定値が読み込まれた後に発生します。
    ' SettingsSaving イベントは、設定値が保存される前に発生します。
    Partial Friend NotInheritable Class MySettings
        <Global.System.Configuration.UserScopedSettingAttribute()>
        Public Property titleList() As System.Collections.Generic.List(Of String)
            Get
                Return CType(Me("titleList"), System.Collections.Generic.List(Of String))
            End Get
            Set(ByVal value As System.Collections.Generic.List(Of String))
                Me("titleList") = value
            End Set
        End Property

        <Global.System.Configuration.UserScopedSettingAttribute()>
        Public Property idNameList() As System.Collections.Generic.List(Of String)
            Get
                Return CType(Me("idNameList"), System.Collections.Generic.List(Of String))
            End Get
            Set(ByVal value As System.Collections.Generic.List(Of String))
                Me("idNameList") = value
            End Set
        End Property

        <Global.System.Configuration.UserScopedSettingAttribute()>
        Public Property amountList() As System.Collections.Generic.List(Of String)
            Get
                Return CType(Me("amountList"), System.Collections.Generic.List(Of String))
            End Get
            Set(ByVal value As System.Collections.Generic.List(Of String))
                Me("amountList") = value
            End Set
        End Property

        <Global.System.Configuration.UserScopedSettingAttribute()>
        Public Property maisuList() As System.Collections.Generic.List(Of String)
            Get
                Return CType(Me("maisuList"), System.Collections.Generic.List(Of String))
            End Get
            Set(ByVal value As System.Collections.Generic.List(Of String))
                Me("maisuList") = value
            End Set
        End Property
    End Class
End Namespace
