'次点：クレカの検証
'500円未満はクレカの義務付けチェック

Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Text.Json
Imports System.Text.Encodings.Web
Imports System.Text.Unicode
Imports System.IO

Public Class Form1

    'チケットデータ保持用
    Public titleList = New List(Of String)
    Public idNameList = New List(Of String)
    Public maisuList = New List(Of String)

    'パスワード表示フラグ
    Public isPasswordVisible = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        imtURL.Text = My.Settings.URL
        If My.Settings.StartDateTime <> "" Then
            imdStartDateTime.Value = My.Settings.StartDateTime
        End If
        cmbStore.SelectedValue = My.Settings.Store
        If My.Settings.MaxUser <> "" Then
            imnUserE.Value = My.Settings.MaxUser
        End If


        Dim titleListw = My.Settings.titleList
        Dim idNameListw = My.Settings.idNameList
        Dim maisuListw = My.Settings.MaisuList

        setDispData(titleListw, idNameListw, maisuListw)

        imtUser1.Text = "dabame8549@mi166.com"
        imtUser2.Text = "natipo5714@mom2kid.com"
        imtUser3.Text = "kumyu449@instaddr.win"
        imtUser4.Text = "mipyagyo@instaddr.win"

        imtPassword1.Text = "6yjAWLhCxVT2"
        imtPassword2.Text = "6yjAWLhCxVT2"
        imtPassword3.Text = "6yjAWLhCxVT2"
        imtPassword4.Text = "6yjAWLhCxVT2"


    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        saveDispData()
    End Sub



    Private Sub imtURL_TextChanged(sender As Object, e As EventArgs) Handles imtURL.TextChanged

    End Sub

    Private Sub btnGetDisp_Click(sender As Object, e As EventArgs) Handles btnGetDisp.Click
        Dim sURL As String
        Dim sHTML As String
        Dim objWC As System.Net.WebClient
        Dim objDOC As HtmlAgilityPack.HtmlDocument
        Dim titleNodes As HtmlAgilityPack.HtmlNodeCollection
        Dim idNodes As HtmlAgilityPack.HtmlNodeCollection
        Dim timeNodes As HtmlAgilityPack.HtmlNodeCollection

        sURL = imtURL.Text
        objWC = New System.Net.WebClient()
        objWC.Encoding = System.Text.Encoding.UTF8
        sHTML = objWC.DownloadString(sURL)

        objDOC = New HtmlAgilityPack.HtmlDocument()
        objDOC.LoadHtml(sHTML)

        'HTMLソースは「<span class=bld>110.5740 JPY</span>」となっています。

        'タイトル取得
        titleNodes = objDOC.DocumentNode.SelectNodes("//*[@id='purchase_form']/section[1]/section/section[*]/section/section[1]/h4")
        'id取得
        idNodes = objDOC.DocumentNode.SelectNodes("//*[@id='purchase_form']/section[1]/section/section[*]/section/section[2]/div")

        titleList = New List(Of String)
        idNameList = New List(Of String)
        maisuList = New List(Of String)

        For Each node As HtmlAgilityPack.HtmlNode In titleNodes
            titleList.Add(node.InnerText)
        Next
        For Each node As HtmlAgilityPack.HtmlNode In idNodes
            idNameList.Add(node.Id.Split("_")(3))
        Next
        For index = 1 To 10
            maisuList.Add(0)
        Next

        '発売時間取得
        timeNodes = objDOC.DocumentNode.SelectNodes("//*[@id='purchase_form']/section[1]/section/section[1]/p/span[1]")

        For Each node As HtmlAgilityPack.HtmlNode In timeNodes
            Dim startTimeText = node.InnerText
            Dim reg1 As New Regex("\(.+?\)") '()を排除

            startTimeText = reg1.Replace(startTimeText, ",").Replace(" ", "")

            Dim startDate = startTimeText.Split(",")(0)
            Dim startTime = startTimeText.Split(",")(1)

            imdStartDateTime.Value = startDate + " " + startTime

        Next


        setDispData(titleList, idNameList, maisuList)

    End Sub

    Private Sub setDispData(titleList As List(Of String), idNameList As List(Of String), maisuList As List(Of String))

        imtTicketName1.Text = ""
        imtTicketName2.Text = ""
        imtTicketName3.Text = ""
        imtTicketName4.Text = ""
        imtTicketName5.Text = ""
        imtTicketName6.Text = ""
        imtTicketName7.Text = ""
        imtTicketName8.Text = ""
        imtTicketName9.Text = ""
        imtTicketName10.Text = ""

        imtTicketId1.Text = ""
        imtTicketId2.Text = ""
        imtTicketId3.Text = ""
        imtTicketId4.Text = ""
        imtTicketId5.Text = ""
        imtTicketId6.Text = ""
        imtTicketId7.Text = ""
        imtTicketId8.Text = ""
        imtTicketId9.Text = ""
        imtTicketId10.Text = ""

        imnMaisu1.Text = ""
        imnMaisu2.Text = ""
        imnMaisu3.Text = ""
        imnMaisu4.Text = ""
        imnMaisu5.Text = ""
        imnMaisu6.Text = ""
        imnMaisu7.Text = ""
        imnMaisu8.Text = ""
        imnMaisu9.Text = ""
        imnMaisu10.Text = ""

        For index = 1 To titleList.Count()
            Select Case index
                Case 1 : imtTicketName1.Text = titleList(index - 1)
                Case 2 : imtTicketName2.Text = titleList(index - 1)
                Case 3 : imtTicketName3.Text = titleList(index - 1)
                Case 4 : imtTicketName4.Text = titleList(index - 1)
                Case 5 : imtTicketName5.Text = titleList(index - 1)
                Case 6 : imtTicketName6.Text = titleList(index - 1)
                Case 7 : imtTicketName7.Text = titleList(index - 1)
                Case 8 : imtTicketName8.Text = titleList(index - 1)
                Case 9 : imtTicketName9.Text = titleList(index - 1)
                Case 10 : imtTicketName10.Text = titleList(index - 1)
            End Select
        Next

        For index = 1 To idNameList.Count()
            Select Case index
                Case 1 : imtTicketId1.Text = idNameList(index - 1)
                Case 2 : imtTicketId2.Text = idNameList(index - 1)
                Case 3 : imtTicketId3.Text = idNameList(index - 1)
                Case 4 : imtTicketId4.Text = idNameList(index - 1)
                Case 5 : imtTicketId5.Text = idNameList(index - 1)
                Case 6 : imtTicketId6.Text = idNameList(index - 1)
                Case 7 : imtTicketId7.Text = idNameList(index - 1)
                Case 8 : imtTicketId8.Text = idNameList(index - 1)
                Case 9 : imtTicketId9.Text = idNameList(index - 1)
                Case 10 : imtTicketId10.Text = idNameList(index - 1)
            End Select
        Next

        For index = 1 To maisuList.Count()
            Select Case index
                Case 1 : imnMaisu1.Text = maisuList(index - 1)
                Case 2 : imnMaisu2.Text = maisuList(index - 1)
                Case 3 : imnMaisu3.Text = maisuList(index - 1)
                Case 4 : imnMaisu4.Text = maisuList(index - 1)
                Case 5 : imnMaisu5.Text = maisuList(index - 1)
                Case 6 : imnMaisu6.Text = maisuList(index - 1)
                Case 7 : imnMaisu7.Text = maisuList(index - 1)
                Case 8 : imnMaisu8.Text = maisuList(index - 1)
                Case 9 : imnMaisu9.Text = maisuList(index - 1)
                Case 10 : imnMaisu10.Text = maisuList(index - 1)
            End Select
        Next

    End Sub

    Private Sub saveDispData()

        Dim titleListw = New List(Of String)
        Dim idNameListw = New List(Of String)
        Dim maisuListw = New List(Of String)

        For index = 1 To 15
            titleListw.Add("")
            Select Case index
                Case 1 : titleListw(index - 1) = imtTicketName1.Text
                Case 2 : titleListw(index - 1) = imtTicketName2.Text
                Case 3 : titleListw(index - 1) = imtTicketName3.Text
                Case 4 : titleListw(index - 1) = imtTicketName4.Text
                Case 5 : titleListw(index - 1) = imtTicketName5.Text
                Case 6 : titleListw(index - 1) = imtTicketName6.Text
                Case 7 : titleListw(index - 1) = imtTicketName7.Text
                Case 8 : titleListw(index - 1) = imtTicketName8.Text
                Case 9 : titleListw(index - 1) = imtTicketName9.Text
                Case 10 : titleListw(index - 1) = imtTicketName10.Text
                Case 11 : titleListw(index - 1) = imtTicketName11.Text
                Case 12 : titleListw(index - 1) = imtTicketName12.Text
                Case 13 : titleListw(index - 1) = imtTicketName13.Text
                Case 14 : titleListw(index - 1) = imtTicketName14.Text
                Case 15 : titleListw(index - 1) = imtTicketName15.Text
            End Select
        Next

        For index = 1 To 15
            idNameListw.Add("")
            Select Case index
                Case 1 : idNameListw(index - 1) = imtTicketId1.Text
                Case 2 : idNameListw(index - 1) = imtTicketId2.Text
                Case 3 : idNameListw(index - 1) = imtTicketId3.Text
                Case 4 : idNameListw(index - 1) = imtTicketId4.Text
                Case 5 : idNameListw(index - 1) = imtTicketId5.Text
                Case 6 : idNameListw(index - 1) = imtTicketId6.Text
                Case 7 : idNameListw(index - 1) = imtTicketId7.Text
                Case 8 : idNameListw(index - 1) = imtTicketId8.Text
                Case 9 : idNameListw(index - 1) = imtTicketId9.Text
                Case 10 : idNameListw(index - 1) = imtTicketId10.Text
                Case 11 : idNameListw(index - 1) = imtTicketId11.Text
                Case 12 : idNameListw(index - 1) = imtTicketId12.Text
                Case 13 : idNameListw(index - 1) = imtTicketId13.Text
                Case 14 : idNameListw(index - 1) = imtTicketId14.Text
                Case 15 : idNameListw(index - 1) = imtTicketId15.Text
            End Select
        Next

        For index = 1 To 15
            maisuListw.Add("")
            Select Case index
                Case 1 : maisuListw(index - 1) = imnMaisu1.Text
                Case 2 : maisuListw(index - 1) = imnMaisu2.Text
                Case 3 : maisuListw(index - 1) = imnMaisu3.Text
                Case 4 : maisuListw(index - 1) = imnMaisu4.Text
                Case 5 : maisuListw(index - 1) = imnMaisu5.Text
                Case 6 : maisuListw(index - 1) = imnMaisu6.Text
                Case 7 : maisuListw(index - 1) = imnMaisu7.Text
                Case 8 : maisuListw(index - 1) = imnMaisu8.Text
                Case 9 : maisuListw(index - 1) = imnMaisu9.Text
                Case 10 : maisuListw(index - 1) = imnMaisu10.Text
                Case 11 : maisuListw(index - 1) = imnMaisu11.Text
                Case 12 : maisuListw(index - 1) = imnMaisu12.Text
                Case 13 : maisuListw(index - 1) = imnMaisu13.Text
                Case 14 : maisuListw(index - 1) = imnMaisu14.Text
                Case 15 : maisuListw(index - 1) = imnMaisu15.Text
            End Select
        Next

        My.Settings.URL = imtURL.Text
        My.Settings.StartDateTime = imdStartDateTime.Value
        My.Settings.Store = cmbStore.SelectedValue
        My.Settings.MaxUser = imnUserE.Value

        My.Settings.titleList = titleListw
        My.Settings.idNameList = idNameListw
        My.Settings.MaisuList = maisuListw


    End Sub

    Private Sub btnDispPassword_Click(sender As Object, e As EventArgs) Handles btnDIspPassword.Click
        If isPasswordVisible = True Then
            imtPassword1.PasswordChar = "*"
            imtPassword2.PasswordChar = "*"
            imtPassword3.PasswordChar = "*"
            imtPassword4.PasswordChar = "*"
            imtPassword5.PasswordChar = "*"
            isPasswordVisible = False
        Else
            imtPassword1.PasswordChar = ControlChars.NullChar
            imtPassword2.PasswordChar = ControlChars.NullChar
            imtPassword3.PasswordChar = ControlChars.NullChar
            imtPassword4.PasswordChar = ControlChars.NullChar
            imtPassword5.PasswordChar = ControlChars.NullChar
            isPasswordVisible = True
        End If
    End Sub

    Private Sub btnMakeJson_Click(sender As Object, e As EventArgs) Handles btnMakeJson.Click
        writeJsonFile()
    End Sub

    Private Sub writeJsonFile()
        Dim enc As Encoding = Encoding.UTF8
        Dim configCommonStr As String = ""
        Dim configCommonFilePath As String = "C:\tool\yt-win-scraping\lpkickrun-common-template.json"
        Dim configStr As String = ""
        Dim configFilePath As String = "C:\tool\yt-win-scraping\lpkickrun-template.json"

        'ファイルからJson文字列を読み込む
        Using sr As New System.IO.StreamReader(configCommonFilePath, enc)
            configCommonStr = sr.ReadToEnd()
        End Using
        Using sr As New System.IO.StreamReader(configFilePath, enc)
            configStr = sr.ReadToEnd()
        End Using

        'Json文字列をJson形式データに復元する
        Dim configCommonObj As ConfigCommonJsonClass = JsonSerializer.Deserialize(Of ConfigCommonJsonClass)(configCommonStr)
        Dim configObj As ConfigJsonClass = JsonSerializer.Deserialize(Of ConfigJsonClass)(configStr)

        '【共通ファイル用】
        configObj.isTest = chkTest.Checked

        '【指令ファイル用】

        '支払先
        If cmbStore.SelectedValue = "クレジットカード" Then
            configObj.isCredit = True
        Else
            configObj.isCredit = False
            If cmbStore.SelectedValue = "ファミリーマート" Then
                configObj.storeTypeId = "016"
            End If
            If cmbStore.SelectedValue = "ローソン" Then
                configObj.storeTypeId = "002"
            End If
        End If

        '時間
        Dim startTime As String = imdStartDateTime.Value

        Dim dt1 = Convert.ToDateTime(startTime)
        dt1 = dt1.AddMinutes(-15)
        Dim dt2 = Convert.ToDateTime(startTime)
        dt2 = dt2.AddHours(1)

        Dim loginTime = dt1.ToString().Replace("/", "-")
        Dim endTime = dt2.ToString().Replace("/", "-")

        configObj.startTime = startTime.Replace("/", "-")
        configObj.loginTime = loginTime.Replace("/", "-")
        configObj.endTime = endTime.Replace("/", "-")

        Dim userListw = New List(Of String)

        If imtUser1.Text <> "" And imnUserE.Value >= 1 Then
            userListw.Add(imtUser1.Text)
        End If
        If imtUser2.Text <> "" And imnUserE.Value >= 2 Then
            userListw.Add(imtUser2.Text)
        End If
        If imtUser3.Text <> "" And imnUserE.Value >= 3 Then
            userListw.Add(imtUser3.Text)
        End If
        If imtUser4.Text <> "" And imnUserE.Value >= 4 Then
            userListw.Add(imtUser4.Text)
        End If
        If imtUser5.Text <> "" And imnUserE.Value >= 5 Then
            userListw.Add(imtUser5.Text)
        End If

        Dim configObjArray(userListw.Count()) As ConfigJsonClass

        '事前に旧ファイルを削除
        For i = 1 To 5
            Dim confFileName = "C:\tool\yt-win-scraping\" + "lpkickrun-config" + i.ToString + ".json"
            System.IO.File.Delete(confFileName)

            Dim batFileName = "C:\tool\yt-win-scraping\" + "lpkickrun-batch" + i.ToString + ".bat"
            System.IO.File.Delete(batFileName)
        Next


        For i = 1 To userListw.Count()
            'ログインユーザ分 JSONファイル作成

            configObjArray(i) = configObj

            Select Case i
                Case 1
                    configObjArray(i).confUser = imtUser1.Text
                    configObjArray(i).confPass = imtPassword1.Text
                Case 2
                    configObjArray(i).confUser = imtUser2.Text
                    configObjArray(i).confPass = imtPassword2.Text
                Case 3
                    configObjArray(i).confUser = imtUser3.Text
                    configObjArray(i).confPass = imtPassword3.Text
                Case 4
                    configObjArray(i).confUser = imtUser4.Text
                    configObjArray(i).confPass = imtPassword4.Text
                Case 5
                    configObjArray(i).confUser = imtUser5.Text
                    configObjArray(i).confPass = imtPassword5.Text
            End Select

            configObjArray(i).targetUrl = imtURL.Text

            Dim tikectNameList = New List(Of String)
            Dim tikectNumberList = New List(Of String)
            For j = 1 To 15
                Select Case j
                    Case 1
                        If imtTicketName1.Text <> "" And imnMaisu1.Text >= 1 Then
                            tikectNameList.Add(imtTicketName1.Text)
                            tikectNumberList.Add(imnMaisu1.Text)
                        End If
                    Case 2
                        If imtTicketName2.Text <> "" And imnMaisu2.Text >= 1 Then
                            tikectNameList.Add(imtTicketName2.Text)
                            tikectNumberList.Add(imnMaisu2.Text)
                        End If
                    Case 3
                        If imtTicketName3.Text <> "" And imnMaisu3.Text >= 1 Then
                            tikectNameList.Add(imtTicketName3.Text)
                            tikectNumberList.Add(imnMaisu3.Text)
                        End If
                    Case 4
                        If imtTicketName4.Text <> "" And imnMaisu4.Text >= 1 Then
                            tikectNameList.Add(imtTicketName4.Text)
                            tikectNumberList.Add(imnMaisu4.Text)
                        End If
                    Case 5
                        If imtTicketName5.Text <> "" And imnMaisu5.Text >= 1 Then
                            tikectNameList.Add(imtTicketName5.Text)
                            tikectNumberList.Add(imnMaisu5.Text)
                        End If
                    Case 6
                        If imtTicketName6.Text <> "" And imnMaisu6.Text >= 1 Then
                            tikectNameList.Add(imtTicketName6.Text)
                            tikectNumberList.Add(imnMaisu6.Text)
                        End If
                    Case 7
                        If imtTicketName7.Text <> "" And imnMaisu7.Text >= 1 Then
                            tikectNameList.Add(imtTicketName7.Text)
                            tikectNumberList.Add(imnMaisu7.Text)
                        End If
                    Case 8
                        If imtTicketName8.Text <> "" And imnMaisu8.Text >= 1 Then
                            tikectNameList.Add(imtTicketName8.Text)
                            tikectNumberList.Add(imnMaisu8.Text)
                        End If
                    Case 9
                        If imtTicketName9.Text <> "" And imnMaisu9.Text >= 1 Then
                            tikectNameList.Add(imtTicketName9.Text)
                            tikectNumberList.Add(imnMaisu9.Text)
                        End If
                    Case 10
                        If imtTicketName10.Text <> "" And imnMaisu10.Text >= 1 Then
                            tikectNameList.Add(imtTicketName10.Text)
                            tikectNumberList.Add(imnMaisu10.Text)
                        End If
                    Case 11
                        If imtTicketName11.Text <> "" And imnMaisu11.Text >= 1 Then
                            tikectNameList.Add(imtTicketName11.Text)
                            tikectNumberList.Add(imnMaisu11.Text)
                        End If
                    Case 12
                        If imtTicketName12.Text <> "" And imnMaisu12.Text >= 1 Then
                            tikectNameList.Add(imtTicketName12.Text)
                            tikectNumberList.Add(imnMaisu12.Text)
                        End If
                    Case 13
                        If imtTicketName13.Text <> "" And imnMaisu13.Text >= 1 Then
                            tikectNameList.Add(imtTicketName13.Text)
                            tikectNumberList.Add(imnMaisu13.Text)
                        End If
                    Case 14
                        If imtTicketName14.Text <> "" And imnMaisu14.Text >= 1 Then
                            tikectNameList.Add(imtTicketName14.Text)
                            tikectNumberList.Add(imnMaisu14.Text)
                        End If
                    Case 15
                        If imtTicketName15.Text <> "" And imnMaisu15.Text >= 1 Then
                            tikectNameList.Add(imtTicketName15.Text)
                            tikectNumberList.Add(imnMaisu15.Text)
                        End If
                End Select
            Next

            If tikectNameList.Count() = 0 Then
                MsgBox("枚数選択されていません")
                Exit Sub
            End If


            configObjArray(i).targetName = String.Join(",", tikectNameList)
            configObjArray(i).targetNumber = String.Join(",", tikectNumberList)


            configObjArray(i).advSec = -245 - ((i - 1) * 3)


            'Json形式データをJson文字列を復元する
            Dim options = New JsonSerializerOptions With {
                .Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
                .WriteIndented = True
            }
            Dim expJsonStr = JsonSerializer.Serialize(configObjArray(i), options)

            'configファイル作成
            Dim confFileName = "C:\tool\yt-win-scraping\" + "lpkickrun-config" + i.ToString + ".json"
            Dim configSW As System.IO.StreamWriter = System.IO.File.CreateText(confFileName)
            configSW.Write(expJsonStr)
            configSW.Close()


            '出力したconfigファイルをlpkickrunフォルダにコピー
            Dim toFolder = "C:\source\yt-scraping\func-lpkick\" + "lpkickrun-config" + i.ToString + ".json"
            File.Copy(confFileName, toFolder, True)

            'batファイル作成
            Dim batFileName = "C:\tool\yt-win-scraping\" + "lpkickrun-batch" + i.ToString + ".bat"
            Dim batSW As System.IO.StreamWriter = System.IO.File.CreateText(batFileName)
            batSW.WriteLine("cd C:\source\yt-scraping\func-lpkick")
            batSW.WriteLine("node lpkickrun " + "C:\tool\yt-win-scraping\lpkickrun-config" + i.ToString + ".json")
            batSW.WriteLine("pause")
            batSW.Close()

        Next

        'Json形式データをJson文字列を復元する
        Dim commonOptions = New JsonSerializerOptions With {
                .Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
                .WriteIndented = True
            }
        Dim expCommonJsonStr = JsonSerializer.Serialize(configCommonObj, commonOptions)

        'commonConfigファイル作成
        Dim commonConfFileName = "C:\tool\yt-win-scraping\" + "lpkickrun-common.json"
        Dim commonConfigSW As System.IO.StreamWriter = System.IO.File.CreateText(commonConfFileName)
        commonConfigSW.Write(expCommonJsonStr)
        commonConfigSW.Close()

        '出力したcommonConfigファイルをlpkickrunフォルダにコピー
        Dim toCommonFolder = "C:\source\yt-scraping\func-lpkick\lpkickrun-common.json"
        File.Copy(commonConfFileName, toCommonFolder, True)

        MsgBox("出力完了しました")



    End Sub

    Private Sub btnOpenFolder_Click(sender As Object, e As EventArgs) Handles btnOpenFolder.Click
        openFolder()
    End Sub

    Private Sub openFolder()
        System.Diagnostics.Process.Start("C:\tool\yt-win-scraping")
    End Sub

End Class

Public Class ConfigCommonJsonClass
    Public Property mypageUrl As String
    Public Property autoLogin As Boolean
    Public Property isJamCheck As Boolean
    Public Property isMailSend As Boolean
    Public Property isClose As Boolean
    Public Property isMobile As Boolean
    Public Property dareSec As Integer
    Public Property isConnect As Boolean
    Public Property isLogger As Boolean
    Public Property jamTryCnt As Integer
    Public Property isProxy As Boolean
    Public Property fromMail As String
    Public Property fromMailPass As String
    Public Property toMail As String
    Public Property toSysMail As String
End Class

Public Class ConfigJsonClass
    Public Property confUser As String
    Public Property confPass As String
    Public Property testUser As String
    Public Property testPass As String
    Public Property targetUrl As String
    Public Property targetName As String
    Public Property targetNumber As String
    Public Property storeTypeId As String
    Public Property loginTime As String
    Public Property startTime As String
    Public Property endTime As String
    Public Property advSec As Integer
    Public Property isBrowse As Boolean
    Public Property isTest As Boolean
    Public Property isCredit As Boolean
    Public Property proxy As String
    Public Property proxyUser As String
    Public Property proxyPass As String
End Class
