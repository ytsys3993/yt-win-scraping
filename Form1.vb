'次点：クレカの検証
'500円未満はクレカの義務付けチェック

'実行したバッチをkillできるように


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
    Public amountList = New List(Of String)
    Public maisuList = New List(Of String)

    'パスワード表示フラグ
    Public isPasswordVisible = False

    Public nodesFolder = "C:\source\yt-scraping\func-lpkick"
    Public winFolder = "C:\tool\yt-win-scraping"

    '一意のキー
    Dim kickId = ""


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'タイトル変更
        kickId = Now().ToString("yyyyMMddhhmmss")

        Me.Text = Me.Text + "_" + kickId

        imtURL.Text = My.Settings.URL
        If My.Settings.StartDateTime <> "" Then
            imdStartDateTime.Value = My.Settings.StartDateTime
        End If
        cmbStore.SelectedValue = My.Settings.Store

        Dim titleListw = My.Settings.titleList
        Dim idNameListw = My.Settings.idNameList
        Dim amountListw = My.Settings.amountList
        Dim maisuListw = My.Settings.maisuList

        setDispData(titleListw, idNameListw, amountListw, maisuListw)

        If My.Settings.LoginUser <> "" Then
            imtUser1.Text = My.Settings.LoginUser
        End If
        'imtUser2.Text = "natipo5714@mom2kid.com"
        'imtUser3.Text = "kumyu449@instaddr.win"
        'imtUser4.Text = "mipyagyo@instaddr.win"

        If My.Settings.LoginPassword <> "" Then
            imtPassword1.Text = My.Settings.LoginPassword
        End If
        'imtPassword2.Text = "6yjAWLhCxVT2"
        'imtPassword3.Text = "6yjAWLhCxVT2"
        'imtPassword4.Text = "6yjAWLhCxVT2"


    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        '画面情報保存
        saveDispData()

        '予約情報削除
        deleteReserveData()

    End Sub



    Private Sub btnGetDisp_Click(sender As Object, e As EventArgs) Handles btnGetDisp.Click
        Dim sURL As String
        Dim sHTML As String
        Dim objWC As System.Net.WebClient
        Dim objDOC As HtmlAgilityPack.HtmlDocument
        Dim titleNodes As HtmlAgilityPack.HtmlNodeCollection
        Dim idNodes As HtmlAgilityPack.HtmlNodeCollection
        Dim amountNodes As HtmlAgilityPack.HtmlNodeCollection
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
        '値段取得
        amountNodes = objDOC.DocumentNode.SelectNodes("//*[@id='purchase_form']/section[1]/section/section[*]/section/section[2]/div/p")

        titleList = New List(Of String)
        idNameList = New List(Of String)
        amountList = New List(Of String)
        maisuList = New List(Of String)

        For Each node As HtmlAgilityPack.HtmlNode In titleNodes
            titleList.Add(node.InnerText)
        Next
        For Each node As HtmlAgilityPack.HtmlNode In idNodes
            idNameList.Add(node.Id.Split("_")(3))
        Next
        For Each node As HtmlAgilityPack.HtmlNode In amountNodes
            Dim workInnerText = node.InnerText
            workInnerText = workInnerText.Replace(" ", "").Replace(Chr(13), "").Replace(Chr(10), "").Replace("&yen;", "").Replace("""", "")
            amountList.Add(workInnerText)
        Next

        For index As Integer = 1 To 10
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


        setDispData(titleList, idNameList, amountList, maisuList)

    End Sub

    Private Sub setDispData(titleList As List(Of String), idNameList As List(Of String), amountList As List(Of String), maisuList As List(Of String))

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

        imtAmount1.Text = ""
        imtAmount2.Text = ""
        imtAmount3.Text = ""
        imtAmount4.Text = ""
        imtAmount5.Text = ""
        imtAmount6.Text = ""
        imtAmount7.Text = ""
        imtAmount8.Text = ""
        imtAmount9.Text = ""
        imtAmount10.Text = ""

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

        If titleList IsNot Nothing Then
            For index As Integer = 1 To titleList.Count()
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
        End If

        If idNameList IsNot Nothing Then
            For index As Integer = 1 To idNameList.Count()
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
        End If

        If amountList IsNot Nothing Then
            For index As Integer = 1 To amountList.Count()
                Select Case index
                    Case 1 : imtAmount1.Text = amountList(index - 1)
                    Case 2 : imtAmount2.Text = amountList(index - 1)
                    Case 3 : imtAmount3.Text = amountList(index - 1)
                    Case 4 : imtAmount4.Text = amountList(index - 1)
                    Case 5 : imtAmount5.Text = amountList(index - 1)
                    Case 6 : imtAmount6.Text = amountList(index - 1)
                    Case 7 : imtAmount7.Text = amountList(index - 1)
                    Case 8 : imtAmount8.Text = amountList(index - 1)
                    Case 9 : imtAmount9.Text = amountList(index - 1)
                    Case 10 : imtAmount10.Text = amountList(index - 1)
                End Select
            Next
        End If

        If maisuList IsNot Nothing Then
            For index As Integer = 1 To maisuList.Count()
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
        End If

    End Sub

    Private Sub saveDispData()

        Dim titleListw = New List(Of String)
        Dim idNameListw = New List(Of String)
        Dim amountListw = New List(Of String)
        Dim maisuListw = New List(Of String)

        For index As Integer = 1 To 15
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

        For index As Integer = 1 To 15
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

        For index As Integer = 1 To 15
            amountListw.Add("")
            Select Case index
                Case 1 : amountListw(index - 1) = imtAmount1.Text
                Case 2 : amountListw(index - 1) = imtAmount2.Text
                Case 3 : amountListw(index - 1) = imtAmount3.Text
                Case 4 : amountListw(index - 1) = imtAmount4.Text
                Case 5 : amountListw(index - 1) = imtAmount5.Text
                Case 6 : amountListw(index - 1) = imtAmount6.Text
                Case 7 : amountListw(index - 1) = imtAmount7.Text
                Case 8 : amountListw(index - 1) = imtAmount8.Text
                Case 9 : amountListw(index - 1) = imtAmount9.Text
                Case 10 : amountListw(index - 1) = imtAmount10.Text
                Case 11 : amountListw(index - 1) = imtAmount11.Text
                Case 12 : amountListw(index - 1) = imtAmount12.Text
                Case 13 : amountListw(index - 1) = imtAmount13.Text
                Case 14 : amountListw(index - 1) = imtAmount14.Text
                Case 15 : amountListw(index - 1) = imtAmount15.Text
            End Select
        Next

        For index As Integer = 1 To 15
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

        My.Settings.titleList = titleListw
        My.Settings.idNameList = idNameListw
        My.Settings.amountList = amountListw
        My.Settings.maisuList = maisuListw

        My.Settings.LoginUser = imtUser1.Text
        My.Settings.LoginPassword = imtPassword1.Text



    End Sub

    Private Sub btnDispPassword_Click(sender As Object, e As EventArgs) Handles btnDIspPassword.Click
        If isPasswordVisible = True Then
            imtPassword1.PasswordChar = "*"
            isPasswordVisible = False
        Else
            imtPassword1.PasswordChar = ControlChars.NullChar
            isPasswordVisible = True
        End If
    End Sub

    Private Sub btnReserve_Click(sender As Object, e As EventArgs) Handles btnReserve.Click

        'メッセージボックスを表示する 
        Dim result As DialogResult = MessageBox.Show(
                                        "画面の内容で実行予約してもいいですか？",
                                        "確認",
                                        MessageBoxButtons.YesNo,
                                        MessageBoxIcon.Question,
                                        MessageBoxDefaultButton.Button2)

        '何が選択されたか調べる 
        If result = DialogResult.Yes Then
            '「はい」が選択された時 

            '予約のために操作不能に
            grpURL.Enabled = False
            grpTicket.Enabled = False
            grpLogin.Enabled = False
            grpCommon.Enabled = False

            btnReserve.Enabled = False

            '実行用情報およびバッチ作成、実行
            writeJsonFile()

            'キャンセルおよび終了時処理
            reserveCancel()

        ElseIf result = DialogResult.No Then
            '「いいえ」が選択された時 

        End If


    End Sub

    Private Sub reserveCancel()

        'キャンセルのために操作可能に
        grpURL.Enabled = True
        grpTicket.Enabled = True
        grpLogin.Enabled = True
        grpCommon.Enabled = True

        btnReserve.Enabled = True

        '実行用情報およびバッチ削除
        deleteReserveData()

    End Sub

    Private Sub writeJsonFile()
        Dim enc As Encoding = Encoding.UTF8
        Dim configStr As String = ""
        Dim configFilePath As String = winFolder + "\template\lpkickrun-template.json"

        'ファイルからJson文字列を読み込む
        Using sr As New System.IO.StreamReader(configFilePath, enc)
            configStr = sr.ReadToEnd()
        End Using

        'Json文字列をJson形式データに復元する
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

        '結局ログインユーザは1種類のみとする
        configObj.confUser = imtUser1.Text
        configObj.confPass = imtPassword1.Text


        configObj.targetUrl = imtURL.Text

        Dim tikectNameList = New List(Of String)
        Dim tikectNumberList = New List(Of String)
        For j As Integer = 1 To 15
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

        Dim targetNameArr() As String = tikectNameList.ToArray()
        Dim targetNumberArr() As String = tikectNumberList.ToArray()

        configObj.targetName = String.Join(",", targetNameArr)
        configObj.targetNumber = String.Join(",", targetNumberArr)


        configObj.advSec = -245


        'Json形式データをJson文字列を復元する
        Dim options = New JsonSerializerOptions With {
            .Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
            .WriteIndented = True
        }
        Dim expJsonStr = JsonSerializer.Serialize(configObj, options)

        'configファイル作成
        Dim confFileName = winFolder + "\" + "lpkickrun-" + kickId + "-config.json"
        Dim configSW As System.IO.StreamWriter = System.IO.File.CreateText(confFileName)
        configSW.Write(expJsonStr)
        configSW.Close()

        'batファイル作成
        Dim batFileName = winFolder + "\" + "lpkickrun-" + kickId + "-batch.bat"
        Dim batSW As System.IO.StreamWriter = System.IO.File.CreateText(batFileName)
        batSW.WriteLine("title " + kickId)
        batSW.WriteLine("cd " + nodesFolder)
        batSW.WriteLine("node lpkickrun " + winFolder + "\" + "lpkickrun-" + kickId + "-config.json")
        batSW.WriteLine("pause")
        batSW.Close()

        'ファイルを開いて終了まで待機する
        Dim p As System.Diagnostics.Process =
            System.Diagnostics.Process.Start(batFileName)
        p.WaitForExit()

        MessageBox.Show("終了しました。" &
            vbLf & "終了コード:" & p.ExitCode.ToString() &
            vbLf & "終了時間:" & p.ExitTime.ToString())


    End Sub

    Private Sub btnOpenFolder_Click(sender As Object, e As EventArgs) Handles btnOpenFolder.Click
        openFolder()
    End Sub

    Private Sub openFolder()
        System.Diagnostics.Process.Start("C:\tool\yt-win-scraping")
    End Sub

    '予約データ削除
    Private Sub deleteReserveData()

        'ファイル一覧格納用の配列
        Dim strArrayFiles As String() = Nothing

        'ファイル一覧を配列に格納
        strArrayFiles = System.IO.Directory.GetFiles(winFolder)

        'ファイル一覧をテキストボックスに表示
        For Each strFile As String In strArrayFiles
            If strFile.Contains(kickId) Then
                System.IO.File.Delete(strFile)
            End If
        Next
    End Sub

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
