
Imports System.Text

Public Class frmMain

    ''' <summary>
    ''' タスククラス
    ''' </summary>
    ''' <remarks></remarks>
    Private Class Task
        Private table As String = "task"
        Public id As String
        Public name As String

        ''' <summary>
        ''' 登録SQLを生成する
        ''' </summary>
        ''' <param name="name">タスク名</param>
        ''' <returns>String INSERT文</returns>
        ''' <remarks></remarks>
        Public Function generateSQL_Add(ByVal name As String) As String
            Dim sql As New StringBuilder
            With sql
                .Append(" insert into task ( ")
                .Append("   name ")
                .Append(" ) values ( ")
                .Append("   @name ")
                .Append(" ); ")
            End With
            Return sql.ToString
        End Function

        ''' <summary>
        ''' 削除SQLを生成する
        ''' </summary>
        ''' <param name="id"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function generateSQL_Delete(ByVal id As String) As String
            Dim sql As New StringBuilder
            With sql
                .Append(" delete from task ")
                .Append(" where Id = @id ")
            End With
            Return sql.ToString
        End Function

        ''' <summary>
        ''' 検索SQLを生成する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function generateSQL_Select() As String
            Dim s As New StringBuilder
            With s
                .Append(" select * ")
                .Append(" from " & Me.table & " ")
                .Append(" order by Id desc ")
            End With
            Return s.ToString
        End Function

    End Class

#Region " Event "

#Region " Form "

    ''' <summary>
    ''' フォームロード時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim task As New Task
        Dim s As SQLiter = SQLiter.getInstance
        With s
            .open()
            .init()
            .close()
        End With
        searchTaskList()
    End Sub

    ''' <summary>
    ''' フォームクローズ後
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub frmMain_FormClosed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.FormClosed
    End Sub

#End Region

#Region " Menu "

    ''' <summary>
    ''' 最前面メニュークリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>最前面設定を切り替える</remarks>
    Private Sub ToolStripMenuItem_TopMost_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem_TopMost.Click

        If Me.TopMost Then
            Me.TopMost = False
            Me.ToolStripMenuItem_TopMost.Text = "最前面"
        Else
            Me.TopMost = True
            Me.ToolStripMenuItem_TopMost.Text = "最前面解除"
        End If

    End Sub

    ''' <summary>
    ''' DB初期化メニュークリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>simple_taskで使用しているSQLiteデータベースファイルを再作成します</remarks>
    Private Sub ToolStripMenuItem_initDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem_initDB.Click

        ' 確認メッセージ
        If MsgBox("SQLiteDBを初期化しますか？", vbOKCancel) = vbCancel Then Exit Sub

        MsgBox("SQLiteDB初期化処理は未実装です！")

    End Sub

#End Region

    ''' <summary>
    ''' タスク名入力時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub txtTask_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTask.TextChanged
        ' 未入力の場合、＋ボタンは押せない
        If Me.txtTask.Text.Trim().Length = 0 Then
            Me.btnAdd.Enabled = False
        Else
            Me.btnAdd.Enabled = True
        End If
    End Sub

    ''' <summary>
    ''' ＋ボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>新しいタスクを作成する</remarks>
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click

        Dim t As New Task
        t.name = txtTask.Text.Trim()

        ' パラメータを作成する
        Dim param As New Dictionary(Of String, String)
        param.Add("name", t.name)

        Dim s As SQLiter = SQLiter.getInstance
        With s
            .open()
            .execute(t.generateSQL_Add(t.name), param)
            .close()
        End With

        searchTaskList()

    End Sub

    ''' <summary>
    ''' ーボタンクリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks>選択されているタスクを削除する</remarks>
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        If Me.dgvTaskList.SelectedRows.Count < 1 Then
            MsgBox("削除するタスクを選択してください")
            Exit Sub
        End If

        Dim t As New Task
        t.id = Me.dgvTaskList.SelectedRows.Item(0).Cells(0).Value.ToString

        ' パラメータを作成する
        Dim param As New Dictionary(Of String, String)
        param.Add("id", t.id)

        Dim s As SQLiter = SQLiter.getInstance
        With s
            .open()
            .execute(t.generateSQL_Delete(t.id), param)
            .close()
        End With

        searchTaskList()

    End Sub

    ''' <summary>
    ''' タスク一覧 行クリック時
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub dgvTaskList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgvTaskList.Click
        ' 行が未選択の場合、－ボタンは押せない
        If Me.dgvTaskList.SelectedRows.Count <= 0 Then
            Me.btnDelete.Enabled = False
        Else
            Me.btnDelete.Enabled = True
        End If
    End Sub

#End Region

    ''' <summary>
    ''' タスク検索結果を一覧に表示する
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub searchTaskList()
        Dim task As New Task
        Dim s As SQLiter = SQLiter.getInstance
        With s
            .open()
            Me.dgvTaskList.DataSource = .select_execute(task.generateSQL_Select())
            .close()
        End With
    End Sub

End Class
