
Imports System.Data.SQLite

''' <summary>
''' SQLite操作用クラス
''' </summary>
''' <remarks></remarks>
Public Class SQLiter

    Private ReadOnly con_string As String = "Data Source=task.db"

#Region " singleton "

    Private Shared ReadOnly _instance As New SQLiter
    Public Shared Function getInstance() As SQLiter
        Return _instance
    End Function

    Private Sub New()
        'Me._con = Me.open()
        'Me.init()
        'Me.close()
    End Sub

#End Region

    Private _con As SQLiteConnection

    ''' <summary>
    ''' DBと切断する
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub close()
        _con.Close()
    End Sub

    ''' <summary>
    ''' DBと接続する
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub open()
        _con = New SQLiteConnection(con_string)
        _con.Open()
    End Sub

    ''' <summary>
    ''' データベース準備
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub init()
        ' todo; DML文はSimple_Task側で生成する

        ' テーブルが存在しない場合、作成する
        Dim sql As New System.Text.StringBuilder
        With sql
            .Append(" create table if not exists task ( ")
            .Append("   Id INTEGER PRIMARY KEY AUTOINCREMENT, ")
            .Append("   Name TEXT, ")
            .Append("   due TEXT,  ")
            .Append("   duration TEXT,  ")
            .Append("   done TEXT  ")
            .Append(" ); ")
        End With
        Me.execute(sql.ToString)
    End Sub

#Region " Execute"

    ''' <summary>
    ''' SQL実行（パラメータ無）
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <remarks></remarks>
    Public Overloads Sub execute(ByVal sql As String)
        Using cmd As SQLiteCommand = Me._con.CreateCommand
            cmd.CommandText = sql.ToString
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>
    ''' SQL実行（パラメータ有）
    ''' </summary>
    ''' <param name="sql">SQL文</param>
    ''' <param name="param">SQLパラメータ(FieldName, Value)</param>
    ''' <remarks></remarks>
    Public Overloads Sub execute(ByVal sql As String, ByVal param As Dictionary(Of String, String))
        Using cmd As SQLiteCommand = Me._con.CreateCommand
            cmd.CommandText = sql.ToString
            ' パラメータを設定
            For Each k As String In param.Keys
                Dim p As SQLiteParameter = cmd.CreateParameter
                p.ParameterName = "@" & k
                p.Value = param.Item(k)
                cmd.Parameters.Add(p)
            Next
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    ''' <summary>
    ''' SELECT SQLを実行し、結果を返す
    ''' </summary>
    ''' <param name="sql"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function select_execute(ByVal sql As String) As DataTable
        Dim dt As New DataTable
        Using adapter As New SQLiteDataAdapter
            adapter.SelectCommand = New SQLiteCommand(sql, Me._con)
            adapter.Fill(dt)
        End Using
        Return dt
    End Function

#End Region

End Class
