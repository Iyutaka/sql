
Dim db As DAO.Database
Set db = CurrentDb()


Dim rs As DAO.Recordset
Set rs = db.OpenRecordset("Employees",dbOpenDynaset)

'レコードをループしてデータを参照'

Do While Not rs.EOF
    Debug.Print rs!ID, rs!Name 'IDとNameのフィールドの値を出力'
    rs.MoveNext
Loop
'リソースの開放'
rs.Close
Set rs = Nothing
Set db = Nothing
'このコードは、Employees テーブルの全レコードをループし、それぞれの ID と Name フィールドの値を Immediate ウィンドウに出力'


Dim rs As DAO.Recordset
Set rs = db.OpenRecordset("SELECT * FROM Employees WHERE Department = 'Sales'")

'結果をループ'
Do While Not rs.EOF
    Debug.Print rs!Name & " works in " & rs!Department
    rs.MoveNext
loop

rs.Close
Set rs = Nothing
Set db = Nothing


'vba'
On Error GoTo ErrorHandler

Dim db As DAO.Database
Dim rs AS DAO.Recordset

Set db = CurrentDb()
Set rs = db.OpenRecordset("SELECT * FROM Employees")

Do While Not rs.EOF
    Debug.Print rs!Name
    rs.MoveNext
loop

ExitHere:
    rs.Close
    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Error" & Err.Number & ": " & Err.Description
    Resume ExitHere

Dim db As DAO.Database
Dim rs As DAO.Recordset
DIm sql As String

Set db = CurrentDb()

'従業員の名前と給料を選択、その給料が所属部署の平均給料よりも高い場合'
sql = "SELECT E.Name, E.Salary FROM Employees E WHERE E.Salary > (SELECT AVG(Salary) FROM Employess WHERE Department = E.Department)"

Set rs = db.OpenRecordset(sql)

If Not rs.EOF Then
    Do While Not rs.EOF
        Debug.Print rs!Name & " earns " & rs!Salary
        rs.MoveNext
    loop
Else
    Debug.Print "No records found."
End if

rs.Close
Set rs = Nothing
Set db = Nothing


Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim sql As String

Set db = CurrentDb()

'副クエリで選択した部署の従業員のみを対象とする'
sql = "SELECT E.Name, E.Department FROM Employees E INNER JOIN (SELECT Department FROM Employees WHERE Location = 'Tokyo') As D ON E.Department = D.Department"

Set rs db = OpenRecordset(sql)

If Not rs.EOF Then
    Do While Not rs.EOF
        Debug.Print rs!Name & " works in " & rs!Department
        rs.MoveNext
    loop
Else
    Debug.Print "No recors found."
End If

rs.Close
Set rs = Nothing
Set db = Nothing

Dim db As DAO.Database
Dim rs As DAO.Recordset
Dim sql As String

Set db = CurrentDb()

' 特定のプロジェクトに関連する従業員のみを選択'
sql = "SELECT Name FROM Employees WHERE ProjectID IN (SELECT ProjectID FROM Projects WHERE ProjectName = 'Development')"

Set rs = db.OpenRecordset(sql)

If Not rs.EOF Then
    Do While Not rs.EOF
        Debug.Print "Employee: " & rs!Name
        rs.MoveNext
    Loop
Else
    Debug.Print "No employees found in the specified projects."
End If

rs.Close
Set rs = Nothing
Set db = Nothing




Sub EfficientDataProcessing()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String

    Set db = CurrentDb()
    sql = "SELECT C, D FROM YourTableName WHERE SomeCondition = True ORDER BY SomeColumn"

    Set rs = db.OpenRecordset(sql, dbOpenSnapshot)  ' Snapshotを使用してメモリ使用量を抑える

    Do While Not rs.EOF
        ' データの処理
        Debug.Print "C: " & rs!C & ", D: " & rs!D
        rs.MoveNext
    Loop

    rs.Close
    Set

 rs = Nothing
    Set db = Nothing
End Sub



Sub RetrieveAndProcessData()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim Cprint As Variant
    Dim Dprint As Variant

    ' データベース接続
    Set db = CurrentDb()

    ' AとBの最小値からデータを取得するクエリ
    sql = "SELECT A, B, C, D FROM YourTableName ORDER BY A ASC, B ASC"
    Set rs = db.OpenRecordset(sql)

    ' レコードセットの走査
    Do While Not rs.EOF
        ' 特定の条件が必要な場合はここに条件を追加
        ' 例: If rs!A = SomeValue And rs!B = SomeOtherValue Then
        Cprint = rs!C
        Dprint = rs!D

        ' 格納したデータの出力
        Debug.Print "Cprint: " & Cprint & ", Dprint: " & Dprint

        rs.MoveNext  ' 次のレコードへ
    Loop

    ' リソースの解放
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub
