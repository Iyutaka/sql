
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


Sub ProcessCSVData()
    Dim conn As ADODB.Connection
    Dim rsCSV As ADODB.Recordset
    Dim rsDB As ADODB.Recordset
    Dim inputFile As String
    Dim outputFile As String
    Dim line As String
    Dim dataA As String
    Dim dataB As Long
    Dim firstDigit As String
    Dim tenthDigit As String
    Dim sql As String
    Dim fso As Object
    Dim textStream As Object

    ' ファイルパス設定
    inputFile = "C:\path\to\yourfile.csv"
    outputFile = "C:\path\to\outputfile.txt"

    ' FileSystemObjectの初期化
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' テキストファイル出力の準備
    Set textStream = fso.CreateTextFile(outputFile, True)

    ' データベース接続
    Set conn = New ADODB.Connection
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\path\to\yourdatabase.accdb"

    ' CSVファイルオープン
    Set rsCSV = New ADODB.Recordset
    rsCSV.Open "SELECT * FROM [Text;HDR=No;FMT=Delimited;CharacterSet=65001;DATABASE=" & _
               fso.GetParentFolderName(inputFile) & "].[" & fso.GetFileName(inputFile) & "]", _
               conn, adOpenStatic, adLockReadOnly

    ' CSVデータの処理
    Do While Not rsCSV.EOF
        line = rsCSV.Fields(0).Value
        firstDigit = Left(line, 1)
        tenthDigit = Mid(line, 10, 1)
        dataA = Mid(line, 2, 3)
        dataB = Val(Mid(line, 5, 5))

        ' データベース検索クエリ
        sql = "SELECT C, D FROM Table1 WHERE A = '" & dataA & "' AND B = " & dataB
        Set rsDB = New ADODB.Recordset
        rsDB.Open sql, conn, adOpenStatic, adLockReadOnly

        If Not rsDB.EOF Then
            ' テキストファイルにデータを書き込み
            textStream.WriteLine firstDigit & ", " & rsDB("C").Value & ", " & rsDB("D").Value & ", " & tenthDigit
        Else
            textStream.WriteLine firstDigit & ", " & dataA & ", " & dataB & ", " & tenthDigit ' マッチしない場合は元の値を出力
        End If

        rsDB.Close
        rsCSV.MoveNext
    Loop

    ' オブジェクトの解放
    rsCSV.Close
    conn.Close
    textStream.Close
    Set rsDB = Nothing
    Set rsCSV = Nothing
    Set conn = Nothing
    Set textStream = Nothing
    Set fso = Nothing
End Sub