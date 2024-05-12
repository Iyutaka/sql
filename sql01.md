## 1. データベースとの接続
まず、データベースに接続する必要があります。以下のコードは、現在のデータベースにDAOを使用して接続する方法を示しています。
'
Dim db As DAO.Database
Set db = CurrentDb()
'
