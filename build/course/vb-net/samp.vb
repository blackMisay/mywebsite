Imports ADODB
Imports Scripting
Public Class Main
    Dim connSQL As ADODB.Connection
    Dim connBO As ADODB.Connection
    Dim connPRODDB As ADODB.Connection
    Dim connLOS As ADODB.Connection

    Dim dbcmd As New ADODB.Command
    
    Dim rsSQL As ADODB.Recordset
    Dim rsBO As ADODB.Recordset
    Dim rsVER As ADODB.Recordset
    Dim rsDEL As ADODB.Recordset
    Dim rsSTAT As ADODB.Recordset
    
    Public fs As FileSystemObject
    Public runref As String
    Private adUseClient As CursorLocationEnum
    
    Private Sub Main_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Parameters As String ' here's where the parameter will go.
        Dim file_str() As String, ct As String
        Dim conn_uid As String, conn_pwd As String, conn_svr As String,
            conn_CTR As String, conn_CTRdb As String, conn_prodsvr As String,
            conn_proddb As String, los_svr As String
        Dim dtFRom As String, dtTO As String 'Line 45 to Line 47
        Dim runref As String
        Dim boolWriteOff As Boolean

        'Parameters = Command$()
        Parameters = "samson_sv|samsonta|UATBOT|SICCTEST|SAMSON|20180228001350|SICCTEST|STAGEDB|INDUSTEST"
        'Parameters = "samson_sv|sams0nsys4sv|SBCBO|FINCONSVR|SAMSON|20170831062046|SICC-DBSRVR|PRODDB|INDUSTEST"
        If Parameters <> "" Then
            file_str = Split(Parameters, "|")
            conn_uid = file_str(0)
            conn_pwd = file_str(1)
            conn_svr = file_str(2)
            conn_CTR = file_str(3)
            conn_CTRdb = file_str(4)
            runref = file_str(5)
            conn_prodsvr = file_str(6)
            conn_proddb = file_str(7)
            los_svr = file_str(8)

            'since date from samson system is yyyymmdd
            'dtFRom = Right(dtFRom, 2) & "/" & Mid(dtFRom, 5, 2) & "/" & Left(dtFRom, 4)
            'dtTO = Right(dtTO, 2) & "/" & Mid(dtTO, 5, 2) & "/" & Left(dtTO, 4)

            connBO = New ADODB.Connection
            connBO.ConnectionTimeout = 0
            connBO.CommandTimeout = 0
            'connBO.ConnectionString = "Provider=MSDAORA.1;User ID=sbcards;Data Source=sbcbo;Persist Security Info=False;password=sbcards"
            connBO.ConnectionString = "Provider=MSDAORA.1;User ID=" & conn_uid & ";Data Source=" & conn_svr & ";Persist Security Info=False;password=" & conn_pwd
            connBO.Open()

            '    Set connLOS = New ADODB.Connection
            '    connLOS.ConnectionTimeout = 0
            '    connLOS.CommandTimeout = 0
            '    'connLOS.ConnectionString = "Provider=MSDAORA.1;User ID=sysdev;Data Source=sbcbot;Persist Security Info=False;password=sysdev"
            '    connLOS.ConnectionString = "Provider=MSDAORA.1;User ID=" & conn_uid & ";Data Source=" & los_svr & ";Persist Security Info=False;password=" & conn_pwd
            '    connLOS.Open

            connSQL = New ADODB.Connection
            connSQL.ConnectionString = "Provider=SQLOLEDB.1;user id=" & conn_uid & ";Persist Security Info=False;password=" & conn_pwd & ";Initial Catalog=" & conn_CTRdb & ";Data Source=" & conn_CTR
            'connSQL.ConnectionString = "Provider=SQLOLEDB.1;user id=sysdev;Persist Security Info=False;password=sysdev;Initial Catalog=sysdevdb;Data Source=SICCTEST"
            'connSQL.ConnectionString = "Provider=SQLOLEDB.1;user id=samson_sv;Persist Security Info=False;password=samsonta;Initial Catalog=samson;Data Source=SICCTEST"
            connSQL.CursorLocation = adUseClient
            connSQL.ConnectionTimeout = 0
            connSQL.CommandTimeout = 0
            connSQL.Open()

            connPRODDB = New ADODB.Connection
            connPRODDB.ConnectionString = "Provider=SQLOLEDB.1;user id=" & conn_uid & ";Persist Security Info=False;password=" & conn_pwd & ";Initial Catalog=" & conn_proddb & ";Data Source=" & conn_prodsvr
            'connPRODDB.ConnectionString = "Provider=SQLOLEDB.1;user id=sysdev;Persist Security Info=False;password=sysdev;Initial Catalog=sysdevdb;Data Source=SICCTEST"
            connPRODDB.CursorLocation = adUseClient
            connPRODDB.ConnectionTimeout = 0
            connPRODDB.CommandTimeout = 0
            connPRODDB.Open()


            'chkCONAMT (runREF)
            chkCONAMT_45371(runref)
            chkStatusAndCMTPTPTP(runref)
            getDetails(runref)
            'GenerateREPORT (runREF)
            GenerateREPORT_45371(runref)

        End If
    End Sub
End Class