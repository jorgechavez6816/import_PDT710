Sub Main
	IgnoreWarning(True)
	Call ReportReaderImport1()	'D:\RUC1\DATA\Archivos fuente.ILB\PDT_710_2021.pdf
	Call ReportReaderImport2()
	Call ReportReaderImport3()
	Call ReportReaderImport4()
	Call ReportReaderImport5()
	Call ReportReaderImport6()
	Call AppendDatabase()	'PDT_710_dd2.IMD
	Client.CloseAll
	Set task = Client.ProjectManagement
	task.CreateFolder "_PDT710"
	Set task = Nothing
	Client.DeleteDatabase "PDT_710_activo.IMD"
	Client.DeleteDatabase "PDT_710_pas_patrim.IMD"
	Client.DeleteDatabase "PDT_710_eerr.IMD"
	Client.DeleteDatabase "PDT_710_costos.IMD"
	Client.DeleteDatabase "PDT_710_dd1.IMD"
	Client.DeleteDatabase "PDT_710_dd2.IMD"
	Dim pm As Object
	Dim SourcePath As String
	Dim DestinationPath As String
	Set SourcePath = Client.WorkingDirectory
	Set DestinationPath = "D:\RUC1\DATA\_PDT710"
	Client.RunAtServer False
	Set pm = Client.ProjectManagement
	pm.MoveDatabase SourcePath + "PDT710.IMD", DestinationPath
	Set pm = Nothing
	Client.RefreshFileExplorer
End Sub


' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport1
	dbName = "PDT_710_activo.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\PDT_710_activo.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\PDT_710_2021.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport2
	dbName = "PDT_710_pas_patrim.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\PDT_710_pas_patrim.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\PDT_710_2021.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport3
	dbName = "PDT_710_costos.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\PDT_710_activo.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\PDT_710_2021.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport4
	dbName = "PDT_710_eerr.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\PDT_710_eerr.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\PDT_710_2021.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport5
	dbName = "PDT_710_dd1.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\PDT_710_rtadd1.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\PDT_710_2021.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Archivo - Asistente de importación: Report Reader
Function ReportReaderImport6
	dbName = "PDT_710_dd2.IMD"
	Client.ImportPrintReportEx "D:\RUC1\DATA\Definiciones de importación.ILB\PDT_710_rtadd2.jpm", "D:\RUC1\DATA\Archivos fuente.ILB\PDT_710_2021.pdf", dbname, FALSE, FALSE
	Client.OpenDatabase (dbName)
End Function

' Archivo: Anexar bases de datos
Function AppendDatabase
	Set db = Client.OpenDatabase("PDT_710_activo.IMD")
	Set task = db.AppendDatabase
	task.AddDatabase "PDT_710_pas_patrim.IMD"
	task.AddDatabase "PDT_710_eerr.IMD"
	task.AddDatabase "PDT_710_costos.IMD"
	task.AddDatabase "PDT_710_dd1.IMD"
	task.AddDatabase "PDT_710_dd2.IMD"
	dbName = "PDT710.IMD"
	task.PerformTask dbName, ""
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function