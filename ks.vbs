'KS 
wscript.echo "################################## KS ##################################"
wscript.echo "thread Compreensive collect script to native info gattering"

'HV Info schema
Class objHVInfo
 Public Hostname
 Public Ipaddress
 Public Status
 Public TimeZone
 Public DlBias
 Public TimeDate
 Public TimeDiff
 Public HVFim
 Public HVIni
 Public Domain
 Public User
 Public SO
 Public UUID 'Win32_ComputerSystemProduct
 Public HVI
 Public HVF
 Public ColIni
 Public ColFim
End Class

'Codigo de traducao
Dim strMonth, strDayOfWeek, strDay
 strDay = Array("","1","2","3","4","Ultim")
 strDayOfWeek = Array("o Dom de ","a Seg de ","a Ter de ","a Qua de ","a Qui de ","a Sex de ","o Sab de ")
 strMonth = Array("","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez")

'WMI Connect
 Const HKEY_LOCAL_MACHINE = &H80000002
 Set mHV = New objHVInfo
 Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
 Err.Clear
Set objWMIService = objwbemLocator.ConnectServer (strInput, "root\cimv2", strUser, strPassword)
'Set objWMIService = objwbemLocator.ConnectServer (strInput, "root\cimv2")

'WMI Collect
 Set colCSes = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
 Set wmi_timezone = objWMIService.ExecQuery("SELECT * FROM Win32_TimeZone")
 Set wmi_localtime = objWMIService.ExecQuery("SELECT * FROM Win32_LocalTime")
 Set wmi_computersystem = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
 Set wmi_operatingsystem = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")

 For Each timezone In wmi_timezone
  mHV.TimeZone = timezone.Caption
  mHV.HVI = timezone.DaylightDay & "-" & timezone.DaylightDayOfWeek & "-" & timezone.DaylightMonth
  mHV.HVIni = strDay(timezone.DaylightDay) & strDayOfWeek(timezone.DaylightDayOfWeek) & strMonth(timezone.DaylightMonth)
  mHV.HVF = timezone.StandardDay & "-" & timezone.StandardDayOfWeek & "-" & timezone.StandardMonth
  mHV.HVFim = strDay(timezone.StandardDay) & strDayOfWeek(timezone.StandardDayOfWeek) & strMonth(timezone.StandardMonth)
 Next
 
 For Each localtime In wmi_localtime
  'getWMI.nCPU = objCS.NumberOfProcessors
 Next

 For Each computersystem In wmi_computersystem
  mHV.Hostname = computersystem.Name
 Next

 For Each operatingsystem In wmi_operatingsystem
  'getWMI.nCPU = objCS.NumberOfProcessors
 Next
 
 
 Out = Join( array(mHV.Hostname, mHV.TimeZone, mHV.HVI, mHV.HVIni, mHV.HVF, mHV.HVFim), ", ")
 
' wscript.Echo Out
 logMsg Out
 
 
 
 
 
Private Sub logMsg(msg)
 hora = FormatDateTime(now())
 Mensagem = hora & " [INFO]: " & msg 
 wscript.echo Mensagem
End Sub
