'KS 
'wscript.echo "################################## KS ##################################"
'wscript.echo "thread Compreensive collect script to native info gattering"

'HV Info schema
Class objDayLight
 Public Erro
 Public ErroDesc
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
Dim strMonth, strDayOfWeek, strDay
strDay = Array("","1","2","3","4","Ultim")
strDayOfWeek = Array("o Dom de ","a Seg de ","a Ter de ","a Qua de ","a Qui de ","a Sex de ","o Sab de ")
strMonth = Array("","Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez")

'Caller
set dmHV = getDayLight("marvin","","")
dmHV.Status = "beta testes"
Out = Join( array(dmHV.Hostname, dmHV.Status, dmHV.TimeZone, dmHV.HVI, dmHV.HVIni, dmHV.HVF, dmHV.HVFim, dmHV.ColIni, dmHV.ColFim), ", ")
logMsg Out

Private Sub logMsg(msg)
 hora = FormatDateTime(now())
 Mensagem = hora & " [INFO]: " & msg 
 wscript.echo Mensagem
End Sub

Function getDayLight (Machine, WMIUser, WMIPass)
 On Error Resume Next 
  
 Set getDayLight = New objDayLight
 Set objWbemLocator = CreateObject("WbemScripting.SWbemLocator")
 getDayLight.ColIni = FormatDateTime(now())
 Err.Clear
 
 Set objWMIService = objwbemLocator.ConnectServer (Machine, "root\cimv2", WMIUser, WMIPass)
'Set objWMIService = objwbemLocator.ConnectServer (strInput, "root\cimv2")
  
 getDayLight.ErroDesc = Err.Description 
 getDayLight.Erro = Err.Number

 If Err.Number = 0 Then
  getDayLight.User = WMIUser
  Set colCSes = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
  Set wmi_timezone = objWMIService.ExecQuery("SELECT * FROM Win32_TimeZone")
  Set wmi_localtime = objWMIService.ExecQuery("SELECT * FROM Win32_LocalTime")
  Set wmi_computersystem = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
  Set wmi_operatingsystem = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
  
  For Each timezone In wmi_timezone
   getDayLight.TimeZone = timezone.Caption
   getDayLight.HVI = timezone.DaylightDay & "-" & timezone.DaylightDayOfWeek & "-" & timezone.DaylightMonth
   getDayLight.HVIni = strDay(timezone.DaylightDay) & strDayOfWeek(timezone.DaylightDayOfWeek) & strMonth(timezone.DaylightMonth)
   getDayLight.HVF = timezone.StandardDay & "-" & timezone.StandardDayOfWeek & "-" & timezone.StandardMonth
   getDayLight.HVFim = strDay(timezone.StandardDay) & strDayOfWeek(timezone.StandardDayOfWeek) & strMonth(timezone.StandardMonth)
  Next
  
  For Each localtime In wmi_localtime
   'getWMI.nCPU = objCS.NumberOfProcessors
  Next
  
  For Each computersystem In wmi_computersystem
   getDayLight.Hostname = computersystem.Name
  Next
  
  For Each operatingsystem In wmi_operatingsystem
   'getWMI.nCPU = objCS.NumberOfProcessors
  Next
  
 End If
  getDayLight.ColFim = FormatDateTime(now()) 
 Return getDayLight
End Function
