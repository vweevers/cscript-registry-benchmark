const HKEY_LOCAL_MACHINE = &H80000002
key = "System\CurrentControlSet\Control\Session Manager\Environment"
valueName = "PROCESSOR_ARCHITECTURE"
expectedValue = Wscript.Arguments(0) 'x86 or x64

If expectedValue = "x64" Then
  expectedValue = "AMD64"
End If

'Agnostic
Set agnostic = GetObject("winmgmts:\root\default:StdRegProv")

' 32
Set reg32_ctx = CreateObject("WbemScripting.SWbemNamedValueSet")
reg32_ctx.Add "__ProviderArchitecture", 32
Set reg32_locator = CreateObject("Wbemscripting.SWbemLocator")
Set reg32_services = reg32_locator.ConnectServer(".", "root\default","","",,,,reg32_ctx)
Set reg32 = reg32_services.Get("StdRegProv") 

' 64
Set reg64_ctx = CreateObject("WbemScripting.SWbemNamedValueSet")
reg64_ctx.Add "__ProviderArchitecture", 64
Set reg64_locator = CreateObject("Wbemscripting.SWbemLocator")
Set reg64_services = reg64_locator.ConnectServer(".", "root\default","","",,,,reg64_ctx)
Set reg64 = reg64_services.Get("StdRegProv") 

Wscript.Echo "cscript path: " & Wscript.FullName
Wscript.Echo "Reading " & key & "\" & valueName & " 500 times" & vbcrlf

For j = 0 To 1
  Wscript.stdout.Write Rpad("Agnostic: ", " ", 30)
  StartTime = Timer()
  For i = 0 To 500
    GetStringValueAgnostic strResult
    VerifyResult strResult
  Next
  endTime = Timer()
  firstDuration = endTime - StartTime
  EndBenchmark StartTime, endTime, null

  Wscript.stdout.Write Rpad("32-bit: ", " ", 30)
  StartTime = Timer()
  For i = 0 To 500
    GetStringValueSpecific reg32, null, reg32_ctx, strResult
    VerifyResult strResult
  Next
  EndBenchmark StartTime, Timer(), firstDuration

  Wscript.stdout.Write Rpad("32-bit (predefined params): ", " ", 30)
  Set params = reg32.Methods_("GetStringValue").Inparameters
  params.hDefKey = HKEY_LOCAL_MACHINE
  params.sSubKeyName = key
  params.sValueName = valueName
  StartTime = Timer()
  For i = 0 To 500
    GetStringValueSpecific reg32, params, reg32_ctx, strResult
    VerifyResult strResult
  Next
  EndBenchmark StartTime, Timer(), firstDuration

  Wscript.stdout.Write Rpad("64-bit: ", " ", 30)
  StartTime = Timer()
  For i = 0 To 500
    GetStringValueSpecific reg64, null, reg64_ctx, strResult
    VerifyResult strResult
  Next
  EndBenchmark StartTime, Timer(), firstDuration

  Wscript.stdout.Write Rpad("64-bit (predefined params): ", " ", 30)
  Set params = reg64.Methods_("GetStringValue").Inparameters
  params.hDefKey = HKEY_LOCAL_MACHINE
  params.sSubKeyName = key
  params.sValueName = valueName
  StartTime = Timer()
  For i = 0 To 500
    GetStringValueSpecific reg64, params, reg64_ctx, strResult
    VerifyResult strResult
  Next
  EndBenchmark StartTime, Timer(), firstDuration

  Wscript.stdout.Write vbcrlf
Next

Sub EndBenchmark(StartTime, EndTime, firstDuration)
  duration = EndTime - StartTime
  Wscript.stdout.Write FormatTime(duration)  & " sec "

  If Not IsNull(firstDuration) Then
    Wscript.stdout.Write "  +" & FormatTime(duration - firstDuration)
  End If

  Wscript.stdout.Write vbcrlf
End Sub

Function FormatTime(inputTime)
  FormatTime = Rpad(Round(inputTime, 3), "0", 5)
End Function

Sub VerifyResult(strResult)
  strResult = Trim(strResult)
  If strResult <> expectedValue Then
    Wscript.Echo "Unexpected value: " & strResult & ", expected " & expectedValue
    Wscript.Quit 1
  End If
End Sub

Sub GetStringValueAgnostic(strResult)
  agnostic.GetStringValue HKEY_LOCAL_MACHINE, key, valueName, strResult
End Sub

Sub GetStringValueSpecific(reg, params, ctx, strResult)  
  If IsNull(params) Then
    Set params = reg.Methods_("GetStringValue").Inparameters
  
    params.hDefKey = HKEY_LOCAL_MACHINE
    params.sSubKeyName = key
    params.sValueName = valueName
  End If

  set Outparams = reg.ExecMethod_("GetStringValue", params,,ctx)
  strResult = Outparams.sValue
End Sub

Function Lpad (inputStr, padChar, lengthStr)  
  Lpad = string(lengthStr - Len(inputStr),padChar) & inputStr  
End Function 

Function Rpad (inputStr, padChar, lengthStr)  
  Rpad = inputStr & string(lengthStr - Len(inputStr), padChar)  
End Function 
