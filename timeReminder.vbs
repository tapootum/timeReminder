setTime=InputBox("Enter time in minutes for reminder","Set Time For Reminder")
if setTime != "" then
    MsgMessage=InputBox("Remine Message","Message")
do until IsNumeric(setTimer)=True
  dTimer=InputBox("Invalid Entry" & vbnewline & vbnewline & _
         "Enter timer interval in minutes","Set Timer")
loop

if dTimer<>"" then
do
  WScript.Sleep dTimer*60*1000 'convert from minutes to milliseconds
  t=MsgBox(MsgMessage & vbnewline & vbnewline & "Restart Timer?", _
    vbYesNo, "It's been " & dTimer &" minute(s)")
  if t=6 then
       'continue loop
  else 'exit loop
       exit do
  end if
loop
end if

