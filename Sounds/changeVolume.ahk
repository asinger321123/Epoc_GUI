SoundGet, MuteState, Master, Mute
if MuteState=On 
{
	MuteState= it's muted
	send {Volume_Mute} ; if it's mute : unmute it
}
SoundSet, 50
return