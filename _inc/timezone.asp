<%
Function getTimeZone(fDefault)

	getTimeZone = "<select name=""selTimeZone"" id=""selTimeZone"" class=""oText"" onChange=""doChange();"" style=""width:260;"">" & vbCrLf & _
				vbTab & "<option value=""-12""" & getDefault(0,fDefault,"-12") & ">(GMT -12:00) Eniwetok</option>" & vbCrLf & _
				vbTab & "<option value=""-11""" & getDefault(0,fDefault,"-11") & ">(GMT -11:00) Samoa</option>" & vbCrLf & _
				vbTab & "<option value=""-10""" & getDefault(0,fDefault,"-10") & ">(GMT -10:00) Hawaii</option>" & vbCrLf & _
				vbTab & "<option value=""-9""" & getDefault(0,fDefault,"-9") & ">(GMT -09:00) Alaska</option>" & vbCrLf & _
				vbTab & "<option value=""-8""" & getDefault(0,fDefault,"-8") & ">(GMT -08:00) Pacific</option>" & vbCrLf & _
				vbTab & "<option value=""-7""" & getDefault(0,fDefault,"-7") & ">(GMT -07:00) Mountain</option>" & vbCrLf & _
				vbTab & "<option value=""-6""" & getDefault(0,fDefault,"-6") & ">(GMT -06:00) Central</option>" & vbCrLf & _
				vbTab & "<option value=""-5""" & getDefault(0,fDefault,"-5") & ">(GMT -05:00) Eastern</option>" & vbCrLf & _
				vbTab & "<option value=""-4""" & getDefault(0,fDefault,"-4") & ">(GMT -04:00) Atlantic</option>" & vbCrLf & _
				vbTab & "<option value=""-3""" & getDefault(0,fDefault,"-3") & ">(GMT -03:00) Buenos Aires</option>" & vbCrLf & _
				vbTab & "<option value=""-2""" & getDefault(0,fDefault,"-2") & ">(GMT -02:00) Mid Atlantic</option>" & vbCrLf & _
				vbTab & "<option value=""-1""" & getDefault(0,fDefault,"-1") & ">(GMT -01:00) Azores, West Africa</option>" & vbCrLf & _
				vbTab & "<option value=""0""" & getDefault(0,fDefault,"0") & ">(GMT 00:00) London, Casablanca</option>" & vbCrLf & _
				vbTab & "<option value=""1""" & getDefault(0,fDefault,"1") & ">(GMT 01:00) Paris, Berlin, Arni</option>" & vbCrLf & _
				vbTab & "<option value=""2""" & getDefault(0,fDefault,"2") & ">(GMT 02:00) Kiev, Helsinki, Athens</option>" & vbCrLf & _
				vbTab & "<option value=""3""" & getDefault(0,fDefault,"3") & ">(GMT 03:00) Kuwait, Moscow</option>" & vbCrLf & _
				vbTab & "<option value=""4""" & getDefault(0,fDefault,"4") & ">(GMT 04:00) Abu Dhabi</option>" & vbCrLf & _
				vbTab & "<option value=""5""" & getDefault(0,fDefault,"5") & ">(GMT 05:00) Maldives, Islamabad</option>" & vbCrLf & _
				vbTab & "<option value=""6""" & getDefault(0,fDefault,"6") & ">(GMT 06:00) Bangladesh, Dhaka</option>" & vbCrLf & _
				vbTab & "<option value=""7""" & getDefault(0,fDefault,"7") & ">(GMT 07:00) Bangkok, Hanoi</option>" & vbCrLf & _
				vbTab & "<option value=""8""" & getDefault(0,fDefault,"8") & ">(GMT 08:00) Hong Kong, Perth</option>" & vbCrLf & _
				vbTab & "<option value=""9""" & getDefault(0,fDefault,"9") & ">(GMT 09:00) Tokio, Seoul</option>" & vbCrLf & _
				vbTab & "<option value=""10""" & getDefault(0,fDefault,"10") & ">(GMT 10:00) Sydney, Melbourne</option>" & vbCrLf & _
				vbTab & "<option value=""11""" & getDefault(0,fDefault,"11") & ">(GMT 11:00) Magadan</option>" & vbCrLf & _
				vbTab & "<option value=""12""" & getDefault(0,fDefault,"12") & ">(GMT 12:00) Wellington</option>" & vbCrLf & _
			"</select>" & vbCrLf
End Function
%>