<%
Function getStates(fWidth,fName,fDefault)

	getStates = "<select name=""" & fName & """ id=""" & fName & """ class=""oText"" style=""width:" & fWidth & "px;"" onChange=""doChange();"">" & vbCrLf & _
				vbTab & "<option></option>" & vbCrLf & _
				vbTab & "<option value=""AK""" & getDefault(0,fDefault,"AK") & ">Alaska</option>" & vbCrLf & _
				vbTab & "<option value=""AL""" & getDefault(0,fDefault,"AL") & ">Alabama</option>" & vbCrLf & _
				vbTab & "<option value=""AR""" & getDefault(0,fDefault,"AR") & ">Arkansas</option>" & vbCrLf & _
				vbTab & "<option value=""AZ""" & getDefault(0,fDefault,"AZ") & ">Arizona</option>" & vbCrLf & _
				vbTab & "<option value=""CA""" & getDefault(0,fDefault,"CA") & ">California</option>" & vbCrLf & _
				vbTab & "<option value=""CO""" & getDefault(0,fDefault,"CO") & ">Colorada</option>" & vbCrLf & _
				vbTab & "<option value=""CT""" & getDefault(0,fDefault,"CT") & ">Conneticut</option>" & vbCrLf & _
				vbTab & "<option value=""DC""" & getDefault(0,fDefault,"DC") & ">District of Colombia</option>" & vbCrLf & _
				vbTab & "<option value=""DE""" & getDefault(0,fDefault,"DE") & ">Delaware</option>" & vbCrLf & _
				vbTab & "<option value=""FL""" & getDefault(0,fDefault,"FL") & ">Florida</option>" & vbCrLf & _
				vbTab & "<option value=""GA""" & getDefault(0,fDefault,"GA") & ">Georgia</option>" & vbCrLf & _
				vbTab & "<option value=""GU""" & getDefault(0,fDefault,"GU") & ">Guam</option>" & vbCrLf & _
				vbTab & "<option value=""HI""" & getDefault(0,fDefault,"HI") & ">Hawaii</option>" & vbCrLf & _
				vbTab & "<option value=""IA""" & getDefault(0,fDefault,"IA") & ">Iowa</option>" & vbCrLf & _
				vbTab & "<option value=""ID""" & getDefault(0,fDefault,"ID") & ">Idaho</option>" & vbCrLf & _
				vbTab & "<option value=""IL""" & getDefault(0,fDefault,"IL") & ">Illinois</option>" & vbCrLf & _
				vbTab & "<option value=""IN""" & getDefault(0,fDefault,"IN") & ">Indiana</option>" & vbCrLf & _
				vbTab & "<option value=""KS""" & getDefault(0,fDefault,"KS") & ">Kansas</option>" & vbCrLf & _
				vbTab & "<option value=""KY""" & getDefault(0,fDefault,"KY") & ">Kentucky</option>" & vbCrLf & _
				vbTab & "<option value=""LA""" & getDefault(0,fDefault,"LA") & ">Louisiana</option>" & vbCrLf & _
				vbTab & "<option value=""MA""" & getDefault(0,fDefault,"MA") & ">Massachusetts</option>" & vbCrLf & _
				vbTab & "<option value=""MD""" & getDefault(0,fDefault,"MD") & ">Maryland</option>" & vbCrLf & _
				vbTab & "<option value=""ME""" & getDefault(0,fDefault,"ME") & ">Maine</option>" & vbCrLf & _
				vbTab & "<option value=""MI""" & getDefault(0,fDefault,"MI") & ">Michigan</option>" & vbCrLf & _
				vbTab & "<option value=""MN""" & getDefault(0,fDefault,"MN") & ">Minnesota</option>" & vbCrLf & _
				vbTab & "<option value=""MO""" & getDefault(0,fDefault,"MO") & ">Missouri</option>" & vbCrLf & _
				vbTab & "<option value=""MS""" & getDefault(0,fDefault,"MS") & ">Mississippi</option>" & vbCrLf & _
				vbTab & "<option value=""MT""" & getDefault(0,fDefault,"MT") & ">Montana</option>" & vbCrLf & _
				vbTab & "<option value=""NC""" & getDefault(0,fDefault,"NC") & ">North Carolina</option>" & vbCrLf & _
				vbTab & "<option value=""ND""" & getDefault(0,fDefault,"ND") & ">North Dakota</option>" & vbCrLf & _
				vbTab & "<option value=""NE""" & getDefault(0,fDefault,"NE") & ">Nebraska</option>" & vbCrLf & _
				vbTab & "<option value=""NH""" & getDefault(0,fDefault,"NH") & ">New Hampshire</option>" & vbCrLf & _
				vbTab & "<option value=""NJ""" & getDefault(0,fDefault,"NJ") & ">New Jersey</option>" & vbCrLf & _
				vbTab & "<option value=""NM""" & getDefault(0,fDefault,"NM") & ">New Mexico</option>" & vbCrLf & _
				vbTab & "<option value=""NV""" & getDefault(0,fDefault,"NV") & ">Nevada</option>" & vbCrLf & _
				vbTab & "<option value=""NY""" & getDefault(0,fDefault,"NY") & ">New York</option>" & vbCrLf & _
				vbTab & "<option value=""OH""" & getDefault(0,fDefault,"OH") & ">Ohio</option>" & vbCrLf & _
				vbTab & "<option value=""OK""" & getDefault(0,fDefault,"OK") & ">Oklahoma</option>" & vbCrLf & _
				vbTab & "<option value=""OR""" & getDefault(0,fDefault,"OR") & ">Oregon</option>" & vbCrLf & _
				vbTab & "<option value=""PA""" & getDefault(0,fDefault,"PA") & ">Pennsylvania</option>" & vbCrLf & _
				vbTab & "<option value=""PR""" & getDefault(0,fDefault,"PR") & ">Puerto Rico</option>" & vbCrLf & _
				vbTab & "<option value=""RI""" & getDefault(0,fDefault,"RI") & ">Rhode Island</option>" & vbCrLf & _
				vbTab & "<option value=""SC""" & getDefault(0,fDefault,"SC") & ">South Carolina</option>" & vbCrLf & _
				vbTab & "<option value=""SD""" & getDefault(0,fDefault,"SD") & ">South Dakota</option>" & vbCrLf & _
				vbTab & "<option value=""TN""" & getDefault(0,fDefault,"TN") & ">Tennessee</option>" & vbCrLf & _
				vbTab & "<option value=""TX""" & getDefault(0,fDefault,"TX") & ">Texas</option>" & vbCrLf & _
				vbTab & "<option value=""UT""" & getDefault(0,fDefault,"UT") & ">Utah</option>" & vbCrLf & _
				vbTab & "<option value=""VT""" & getDefault(0,fDefault,"VT") & ">Vermont</option>" & vbCrLf & _
				vbTab & "<option value=""VA""" & getDefault(0,fDefault,"VA") & ">Virginia</option>" & vbCrLf & _
				vbTab & "<option value=""WA""" & getDefault(0,fDefault,"WA") & ">Washington</option>" & vbCrLf & _
				vbTab & "<option value=""WI""" & getDefault(0,fDefault,"WA") & ">Wisconsin</option>" & vbCrLf & _
				vbTab & "<option value=""WV""" & getDefault(0,fDefault,"WV") & ">West Virginia</option>" & vbCrLf & _
				vbTab & "<option value=""WY""" & getDefault(0,fDefault,"WY") & ">Wyoming</option>" & vbCrLf & _
				vbTab & "<option>------------</option>" & vbCrLf & _
				vbTab & "<option value=""AB""" & getDefault(0,fDefault,"AB") & ">Alberta</option>" & vbCrLf & _
				vbTab & "<option value=""BC""" & getDefault(0,fDefault,"BC") & ">British Columbia</option>" & vbCrLf & _
				vbTab & "<option value=""MB""" & getDefault(0,fDefault,"MB") & ">Manitoba</option>" & vbCrLf & _
				vbTab & "<option value=""NB""" & getDefault(0,fDefault,"NB") & ">New Brunswick</option>" & vbCrLf & _
				vbTab & "<option value=""NF""" & getDefault(0,fDefault,"NF") & ">Newfoundland</option>" & vbCrLf & _
				vbTab & "<option value=""NS""" & getDefault(0,fDefault,"NS") & ">Nova Scotia</option>" & vbCrLf & _
				vbTab & "<option value=""ON""" & getDefault(0,fDefault,"ON") & ">Ontario</option>" & vbCrLf & _
				vbTab & "<option value=""PE""" & getDefault(0,fDefault,"PE") & ">Prince Edward Island</option>" & vbCrLf & _
				vbTab & "<option value=""QC""" & getDefault(0,fDefault,"QC") & ">Quebec</option>" & vbCrLf & _
				vbTab & "<option value=""SK""" & getDefault(0,fDefault,"SK") & ">Saskatchewan</option>" & vbCrLf & _
				vbTab & "<option value=""NT""" & getDefault(0,fDefault,"NT") & ">Northwest Territories</option>" & vbCrLf & _
				vbTab & "<option value=""YT""" & getDefault(0,fDefault,"YT") & ">Yukon</option>" & vbCrLf & _
				"</select>" & vbCrLf
End Function

Function getCountries(fWidth,fName,fDefault)

	getCountries = "<select name=""" & fName & """ id=""" & fName & """ class=""oText"" style=""width:" & fWidth & "px;"" onChange=""doChange();"">" & vbCrLf & _
			vbTab & "<option></option>" & _
			vbTab & "<option value=""CA""" & getDefault(0,fDefault,"CA") & ">Canada</option>" & vbCrLf & _
			vbTab & "<option value=""US""" & getDefault(0,fDefault,"US") & ">United States</option>" & vbCrLf & _
			vbTab & "<option value=""UK""" & getDefault(0,fDefault,"UK") & ">United Kingdom</option>" & vbCrLf & _
			"</select>" & vbCrLf
End Function
%>