<!-- #include file="../inc/connection.inc" -->
<!-- #include file="../scripts/datestring.asp" -->

<%      cr_runnum 			= session("veri_runnum")	
		cr_prostatus		= Trim(Request("cr_prostatus"))
		cr_proremark 		= Replace(Trim(Request("cr_proremark")), "'", "''")
		cr_protime       	= Trim(Request("cr_protime"))
		cr_emailadd			= Trim(Request("cr_emailadd"))
		reportJust			= Trim(Request("reportJust"))

		''''Sending Email''''
		cr_proname 	    	= Trim(Request("cr_proname"))
		cr_hod 		        = Trim(Request("cr_hod"))
		cr_buPIC			= Trim(Request("cr_buPIC"))
		cr_amname			= Trim(Request("cr_amname"))
		cr_scmname 			= Trim(Request("cr_scmname"))
		cr_fdname 			= Trim(Request("cr_fdname"))
		cr_buoname 			= Trim(Request("cr_buoname"))
		cr_executivename 	= Trim(Request("cr_executivename"))
		cr_requester 		= Replace(TRIM(Request("cr_requester")), "'", "''")

	'--------------------
	 	Dim ip_address
			ip_address = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If ip_address = "" Then
  			ip_address = Request.ServerVariables("REMOTE_ADDR")
		End If
	
		dim changes
	    changes = "E-CER No: " & cr_runnum & " - E-CER Approval : (Name : " & cr_proname & "; Status : " & cr_prostatus & "; Process Time : " & cr_protime & "; Remark  : " & cr_proremark & ")"
				
		SQLAudit = "INSERT INTO tbceraudit (cr_runnum, " & _
					"action_user, " & _
					"action_taken, " & _
					"cr_smveristatus, " & _
					"action_date, " & _
					"action_time," & _ 
					"ip_address," & _ 
					"changes" & _ 
					") VALUES (" &_
					"'" & cr_runnum & "'," &_
					"'" & session("EmpNameAccess") & "'," & _
					"'" & "Approval" & "'," & _
					"'" & "-" & "'," & _
					"'" & Day(Date()) & "/" & Month(Date()) & "/" & Year(Date()) & "'," & _
					"'" & time & "'," & _ 
					"'" & ip_address & "'," & _ 
					"'" & changes & "'" & _  
					")"
		Con.Execute SQLAudit
		
		'''''''''''''''''''''''''''''
			
		dim rsSM
		set rsSM =Server.CreateObject("ADODB.Recordset")
		SQLSM = "SELECT EmpNo FROM tbEProfile WHERE EmpName = '" & cr_hod & "'"
		rsSM.Open SQLSM, con
		if Not rsSM.EOF Then
			sm_empno = rsSM("EmpNo")
		else
			set rsSM =Server.CreateObject("ADODB.Recordset")
			SQLSM = "SELECT EmpNo FROM tbEProfileC WHERE EmpName = '" & cr_hod & "'"
			rsSM.Open SQLSM, con
			if Not rsSM.EOF Then
				sm_empno = rsSM("EmpNo")
			end if
		end if
		rsSM.Close
		Set rsSM = Nothing
		
		dim rsSMEmail
		set rsSMEmail =Server.CreateObject("ADODB.Recordset")
		SQLSMEmail = "SELECT email_add FROM tbEmail WHERE emp_no = '" & sm_empno & "'"
		rsSMEmail.Open SQLSMEmail, con
		if Not rsSMEmail.EOF Then
			sm_email = rsSMEmail("email_add")
		end if
		rsSMEmail.Close
		Set rsSMEmail = Nothing
		
		dim rsPRO
		set rsPRO =Server.CreateObject("ADODB.Recordset")
		SQLPRO = "SELECT EmpNo FROM tbEProfile WHERE EmpName = '" & cr_proname & "'"
		rsPRO.Open SQLPRO, con
		if Not rsPRO.EOF Then
			pro_empno = rsPRO("EmpNo")
		else
			set rsPRO =Server.CreateObject("ADODB.Recordset")
			SQLPRO = "SELECT EmpNo FROM tbEProfileC WHERE EmpName = '" & cr_proname & "'"
			rsPRO.Open SQLPRO, con
			if Not rsPRO.EOF Then
				pro_empno = rsPRO("EmpNo")
			end if
		end if
		rsPRO.Close
		Set rsPRO = Nothing
	
		dim rsPROEmail
		set rsPROEmail =Server.CreateObject("ADODB.Recordset")
		SQLPROEmail = "SELECT email_add FROM tbEmail WHERE emp_no = '" & pro_empno & "'"
		rsPROEmail.Open SQLPROEmail, con
		if Not rsPROEmail.EOF Then
			pro_email = rsPROEmail("email_add")
		end if
		rsPROEmail.Close
		Set rsPROEmail = Nothing
		
		dim rsBU
		set rsBU =Server.CreateObject("ADODB.Recordset")
		SQLBU = "SELECT EmpNo FROM tbEProfile WHERE EmpName = '" & cr_buoname & "'"
		rsBU.Open SQLBU, con
		if Not rsBU.EOF Then
			bu_empno = rsBU("EmpNo")
		else
			set rsBU =Server.CreateObject("ADODB.Recordset")
			SQLBU = "SELECT EmpNo FROM tbEProfileC WHERE EmpName = '" & cr_buoname & "'"
			rsBU.Open SQLBU, con
			if Not rsBU.EOF Then
				bu_empno = rsBU("EmpNo")
			end if
		end if
		rsBU.Close
		Set rsBU = Nothing
	
		dim rsBUEmail
		set rsBUEmail =Server.CreateObject("ADODB.Recordset")
		SQLBUEmail = "SELECT email_add FROM tbEmail WHERE emp_no = '" & bu_empno & "'"
		rsBUEmail.Open SQLBUEmail, con
		if Not rsBUEmail.EOF Then
			bu_email = rsBUEmail("email_add")
		end if
		rsBUEmail.Close
		Set rsBUEmail = Nothing
		
		dim rsBUPIC
		set rsBUPIC =Server.CreateObject("ADODB.Recordset")
		SQLBUPIC = "SELECT EmpNo FROM tbEProfile WHERE EmpName = '" & cr_buPIC & "'"
		rsBUPIC.Open SQLBUPIC, con
		if Not rsBUPIC.EOF Then
			buPIC_empno = rsBUPIC("EmpNo")

		end if
		rsBUPIC.Close
		Set rsBUPIC = Nothing
	
		dim rsBUPICEmail
		set rsBUPICEmail =Server.CreateObject("ADODB.Recordset")
		SQLBUPICEmail = "SELECT email_add FROM tbEmail WHERE emp_no = '" & buPIC_empno & "'"
		rsBUPICEmail.Open SQLBUPICEmail, con
		if Not rsBUPICEmail.EOF Then
			buPIC_email = rsBUPICEmail("email_add")
		end if
		rsBUPICEmail.Close
		Set rsBUPICEmail = Nothing

		dim rsUM
		set rsUM =Server.CreateObject("ADODB.Recordset")
		SQLUM = "SELECT EmpNo FROM tbEProfile WHERE EmpName = '" & cr_amname & "'"
		rsUM.Open SQLUM, con
		if Not rsUM.EOF Then
			um_empno = rsUM("EmpNo")
		else
			set rsUM =Server.CreateObject("ADODB.Recordset")
			SQLUM = "SELECT EmpNo FROM tbEProfileC WHERE EmpName = '" & cr_amname & "'"
			rsUM.Open SQLUM, con
			if Not rsUM.EOF Then
				um_empno = rsUM("EmpNo")
			end if
		end if
		rsUM.Close
		Set rsUM = Nothing
		
		dim rsUMEmail
		set rsUMEmail =Server.CreateObject("ADODB.Recordset")
		SQLUMEmail = "SELECT email_add FROM tbEmail WHERE emp_no = '" & um_empno & "'"
		rsUMEmail.Open SQLUMEmail, con
		if Not rsUMEmail.EOF Then
			um_email = rsUMEmail("email_add")
		end if
		rsUMEmail.Close
		Set rsUMEmail = Nothing
		
		dim rsFM
		set rsFM =Server.CreateObject("ADODB.Recordset")
		SQLFM = "SELECT EmpNo FROM tbEProfile WHERE EmpName = '" & cr_scmname & "'"
		rsFM.Open SQLFM, con
		if Not rsFM.EOF Then
			fm_empno = rsFM("EmpNo")
		else
			set rsFM =Server.CreateObject("ADODB.Recordset")
			SQLFM = "SELECT EmpNo FROM tbEProfileC WHERE EmpName = '" & cr_scmname & "'"
			rsFM.Open SQLFM, con
			if Not rsFM.EOF Then
				fm_empno = rsFM("EmpNo")
			end if
		end if
		rsFM.Close
		Set rsFM = Nothing

		dim rsFMEmail
		set rsFMEmail =Server.CreateObject("ADODB.Recordset")
		SQLFMEmail = "SELECT email_add FROM tbEmail WHERE emp_no = '" & fm_empno & "'"
		rsFMEmail.Open SQLFMEmail, con
		if Not rsFMEmail.EOF Then
			fm_email = rsFMEmail("email_add")
		end if
		rsFMEmail.Close
		Set rsFMEmail = Nothing

		dim rsCFO
		set rsCFO =Server.CreateObject("ADODB.Recordset")
		SQLCFO = "SELECT EmpNo FROM tbEProfile WHERE EmpName = '" & cr_fdname & "'"
		rsCFO.Open SQLCFO, con
		if Not rsCFO.EOF Then
			cfo_empno = rsCFO("EmpNo")
		else
			set rsCFO =Server.CreateObject("ADODB.Recordset")
			SQLCFO = "SELECT EmpNo FROM tbEProfileC WHERE EmpName = '" & cr_fdname & "'"
			rsCFO.Open SQLCFO, con
			if Not rsCFO.EOF Then
				cfo_empno = rsCFO("EmpNo")
			end if
		end if
		rsCFO.Close
		Set rsCFO = Nothing

		dim rsCFOEmail
		set rsCFOEmail =Server.CreateObject("ADODB.Recordset")
		SQLCFOEmail = "SELECT email_add FROM tbEmail WHERE emp_no = '" & cfo_empno & "'"
		rsCFOEmail.Open SQLCFOEmail, con
		if Not rsCFOEmail.EOF Then
			cfo_email = rsCFOEmail("email_add")
		end if
		rsCFOEmail.Close
		Set rsCFOEmail = Nothing
		
		dim rsCEO
		set rsCEO =Server.CreateObject("ADODB.Recordset")
		SQLCEO = "SELECT EmpNo FROM tbEProfile WHERE EmpName = '" & cr_executivename & "'"
		rsCEO.Open SQLCEO, con
		if Not rsCEO.EOF Then
			ceo_empno = rsCEO("EmpNo")
		else
			set rsCEO =Server.CreateObject("ADODB.Recordset")
			SQLCEO = "SELECT EmpNo FROM tbEProfileC WHERE EmpName = '" & cr_executivename & "'"
			rsCEO.Open SQLCEO, con
			if Not rsCEO.EOF Then
				ceo_empno = rsCEO("EmpNo")
			end if
		end if
		rsCEO.Close
		Set rsCEO = Nothing

		dim rsCEOEmail
		set rsCEOEmail =Server.CreateObject("ADODB.Recordset")
		SQLCEOEmail = "SELECT email_add FROM tbEmail WHERE emp_no = '" & ceo_empno & "'"
		rsCEOEmail.Open SQLCEOEmail, con
		if Not rsCEOEmail.EOF Then
			ceo_email = rsCEOEmail("email_add")
		end if
		rsCEOEmail.Close
		Set rsCEOEmail = Nothing

		dim rsMLevel
		set rsMLevel =Server.CreateObject("ADODB.Recordset")
		SQLMLevel = "SELECT tbModule_Level.* FROM tbModule_Level WHERE e_module = 'E-CER'"
		rsMLevel.Open SQLMLevel, con

		if cr_prostatus = "InfoReq"  Then
				SQLSMVeri = "UPDATE tbCERMaster SET " & _
							"cr_smveristatus='', cr_notify='' WHERE cr_runnum='" & cr_runnum & "';"
				Con.Execute SQLSMVeri
                
                
		elseif cr_prostatus = "Not Approved" then
				SQLSMVeri = "UPDATE tbCERMaster SET " & _
							"cr_smveristatus='', cr_amstatus='', cr_scmstatus='', cr_prostatus =''," & _
							"cr_buostatus='', cr_fdstatus='', cr_executivestatus='' " & _
							"WHERE cr_runnum='" & cr_runnum & "';"
				Con.Execute SQLSMVeri

		elseif cr_prostatus ="Approved" and reportJust="1" then
		        SQLROI = " update tbCERMaster set cr_roiflag=1 WHERE cr_runnum='" & cr_runnum & "';"
				Con.Execute SQLROI
	
				SQLSMVeri = "UPDATE tbCERMaster SET cr_notify='Notice', cr_prostatus ='', cr_buostatus=''  WHERE cr_runnum='" & cr_runnum & "';"
				Con.Execute SQLSMVeri
		end if

		SQLNotify = "INSERT INTO cer_remark (cr_name, " & _
					"cr_remark, " &_
					"cr_date, " &_
					"cr_runnum " &_
					") VALUES (" &_
					"'" & cr_proname & "'," &_
					"'" & cr_proremark & "'," & _
					"'" & cr_protime & "'," & _
					"'" & cr_runnum & "'" &_ 
					")"
		Con.Execute SQLNotify
		
		SQL = "UPDATE tbCERMaster SET " &_
				"cr_prostatus='" & cr_prostatus & "'," &_
				"cr_proremark='" & cr_proremark & "'," &_
				"cr_protime='" & cr_protime & "'" &_
				"WHERE cr_runnum='" & cr_runnum & "';"					
		Con.Execute SQL
		
%>

<%
	Dim Conn,strSQL,objRec   
	Dim xlApp,xlBook,xlSheet1,FileName,intRows   
	Dim Fso,MyFile   
	dim Mail
	dim strBody
	dim mailbody
	
	set Mail = server.createobject("CDO.Message")
	
	Mail.From = session("cia_email") 
	Mail.BCC = pro_email
	
	if cr_prostatus = "Approved" and reportJust="1" Then
	
		'Mail.To = bu_email 'fm_email & ","  & um_email & "," & cfo_email & "," & ceo_email
		'Mail.Subject = "E-CER New " & cr_runnum & " - " & Request("cr_equidesc") 
		
		'strBody = "E-CER New " & "<br/>" 
		'strBody = strBody & "---------------------------------------" & "<br/>"
		''strBody = strBody & "Please Check The E-CER : " & cr_runnum & "<br/>"
		''strBody = strBody & "Equipment Description : " & Request("cr_equidesc") & "<br/>"
		'strBody = strBody & "Purpose : " & Request("cr_purpose") & "<br/>"
		'strBody = strBody & "Approval Needed : " & rsMLevel("level_3") & "<br/>"
		'strBody = strBody & "Please login " & "<a href='" & session("website") & "'>" & session("website") & "</a><br/>"
			
		Mail.To = cr_emailadd
		Mail.Subject = "E-CER ROI Reminder " & cr_runnum & " - " & Request("cr_equidesc") 
	
		strBody = "<html><body>E-CER New " & "<br/>" 
		strBody = "<p>"&strBody & "---------------------------------------" & "</p><br/>"
		'strBody = strBody & "Please Check The E-CER : " & cr_runnum & "<br/>"
		'strBody = strBody & "Equipment Description : " & Request("cr_equidesc") & "<br/>"
		strBody = "<p>"&strBody & "Purpose : " & Request("cr_purpose") & "</p><br/>"
		strBody = "<p>"&strBody & "ROI Needed : " & cr_requester & "</p><br/>"
		strBody = "<p>"&strBody & "Please login " & "<a href='" & session("website") & "'>" & session("website") & "</a></p><br/></body></html>"
		
	elseif cr_prostatus = "Approved" and reportJust<>"1" Then
	
		'Mail.To =  fm_email & ","  & um_email & "," & buPIC_email & "," & cfo_email 
        'Remove fm_email because Johnny Khong keep received email because his name in cr_scmname
        Mail.To =  um_email & "," & buPIC_email & "," & cfo_email & "," & bu_email
		Mail.Subject = "E-CER New " & cr_runnum & " - " & Request("cr_equidesc") 
		
		strBody = "<html><body>E-CER New " & "<br/>" 
		strBody = "<p>"&strBody & "---------------------------------------" & "</p><br/>"
		'strBody = strBody & "Please Check The E-CER : " & cr_runnum & "<br/>"
		'strBody = strBody & "Equipment Description : " & Request("cr_equidesc") & "<br/>"
		strBody = "<p>"&strBody & "Purpose : " & Request("cr_purpose") & "</p><br/>"
		strBody = "<p>"&strBody & "Approval Needed : " & rsMLevel("level_4") & ", " & rsMLevel("level_5") & ", " & rsMLevel("level_6") & ", " & rsMLevel("level_7") & "." & "</p><br/>"
		strBody = "<p>"&strBody & "Please login " & "<a href='" & session("website") & "'>" & session("website") & "</a></p><br/></body></html>"
	
	elseif cr_prostatus = "InfoReq" Then
	
		Mail.To =  cr_emailadd & ","  & sm_email
		Mail.Subject = "E-CER New " & cr_runnum & " - " & Request("cr_equidesc")
		
		'Initialse strBody string with the body of the e-mail
		strBody = "<html><body>E-CER New " & "<br/>" 
		strBody = "<p>"&strBody & "---------------------------------------" & "</p><br/>"
		'strBody = strBody & "Please Check The E-CER : " & cr_runnum & "<br/>"
		'strBody = strBody & "Equipment Description : " & Request("cr_equidesc") & "<br/>"
		strBody = "<p>"&strBody & "Purpose : " & Request("cr_purpose") & "</p><br/>"
		strBody = "<p>"&strBody & "Info Required By " & rsMLevel("level_2") & "." & "</p><br/>"
        'below one line added on 12March2024
        strBody = "<p>"&strBody & "If HOD have any concern please highlight to " & rsMLevel("level_2") & "." & "</p><br/>"
		strBody = "<p>"&strBody & "Please login " & "<a href='" & session("website") & "'>" & session("website") & "</a></p><br/></body></html>"
		
	elseif cr_prostatus = "Not Approved" Then
	
		Mail.To = cr_emailadd & ","  & sm_email 
		Mail.Subject = "E-CER New " & cr_runnum & " - " & Request("cr_equidesc")
		
		'Initialse strBody string with the body of the e-mail
		strBody = "<html><body>E-CER New " & "<br/>" 
		strBody = "<p>"&strBody & "---------------------------------------" & "</p><br/>"
		'strBody = strBody & "Please Check The E-CER : " & cr_runnum & "<br/>"
		'strBody = strBody & "Equipment Description : " & Request("cr_equidesc") & "<br/>"
		strBody = "<p>"&strBody & "Purpose : " & Request("cr_purpose") & "</p><br/>"
		strBody = "<p>"&strBody & "Not Approved By " & rsMLevel("level_2") & "." & "</p><br/>"
		strBody = "<p>"&strBody & "Please login " & "<a href='" & session("website") & "'>" & session("website") & "</a></p><br/></body></html>"
			
	end if

	Mail.HTMLBody = strBody
	Mail.Send
	set Mail = nothing
	
	Response.Write(strBody)
	Response.Redirect "level2_pending.asp"

rsMLevel.Close
Set rsMLevel = Nothing
%>
