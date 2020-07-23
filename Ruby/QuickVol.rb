script_directory = File.dirname(__FILE__)
require File.join(script_directory,"Nx.jar")
java_import "com.nuix.nx.NuixConnection"
java_import "com.nuix.nx.LookAndFeelHelper"
java_import "com.nuix.nx.dialogs.ChoiceDialog"
java_import "com.nuix.nx.dialogs.TabbedCustomDialog"
java_import "com.nuix.nx.dialogs.CommonDialogs"
java_import "com.nuix.nx.dialogs.ProgressDialog"
java_import "com.nuix.nx.dialogs.ProcessingStatusDialog"
java_import "com.nuix.nx.digest.DigestHelper"
java_import "com.nuix.nx.controls.models.Choice"

LookAndFeelHelper.setWindowsIfMetal
NuixConnection.setUtilities($utilities)
NuixConnection.setCurrentNuixVersion(NUIX_VERSION)

require File.join(script_directory,"SuperUtilities.jar")
java_import com.nuix.superutilities.SuperUtilities
java_import com.nuix.superutilities.reporting.SimpleXlsx
java_import com.nuix.superutilities.cases.BulkCaseProcessor

$su = SuperUtilities.init($utilities,NUIX_VERSION)

java_import org.joda.time.DateTime

dialog = TabbedCustomDialog.new("Quick Vol")
main_tab = dialog.addTab("main_tab","Main")
main_tab.appendPathList("case_search_paths")
main_tab.appendSaveFileChooser("report_file","Report File","Excel (*.xlsx)","xlsx")
main_tab.appendCheckBox("allow_migrations","Allow Case Migrations",false)
main_tab.getControl("case_search_paths").setFilesButtonVisible(false)

dialog.validateBeforeClosing do |values|
	if values["case_search_paths"].size < 1
		CommonDialogs.showWarning("Please provide at least one case search path.")
		next false
	end

	if values["report_file"].strip.empty?
		CommonDialogs.showWarning("Please provide a report file path")
		next false
	end

	next true
end

dialog.display
if dialog.getDialogResult == true
	values = dialog.toMap
	case_search_paths = values["case_search_paths"]
	report_file = values["report_file"]

	# Are we appending to an existing file?
	needs_headers = !java.io.File.new(report_file).exists

	xlsx = SimpleXlsx.new(report_file)
	report_sheet = xlsx.getSheet("Report")
	log_sheet = xlsx.getSheet("Log")

	# Looks like this is a new file so we need to write headers
	if needs_headers
		report_sheet.appendRow([
			"Case GUID","Batch Load Guid","Batch Load Date","Total Items",
			"Total Audit Size Bytes","Total File Size Bytes"
		])

		log_sheet.appendRow([
			"Time","Type","Case Directory","Message"
		])
	end

	ProgressDialog.forBlock do |pd|
		pd.embiggen(100)
		case_utility = $su.getCaseUtility
		found_cases = []
		case_search_paths.each do |case_search_path|
			break if pd.abortWasRequested
			pd.setMainStatusAndLogIt("Looking for cases: #{case_search_path}")
			begin
				case_infos = case_utility.findCaseInformation(case_search_path)
				found_cases += case_infos
				pd.logMessage("Found #{case_infos.size} cases")
			rescue Exception => exc
				log_sheet.appendRow([DateTime.new.toString,"ERROR",case_search_path,
					"Error while searching for cases: #{exc.message}"])
			end
		end

		pd.logMessage("Total Cases Found: #{found_cases.size}")
		
		try_again_cases = []

		# First pass, cases with issues are added to try_again_cases and tried again later.  If they have
		# issues a second time they are not retried again after that
		bcp = BulkCaseProcessor.new
		found_cases.each{|ci|bcp.addCaseDirectory(ci.getCaseDirectory)}
		bcp.beforeOpeningCase do |case_info|
			pd.setMainStatusAndLogIt("Processing #{case_info.getCaseDirectory}")
		end

		# If case is locked
		bcp.onCaseIsLocked do |case_locked_info|
			message = "Case locked (will retry later)"
			pd.logMessage(message)
			log_sheet.appendRow([DateTime.new.toString,"ERROR",
				case_locked_info.getCaseInfo.getCaseDirectory,message])
			try_again_cases << case_locked_info.getCaseInfo
		end

		# If there is an error opening case
		bcp.onErrorOpeningCase do |case_open_error|
			message = "Error opening case, '#{case_open_error.getError.getMessage}' (will retry later)"
			log_sheet.appendRow([DateTime.new.toString,"ERROR",
				case_open_error.getCaseInfo.getCaseDirectory,message])
			pd.logMessage(message)
			try_again_cases << case_open_error.getCaseInfo
		end

		# If work function has an error
		bcp.onUserFunctionError do |work_function_error|
			message = "Error reporting on case, '#{work_function_error.getError.getMessage}' (will not retry later)"
			log_sheet.appendRow([DateTime.new.toString,"ERROR",
				work_function_error.getCaseInfo.getCaseDirectory,message])
			pd.logMessage("#{message}:\n#{work_function_error.getError.backtrace.join("\n")}")
		end

		# Attempt to collect data from each case
		bcp.withEachCase do |nuix_case, case_info, case_index, total_cases|
			next false if pd.abortWasRequested
			pd.setMainProgress(case_index,total_cases)
			pd.logMessage("Processing #{nuix_case.getLocation}")
			stats = nuix_case.getStatistics
			nuix_case.getBatchLoads.each do |batch_load_details|
				batch_load_guid = batch_load_details.getBatchId
				batch_load_date = batch_load_details.getLoaded
				query = "batch-load-guid:#{batch_load_guid}"
				total_items = nuix_case.count(query)
				total_audited_bytes = stats.getAuditSize(query)
				total_file_bytes = stats.getFileSize(query)

				report_sheet.appendRow([
					nuix_case.getGuid,
					batch_load_guid,
					batch_load_date.toString,
					total_items,
					total_audited_bytes,
					total_file_bytes
				])
			end
			next true
		end

		# Second pass, we try anything that may have had issue on the first pass, note that this time if something still doesn't
		# work we just leave it at that
		if try_again_cases.size > 0
			pd.logMessage("Retrying #{try_again_cases.size} cases")

			bcp = BulkCaseProcessor.new
			try_again_cases.each{|ci|bcp.addCaseDirectory(ci.getCaseDirectory)}
			bcp.beforeOpeningCase do |case_info|
				pd.setMainStatusAndLogIt("Processing #{case_info.getCaseDirectory}")
			end

			# If case is locked
			bcp.onCaseIsLocked do |case_locked_info|
				message = "Case locked (giving up)"
				log_sheet.appendRow([DateTime.new.toString,"ERROR",
					case_locked_info.getCaseInfo.getCaseDirectory,message])
				pd.logMessage(message)
			end

			# If we have error opening case
			bcp.onErrorOpeningCase do |case_open_error|
				message = "Error opening case, '#{case_open_error.getError.getMessage}' (giving up)"
				log_sheet.appendRow([DateTime.new.toString,"ERROR",
					case_open_error.getCaseInfo.getCaseDirectory,message])
				pd.logMessage(message)
			end

			# If our work function errors
			bcp.onUserFunctionError do |work_function_error|
				message = "Error reporting on case, '#{work_function_error.getError.getMessage}' (giving up)"
				log_sheet.appendRow([DateTime.new.toString,"ERROR",
					work_function_error.getCaseInfo.getCaseDirectory,message])
				pd.logMessage("#{message}:\n#{work_function_error.getError.backtrace.join("\n")}")
			end

			# Attempt to collect data from each case
			bcp.withEachCase do |nuix_case, case_info, case_index, total_cases|
				next false if pd.abortWasRequested
				pd.setMainProgress(case_index,total_cases)
				pd.logMessage("Processing #{nuix_case.getLocation}")
				stats = nuix_case.getStatistics
				nuix_case.getBatchLoads.each do |batch_load_details|
					batch_load_guid = batch_load_details.getBatchId
					batch_load_date = batch_load_details.getLoaded
					query = "batch-load-guid:#{batch_load_guid}"
					total_items = nuix_case.count(query)
					total_audited_bytes = stats.getAuditSize(query)
					total_file_bytes = stats.getFileSize(query)

					report_sheet.appendRow([
						nuix_case.getGuid,
						batch_load_guid,
						batch_load_date.toString,
						total_items,
						total_audited_bytes,
						total_file_bytes
					])

				end
				next true
			end
		end

		report_sheet.autoFitColumns
		log_sheet.autoFitColumns
		xlsx.save

		pd.setCompleted
	end
end