java_import org.joda.time.Days

class HistoryHelper
	attr_accessor :days_ago

	def initialize(days_ago=90)
		@days_ago = days_ago
		@range_start = DateTime.new(DateTimeZone::UTC).minusDays(days_ago).millisOfDay.withMinimumValue
	end

	def days_since_last(event_type,nuix_case)
		range_end = DateTime.new(DateTimeZone::UTC)
		history_settings = {
			"startDateAfter" => @range_start,
			"startDateBefore" => range_end,
			"type" => event_type,
			"order" => "start_date_descending",
		}
		most_recent_event = nil
		nuix_case.getHistory(history_settings).each do |event|
			most_recent_event = event
			break
		end

		if most_recent_event.nil?
			return @days_ago+1
		else
			return Days.daysBetween(most_recent_event.getStartDate,range_end).getDays
		end
	end
end