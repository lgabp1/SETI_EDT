from pathlib import Path
import datetime
import re
import icalendar

PATH_FOLDER = Path(__file__).parent
PATH_RAW_CAL = PATH_FOLDER / 'EDT_raw.csv'
PATH_CAL_EXPORT = PATH_FOLDER / 'EDT.ics'

# Mapping of French month abbreviations to month numbers
french_months = {
    "janv.": 1, "févr.": 2, "mars": 3, "avr.": 4, "mai": 5, "juin": 6,
    "juil.": 7, "août": 8, "sept.": 9, "oct.": 10, "nov.": 11, "déc.": 12
}

def parse_french_date(date_str: str, year:int=2025):
    day, month_abbr = date_str.split("-")
    month = french_months[month_abbr.lower()]
    return datetime.datetime.strptime(f"{day}-{month}-{year}", "%d-%m-%Y").date()

class Event:
    def __init__(self, date: datetime.date, start_time: str, end_time: str, description: str):
        self.date = date
        self.start_time = start_time
        self.end_time = end_time
        self.description = description

    def __repr__(self):
        return f"Event(date={self.date}, start_time={self.start_time}, end_time={self.end_time}, description={self.description.replace('\n', '\t')})"

PATH_FILES = PATH_FOLDER.glob('s*.txt')

RE_TIME = re.compile(r'(?P<hours1>\d{1,2})h(?P<minutes1>\d{1,2})? - (?P<hours2>\d{1,2})h(?P<minutes2>\d{1,2})?')

def main():
    # Parse for events
    events: list[Event] = []

    for pf in PATH_FILES:
        with pf.open('r', encoding='utf-8') as f:
            content = f.read()
        entries = content.split('\t')
        # == Get monday date ===
        date_monday = parse_french_date(entries[0])

        i = 0
        day_counter = 0
        while i + 1 < len(entries):
            i += 1

            entry = entries[i].strip()
            if entry:
                # print(f"Entry: {entry}, {i}")
                # Process the entry as needed
                m = RE_TIME.search(entry)
                if m:
                    time_range = (m.group('hours1'), m.group('minutes1') or '00', m.group('hours2'), m.group('minutes2') or '00')
                    # print(f"  Found time range: {time_range[0]}:{time_range[1]} to {time_range[2]}:{time_range[3]}")
                else:
                    print("  WARNING: No time range found.")
                    continue
                
                start_time = f"{time_range[0]}:{time_range[1]}"
                end_time = f"{time_range[2]}:{time_range[3]}"
                event_date = date_monday + datetime.timedelta(days=day_counter)
                description = entry
                events.append(Event(event_date, start_time, end_time, description))
                day_counter += 1
                if day_counter > 4:
                    day_counter = 0  # Reset for next week

    # for event in events:
    #     print(event)

    # Create iCalendar file
    cal = icalendar.Calendar()
    for event in events:
        ical_event = icalendar.Event()
        start_dt = datetime.datetime.combine(event.date, datetime.datetime.strptime(event.start_time, "%H:%M").time())
        end_dt = datetime.datetime.combine(event.date, datetime.datetime.strptime(event.end_time, "%H:%M").time())
        ical_event.add('dtstart', start_dt)
        ical_event.add('dtend', end_dt)
        ical_event.add('summary', event.description)
        cal.add_component(ical_event)

    with PATH_CAL_EXPORT.open('wb') as f:
        f.write(cal.to_ical())




if __name__ == '__main__':
    main()