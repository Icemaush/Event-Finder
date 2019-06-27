# Created by Reece Pieri, 2/6/19.

# Import modules required.
import wmi
from tkinter import *
from tkinter import ttk
from datetime import datetime
import time
import os


class EventLog(object):
    # ----- CREATE GUI WINDOW ---- #
    # Creates GUI window, frames and widgets.
    def __init__(self):
        # Main window and frame settings.
        self.window = Tk()
        self.window.title("Event Finder")
        frame = Frame(self.window)
        frame.pack()
        filter_frame = Frame(frame)
        filter_frame.pack()
        result_frame = Frame(frame)
        result_frame.pack()

        # Variables.
        self.status_text = StringVar()
        self.desktop = os.path.expanduser("~\Desktop\\")

        # Combobox value lists.
        eventtype_list = ["", "Error", "Warning", "Information", "Audit Success", "Audit Failure"]
        log_list = ["", "Application", "Security", "Setup", "System"]
        time_list = ["Any time", "Last hour", "Last 12 hours", "Last 24 hours", "Last 7 days", "Last 30 days"]

        # Label, combobox and button widgets.
        eventtype_lbl = Label(filter_frame, text="Event Type:")
        eventtype_lbl.grid(row=0, column=0, padx=5, pady=10)
        self.eventtype_cb = ttk.Combobox(filter_frame, state="readonly", values=eventtype_list)
        self.eventtype_cb.current(0)
        self.eventtype_cb.grid(row=0, column=1, padx=5, pady=10)
        log_lbl = Label(filter_frame, text="Logfile:")
        log_lbl.grid(row=1, column=0, padx=5)
        self.log_cb = ttk.Combobox(filter_frame, state="readonly", values=log_list)
        self.log_cb.current(0)
        self.log_cb.grid(row=1, column=1, padx=5)
        time_lbl = Label(filter_frame, text="Time Filter:")
        time_lbl.grid(row=2, column=0, padx=5, pady=10)
        self.time_cb = ttk.Combobox(filter_frame, state="readonly", values=time_list)
        self.time_cb.current(0)
        self.time_cb.grid(row=2, column=1, padx=5, pady=10)
        status_lbl = Label(filter_frame, textvariable=self.status_text)
        status_lbl.grid(row=3, column=0, columnspan=3, padx=5, pady=5)
        self.search_btn = Button(filter_frame, height=5, text="Search", command=self.search_log)
        self.search_btn.grid(row=0, column=2, rowspan=3, padx=5, pady=5)

        self.status_text.set(" ")

        mainloop()

    # ----- FORMATS TIME INTEGERS ----- #
    # Adds a "0" to single digit integers in time tuples for clean display when printing.  ----- #
    def time_format(self):
        self.timegen_formatted = []
        for i in self.timegen:
            if i < 10:
                i = str(i).zfill(2)
                self.timegen_formatted.append(i)
            else:
                self.timegen_formatted.append(i)

    # ----- VALIDATION FUNCTIONS ADJUST TIME INTEGERS BASED ON TIME FILTER SELECTED ----- #
    # Example: If date is January 1, 2020 and you select time filter "Last 24 hours", this will adjust day to 31, month
    # to 12 (December) and year to 2019.
    def validate_time_filter_24_hours(self):
        if self.timenow[2] - 1 <= 0:
            if self.timenow[1] == 1:
                self.timenow[0] = self.timenow[0] - 1
                self.timenow[1] = 12
                self.timenow[2] = 31
            if self.timenow[1] == 2:
                self.timenow[1] = 1
                self.timenow[2] = 31
            if self.timenow[1] == 3:
                self.timenow[1] = 2
                self.timenow[2] = 28
            if self.timenow[1] == 4:
                self.timenow[1] = 3
                self.timenow[2] = 31
            if self.timenow[1] == 5:
                self.timenow[1] = 4
                self.timenow[2] = 30
            if self.timenow[1] == 6:
                self.timenow[1] = 5
                self.timenow[2] = 31
            if self.timenow[1] == 7:
                self.timenow[1] = 6
                self.timenow[2] = 30
            if self.timenow[1] == 8:
                self.timenow[1] = 7
                self.timenow[2] = 31
            if self.timenow[1] == 9:
                self.timenow[1] = 8
                self.timenow[2] = 31
            if self.timenow[1] == 10:
                self.timenow[1] = 9
                self.timenow[2] = 30
            if self.timenow[1] == 11:
                self.timenow[1] = 10
                self.timenow[2] = 31
            if self.timenow[1] == 12:
                self.timenow[1] = 11
                self.timenow[2] = 30
        else:
            self.timenow[2] = self.timenow[2] - 1

    def validate_time_filter_7_days(self):
        if self.timenow[2] - 7 <= 0:
            if self.timenow[1] == 1:
                self.timenow[0] = self.timenow[0] - 1
                self.timenow[1] = 12
                self.timenow[2] = self.timenow[2] - 7 + 31
            if self.timenow[1] == 2:
                self.timenow[1] = 1
                self.timenow[2] = self.timenow[2] - 7 + 31
            if self.timenow[1] == 3:
                self.timenow[1] = 2
                self.timenow[2] = self.timenow[2] - 7 + 28
            if self.timenow[1] == 4:
                self.timenow[1] = 3
                self.timenow[2] = self.timenow[2] - 7 + 31
            if self.timenow[1] == 5:
                self.timenow[1] = 4
                self.timenow[2] = self.timenow[2] - 7 + 30
            if self.timenow[1] == 6:
                self.timenow[1] = 5
                self.timenow[2] = self.timenow[2] - 7 + 31
            if self.timenow[1] == 7:
                self.timenow[1] = 6
                self.timenow[2] = self.timenow[2] - 7 + 30
            if self.timenow[1] == 8:
                self.timenow[1] = 7
                self.timenow[2] = self.timenow[2] - 7 + 31
            if self.timenow[1] == 9:
                self.timenow[1] = 8
                self.timenow[2] = self.timenow[2] - 7 + 31
            if self.timenow[1] == 10:
                self.timenow[1] = 9
                self.timenow[2] = self.timenow[2] - 7 + 30
            if self.timenow[1] == 11:
                self.timenow[1] = 10
                self.timenow[2] = self.timenow[2] - 7 + 31
            if self.timenow[1] == 12:
                self.timenow[1] = 11
                self.timenow[2] = self.timenow[2] - 7 + 30
        else:
            self.timenow[2] = self.timenow[2] - 7

    def validate_time_filter_30_days(self):
        if self.timenow[2] - 30 <= 0:
            if self.timenow[1] == 1:
                self.timenow[0] = self.timenow[0] - 1
                self.timenow[1] = 12
                self.timenow[2] = self.timenow[2] - 30 + 31
            if self.timenow[1] == 2:
                self.timenow[1] = 1
                self.timenow[2] = self.timenow[2] - 30 + 31
            if self.timenow[1] == 3:
                if self.timenow[2] - 30 + 28 <= 0:
                    self.timenow[1] = 1
                    self.timenow[2] = self.timenow[2] - 30 + 31
                else:
                    self.timenow[1] = 2
                    self.timenow[2] = self.timenow[2] - 30 + 28
            if self.timenow[1] == 4:
                self.timenow[1] = 3
                self.timenow[2] = self.timenow[2] - 30 + 31
            if self.timenow[1] == 5:
                self.timenow[1] = 4
                self.timenow[2] = self.timenow[2] - 30 + 30
            if self.timenow[1] == 6:
                self.timenow[1] = 5
                self.timenow[2] = self.timenow[2] - 30 + 31
            if self.timenow[1] == 7:
                self.timenow[1] = 6
                self.timenow[2] = self.timenow[2] - 30 + 30
            if self.timenow[1] == 8:
                self.timenow[1] = 7
                self.timenow[2] = self.timenow[2] - 30 + 31
            if self.timenow[1] == 9:
                self.timenow[1] = 8
                self.timenow[2] = self.timenow[2] - 30 + 31
            if self.timenow[1] == 10:
                self.timenow[1] = 9
                self.timenow[2] = self.timenow[2] - 30 + 30
            if self.timenow[1] == 11:
                self.timenow[1] = 10
                self.timenow[2] = self.timenow[2] - 30 + 31
            if self.timenow[1] == 12:
                self.timenow[1] = 11
                self.timenow[2] = self.timenow[2] - 30 + 30
        else:
            self.timenow[2] = self.timenow[2] - 30

    # ----- GET LOCAL TIMEZONE OFFSET ----- #
    def get_timezone_offset(self):
        self.tz_offset = time.localtime()[3] - time.gmtime()[3]

    # ----- CONVERT EVENT VIEWER TIME ----- #
    # Converts event viewer time as all logs are written in GMT +0 offset but displayed using your current timezone
    # offset. This can be problematic when trying to search for events as it searches by the time written to the log
    # (GMT +0).
    # This function applies your current timezone offset and finds/displays events as you see them in the event viewer.
    def convert_eventlog_time(self):
        self.get_timezone_offset()
        if self.timegen[3] + self.tz_offset >= 24:
            self.timegen[3] = self.timegen[3] + self.tz_offset - 24
            if self.timegen[2] == 28 and self.timegen[1] == 2:
                self.timegen[2] = 1
                self.timegen[1] = 3
            if self.timegen[2] == 30:
                if self.timegen[1] == 4 or self.timegen[1] == 6 or self.timegen[1] == 9 or self.timegen[1] == 11:
                    self.timegen[2] = 1
                    self.timegen[1] = self.timegen[1] + 1
            if self.timegen[2] == 31:
                if self.timegen[1] == 1 or self.timegen[1] == 3 or self.timegen[1] == 5 or self.timegen[1] == 7 or \
                        self.timegen[1] == 8 or self.timegen[1] == 10 or self.timegen[1] == 12:
                    self.timegen[2] = 1
                    if self.timegen[1] != 12:
                        self.timegen[1] = self.timegen[1] + 1
                    else:
                        self.timegen[1] = 1
                        self.timegen[0] = self.timegen[0] + 1
        else:
            self.timegen[3] = self.timegen[3] + self.tz_offset

    # ----- PRINT EVENTS FOUND TO CONSOLE ----- #
    def print_to_console(self):
        print("Record Number: " + str(self.log.RecordNumber))
        print("Event Type: " + self.eventtype_cb.get())
        print("Time Generated: " + self.timegenerated)
        print("Source: " + self.log.SourceName)
        print("Event ID: " + str(self.log.EventIdentifier))
        print("\n")

    # ----- WRITE EVENTS FOUND TO TEXT FILE ----- #
    def write_to_file(self):
        recnum = "Record Number: " + str(self.log.RecordNumber) + "\n"
        evnttype = "Event Type: " + "Error\n"
        timegen = "Time Generated: " + self.timegenerated + "\n"
        source = "Source: " + self.log.SourceName + "\n"
        evntid = "Event ID: " + str(self.log.EventIdentifier) + "\n\n"
        message = str(self.log.Message) + "\n\n"
        break_txt = "---------------------------------------\n"
        textlist = [recnum, evnttype, timegen, source, evntid,
                    message, break_txt]
        self.file.writelines(textlist)

    # ----- MAIN SEARCH FUNCTION ----- #
    def search_log(self):
        if self.eventtype_cb.get() == "" or self.log_cb.get() == "":
            if self.eventtype_cb.get() == "":
                self.status_text.set("Please select an Event Type.")
            elif self.log_cb.get() == "":
                self.status_text.set("Please select a Logfile.")
        else:
            # Set status message.
            self.status_text.set("Searching. This may take a while.")
            self.window.update_idletasks()
            print("Searching. This may take a while.")

            # Gets selected filters.
            eventtype = self.eventtype_cb.current()
            eventlog = self.log_cb.get()

            # Prints selected filter information.
            print("---------------------------------------")
            print("Event Type: \t" + self.eventtype_cb.get())
            print("Logfile: \t" + eventlog)
            print("Time Filter: \t" + self.time_cb.get())
            print("---------------------------------------\n")

            # Creates new .txt file to record events.
            self.file = open(os.path.join(self.desktop, "Event Log.txt"), "w", encoding="utf-8")
            break_txt = "---------------------------------------\n"
            eventtype_txt = "Event Type: \t" + self.eventtype_cb.get() + "\n"
            logfile_txt = "Logfile: \t" + eventlog + "\n"
            timefilter_txt = "Time Filter: \t" + self.time_cb.get() + "\n"
            break2_txt = "---------------------------------------\n\n"
            textlist = [break_txt, eventtype_txt, logfile_txt,
                        timefilter_txt, break2_txt]
            self.file.writelines(textlist)

            # PC to connect to.
            c = wmi.WMI()

            # Run through events and filter based on filter selections.
            count = 0
            for self.log in c.Win32_NTLogEvent(EventType=eventtype, Logfile=eventlog):
                self.timegen = list(wmi.to_time(self.log.timegenerated))
                self.timegen.remove(self.timegen[7])
                self.timegen.remove(self.timegen[6])
                self.convert_eventlog_time()
                self.time_format()
                now = datetime.now()
                self.timenow = [now.year, now.month, now.day, now.hour, now.minute, now.second]
                self.timegenerated = str(self.timegen_formatted[2]) + "/" + str(self.timegen_formatted[1]) + "/" + \
                                     str(self.timegen_formatted[0]) + " " + str(self.timegen_formatted[3]) + ":" + \
                                     str(self.timegen_formatted[4]) + ":" + str(self.timegen_formatted[5])

                # IF NO TIME FILTER SELECTED #
                if self.time_cb.get() == "":
                    count += 1
                    if count != 0:
                        self.print_to_console()
                        self.write_to_file()

                # IF TIME FILTER SELECTED = 1 HOUR #
                if self.time_cb.get() == "Last hour":
                    if self.timenow[3] - 1 < 0:
                        self.timenow[3] = 23
                        self.validate_time_filter_24_hours()
                    else:
                        self.timenow[3] = self.timenow[3] - 1

                    if self.timegen >= self.timenow:
                        count += 1
                        if count != 0:
                            self.print_to_console()
                            self.write_to_file()

                # IF TIME FILTER SELECTED = 12 HOURS #
                if self.time_cb.get() == "Last 12 hours":
                    if self.timenow[3] - 12 < 0:
                        self.timenow[3] = 24 + (self.timenow[3] - 12)
                        self.validate_time_filter_24_hours()
                    else:
                        self.timenow[3] = self.timenow[3] - 12

                    if self.timegen >= self.timenow:
                        count += 1
                        if count != 0:
                            self.print_to_console()
                            self.write_to_file()

                # IF TIME FILTER SELECTED = 24 HOURS #
                if self.time_cb.get() == "Last 24 hours":
                    self.validate_time_filter_24_hours()
                    if self.timegen >= self.timenow:
                        count += 1
                        if count != 0:
                            self.print_to_console()
                            self.write_to_file()

                # IF TIME FILTER SELECTED = 7 DAYS #
                if self.time_cb.get() == "Last 7 days":
                    self.validate_time_filter_7_days()
                    if self.timegen >= self.timenow:
                        count += 1
                        if count != 0:
                            self.print_to_console()
                            self.write_to_file()

                # IF TIME FILTER SELECTED = 30 DAYS #
                if self.time_cb.get() == "Last 30 days":
                    self.validate_time_filter_30_days()
                    if self.timegen >= self.timenow:
                        count += 1
                        if count != 0:
                            self.print_to_console()
                            self.write_to_file()

            self.file.close()
            if count == 0:
                os.remove(self.desktop + "Event Log.txt")
            else:
                os.startfile(self.desktop + "Event Log.txt")
            if count == 1:
                self.status_text.set("Search complete. " + str(count) + " event found.")
                print("Search complete. " + str(count) + " event found.")
            else:
                self.status_text.set("Search complete. " + str(count) + " events found.")
                print("Search complete. " + str(count) + " events found.")


EventLog()
