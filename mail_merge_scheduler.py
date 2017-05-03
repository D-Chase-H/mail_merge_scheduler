
# This module is part of mail_merge_scheduler and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""This is the module that the user will import and uses to set up a scheduled
mail merges, with data taken from database types supported by sqlalchemy. It
will write the information to the schedules.ini file and uses the xml template
to schedule the task through Windows Task Scheduler. The user can also use this
module to remove scheduled mail merges from both the schedules.ini file and the
Windows Task Scheduler, provided the name of the task, and the key in the
schedules.ini file have not been altered.
"""

# Standard library imports
import configparser
from datetime import time
from datetime import date
from datetime import datetime
from datetime import timedelta
import os
import sys
import subprocess
# Third-party imports
from lxml import etree
import sqlalchemy



def remove_scheduled_merge(scheduled_merge_key):
    """Removes a scheduled mail merge.

    This function takes the key from an entry in the schedules.ini files,
    and removes it from the ini file, and deletes the task from the
    Windows Task Scheduler.

    Args:
        scheduled_merge_key: A string of a key from a scheduled merge in the
            schedules.ini file.

    Returns:
        None.

    Raises:
        KeyError: scheduled_merge_key.
    """

    config = configparser.ConfigParser()
    config.optionxform = str
    config.read("scheduled_merges.ini")
    sect = "SCHEDULED_MERGES"
    del config[sect][scheduled_merge_key]

    with open('scheduled_merges.ini', 'w') as config_file:
        config.write(config_file)
        config_file.close()

    # subprocess batch script to Windows Task Scheduler, to delete the task
    # whose name correlates to the key id in the schedules.ini file.
    task = "schtasks.exe /delete /tn {} /f".format(scheduled_merge_key)
    process = subprocess.Popen(task, shell=False)
    process.wait()
    return



class ScheduledMerge(object):
    """This is the class that the end-user will use to set up a scheduled mail
    merge with data taken from a database.

    Should work for all database types that are supported by sqlalchemy.

    Usage of this class requires 3 steps.
    Here is an example:

        from mail_merge_scheduler import ScheduledMerge

        new_merge = ScheduledMerge(connection_string, query, template_docx_path)
        new_merge.set_scheudle(week_interval, days, hour, minute)
        new_merge.generate_scheduled_merge()

    Attributes:
        db_connection_string: A connection string that can be used to connect
            to the desired database. Instructions can be found on sqlalchemy's
            website at: http://docs.sqlalchemy.org/en/latest/core/engines.html
        db_query: A string of a database query that will be used to take data
            from the database that the db_connection_string attribute connects
            with.
        template_docx_file_path: A string  giving the full path to a .docx or
            .dotx document that has mail merge fields with names that correlate
            to the field names of the table in the db_query attribute.
        output_docx_name: OPTIONAL A string of the file name only, and not the
            full path, of the the .docx document created from the mail merge.
            This is a new document, and does not overwrite the
            template_docx_file_path.
        start_day: A datetime.date object that specifies the day the user wants
            the scheduled mail merge task to begin.
        week_int: An integer indicating the week interval/frequency for the
            scheduled mail merge task to occur.
        sched_days: A list of strings indicating the days of the week that the
            user would like the scheduled mail merge task to occur on.
        sched_time: A datetime.time object indicating the hour and minute of
            the day the user wants the scheduled mail merge task to occur.
    """

    # pylint: disable=too-many-instance-attributes
    def __init__(
            self, db_connection_string, db_query,
            template_docx_file_path, output_docx_name=None):
        """Init for ScheduledMerge
        Raises:
            sqlalchemy.exc.OperationalError: no such table.
            FileNotFoundError: template_docx_file_path.
        """

        # Check if the given connection string and query are valid.
        # If invalid, sqlalchemy will raise an error.
        engine = sqlalchemy.create_engine(db_connection_string)
        engine = engine.connect()
        engine.execute(db_query)
        engine.close()

        # Check if the docx file path is valid.
        if os.path.isfile(template_docx_file_path) is False:
            raise FileNotFoundError(template_docx_file_path)

        ## Database Information
        self.db_connection_string = db_connection_string
        self.db_query = db_query

        ## Path Information
        self.template_docx_file_path = template_docx_file_path
        # Output docx is saved to the same folder as the template docx file.
        if output_docx_name is None:
            self.output_docx_name = None
        else:
            self.output_docx_name = output_docx_name

        ## Scheduling Information.
        self.start_day = None
        self.week_int = None
        self.sched_days = []
        self.sched_time = None


    # pylint: disable=too-many-arguments
    def set_scheudle(self, week_interval, days, hour, minute, start_day=None):
        """Allows the end-user to set the attributes for scheduling.

        Args:
            week_interval: An integer indicating the week interval/frequency
                for the scheduled mail merge task to occur.
            days: A list of strings indicating the days of the week that the
                user would like the scheduled mail merge task to occur on. Each
                day must be in a specific format:
                1.) The first letter must be uppercase.
                2.) The the following letters must be lowercase.
                3.) The full name of the day must be used. No abbreviations.
                Example: ['Monday', 'Wednesday', 'Friday']
            hour: An integer between 0 and 23 (inclusive), indicating the hour,
                in military time, the user wants the scheduled mail merge to
                occur.
            minute: An integer between 0 and 59 (inclusive), indicating the
                minute, in military time, the user wants the scheduled mail
                merge to occur.
            start_day: OPTIONAL A datetime.date object that specifies the day
                the user wants the scheduled mail merge task to begin.

        Returns:
            None.

        Raises:
            AssertionError: Scheduling Error: Cannot schedule a mail merge with
                an empty list of days.
            AssertionError: Scheduling Error: List of days contains duplicates.
            AssertionError: Scheduling Error: Incorrect format for a day(s) in
                your list of days.
                "The required format for a day has:
                "1.) The full name of the day.
                "2.) The first letter is uppercase.
                "3.) The following letters are all lowercase.
                "Example: ['Monday', 'Wednesday', 'Friday'].
        """

        if start_day is None:
            self.start_day = date.today()
        else:
            self.start_day = start_day

        self.week_int = int(week_interval)
        self.sched_time = time(hour=int(hour), minute=int(minute))

        assert days, \
        ("\nScheduling Error: "
         "Cannot schedule a mail merge with an empty list of days.")

        assert len(days) == len(set(days)), \
        ("\nScheduling Error: "
         "List of days contains duplicates.")

        # days_format_checker and check_days ensures the days are in the proper
        # format for use in the Windows Task Scheduler xml schema.
        days_format_checker = set([
            "Monday", "Tuesday", "Wednesday", "Thursday",
            "Friday", "Saturday", "Sunday"])
        check_days = [item for item in days if item in days_format_checker]

        assert days == check_days, \
        ("\nScheduling Error:\n"
         "Incorrect format for a day(s) in your list of days.\n"
         "The required format for a day has:\n"
         "1.) The full name of the day.\n"
         "2.) The first letter is uppercase.\n"
         "3.) The following letters are all lowercase.\n"
         "Example: ['Monday', 'Wednesday', 'Friday'].")

        self.sched_days = self.generate_list_of_next_days(days)

        # Check for errors in other attributes. This was seperated because it
        # is used in other areas, and to maintain readability.
        self.error_check_attributes()
        return


    def error_check_attributes(self):
        """Checks to see if all attributes are filled with correct data.

        Returns:
            None.

        Raises:
            AssertionError: Scheduling Error: start_day must be a datetime.date
                object.
            AssertionError: Scheduling Error: Week interval must be an integer
                of 1 or above.
            AssertionError: Scheduling Error: There was an error with hour or
                minute input. Hour must be between 0 and 23, inclusive. Minute
                must be between 0 and 59, inclusive.
        """

        assert isinstance(self.start_day, date), \
        ("\nScheduling Error: "
         "start_day must be a datetime.date object.")

        assert isinstance(self.week_int, int) and self.week_int >= 1, \
        ("\nScheduling Error: "
         "Week interval must be an integer of 1 or above.")

        assert isinstance(self.sched_time, time), \
        ("\nScheduling Error: "
         "There was an error with hour or minute input."
         "Hour must be between 0 and 23, inclusive."
         "Minute must be between 0 and 59, inclusive.")

        # Ensure each day in sched_days is a datetime.datetime object.
        sched_days_format_check = True
        for day in self.sched_days:
            if isinstance(day, datetime) is False:
                sched_days_format_check = False
                break

        assert sched_days_format_check is True, \
        ("\nScheduling Error: "
         "One or more of the items in the list of self.sched_days is not a"
         "datetime.datetime object.")
        return


    def load_data_into_list_of_dicts(self):
        """Takes the instance attributes of the class, and converts them into
        an appropriate format for storing in the schedules.ini file.

        Returns:
            A list of dictionaries of attribute names and thier data. This
            format was chosen so the data stays in the same place in the
            schedules.ini file, so it is easier to read.
        """

        # Create a list of dictionaries instead of a single dictionary, so the
        # schedules.ini file has better human readability.
        days = [str(d) for d in self.sched_days]
        list_of_dicts = [
            {"db_connection_string":self.db_connection_string},
            {"db_query":self.db_query},
            {"template_docx_file_path":self.template_docx_file_path},
            {"output_docx_name":self.output_docx_name},
            {"week_int":self.week_int},
            {"sched_days":days}]

        return list_of_dicts


    def generate_unique_config_key_id(self):
        """Creates a unique string to use as a key in the schedules.ini file.

        Returns:
            A unique string that is used as the key in the schedules.ini file,
            as well as for the task name in the Windows Task Scheduler.
        """

        config = configparser.ConfigParser()
        config.optionxform = str
        config.read("scheduled_merges.ini")
        sect = "SCHEDULED_MERGES"

        # Get all keys from the .ini file, so a duplicate key is not made.
        config_keys_set = set([i for i in config[sect]])

        prefix = "Scheduled_Mail_Merge_for"
        docx_name = os.path.basename(self.template_docx_file_path)
        id_num = 1
        hour = self.sched_time.hour
        minute = self.sched_time.minute
        week_int = self.week_int

        config_key_front = "{}_{}_at_{}_{}_every_{}_week(s)".format(
            prefix, docx_name, hour, minute, week_int)
        config_key = "{}_{}".format(config_key_front, id_num)

        # If the key already exists, keep adding 1 to id_num, until a unique
        # key is found.
        while True:
            if config_key in config_keys_set:
                id_num += 1
                config_key = "{}_{}".format(config_key_front, id_num)
            else:
                break
        return config_key


    # pylint: disable=no-self-use
    def import_task_to_win_task_sched(self, xml_name, task_name):
        """Imports an xml into Windows Task Scheduler via a subprocess.

        Args:
            xml_name: A string of the file name of xml that was generated, to
                be imported into Windows Task Scheduler.
            task_name: A string to be used as the task name for the task in
                Windows Task Scheduler.

        Returns:
            None.
        """

        task = "schtasks.exe /create /XML {} /tn {}".format(xml_name, task_name)
        process = subprocess.Popen(task, shell=False)
        process.wait()
        return


    # pylint: disable=too-many-locals
    def generate_list_of_next_days(self, days):
        """Given the list of days, this method finds when the next occurance of
        those days of the week, and converts them into datetime objects and
        returns a list of datetime objects.

        Args:
            days: A list of strings indicating the days of the week that the
                user would like the scheduled mail merge task to occur on.

        Returns:
            A list of datetime.datetime objects.
        """

        # A Dictionary for converting the number that is returned from
        # datetime.weekday(), into the string representation of the day's name.
        weekday_conv = {
            "Monday":0, "Tuesday":1, "Wednesday":2, "Thursday":3,
            "Friday":4, "Saturday":5, "Sunday":6}

        start_hour = self.sched_time.hour
        start_minute = self.sched_time.minute
        start_day = self.start_day
        start_day_num = self.start_day.weekday()
        datetime_days = []

        for day in days:
            day_num = weekday_conv[day]
            if start_day_num > day_num:
                next_day_num = 6 - (start_day_num - abs(start_day_num -
                                                        day_num))

            elif start_day_num < day_num:
                next_day_num = 7 - (start_day_num + abs(start_day_num -
                                                        day_num))

            elif start_day_num == day_num:
                now = datetime.today().time()
                curr_hour = now.hour
                curr_minute = now.minute
                if curr_hour <= start_hour and curr_minute < start_minute:
                    next_day_num = 0
                else:
                    next_day_num = 7

            start_day_copy = start_day
            start_day_copy += timedelta(days=next_day_num)
            start_day_copy = str(start_day_copy)
            year, month, day = map(int, start_day_copy.split("-"))

            next_day = datetime(
                year=year, month=month, day=day,
                hour=start_hour, minute=start_minute)

            datetime_days.append(next_day)

        return datetime_days


    # pylint: disable=too-many-locals
    # pylint: disable=too-many-branches
    # pylint: disable=too-many-statements
    # pylint: disable=no-member
    # pylint: disable=anomalous-backslash-in-string
    def xml_gen(self, task_name):
        """Uses the weekly schedule xml template for Windows Task Scheduler,
        and replaces specific data in the xml tree, then writes the xml as
        out_xml.xml, imports the xml into Windows Task Scheduler via a
        subprocess batch script, then deletes out_xml.xml.

        Args:
            task_name: A string to be used as the task name for the task in
                Windows Task Scheduler.

        Returns:
            None.

        Raises:
            If the weekly_xml_schedule_template.xml is edited, it may raise
                errors pertaining to the Windows Task Scheduler xml schema. If
                you would like to edit this xml, you can find information on
                how to do so here:
                https://msdn.microsoft.com/en-us/library/windows/desktop/aa38360
                9(v=vs.85).aspx .
        """

        weekday_conv = {
            0:"Monday", 1:"Tuesday", 2:"Wednesday", 3:"Thursday",
            4:"Friday", 5:"Saturday", 6:"Sunday"}

        next_date = self.start_day
        next_time = self.sched_time
        week_int = self.week_int
        days = self.sched_days

        days = [weekday_conv[d.weekday()] for d in days]

        xml_template = "weekly_xml_schedule_template.xml"

        dom = os.getenv("USERDOMAIN")
        user = os.getenv("USERNAME")
        dn_un = "{}\{}".format(dom, user)

        today = datetime.today().date()
        curr_time = datetime.today().time()
        creation_datetime = "{}T{}".format(today, curr_time)

        next_date = "{}T{}".format(next_date, next_time)

        # Tries to find the windowless version of the version of python being
        # used to run the script. The search is done using typical naming
        # conventions for python, but if it cannot be found, then the normal
        # version of python will be used.
        pyth_path = '\\'.join(sys.executable.split('\\')[0:-1])
        files_in_pyth_path = os.listdir(pyth_path)
        python_variations = ["python", "Python", "PYTHON"]
        windowless_pyth_path = ""

        for file_name in files_in_pyth_path:
            for var in  python_variations:
                if var in file_name:
                    if "w" in file_name or "W" in file_name:
                        windowless_pyth_path = "{}\\{}".format(
                            pyth_path, file_name)

        # If the windowless python executable couldnt be found, then use the
        # windowed executable that is being used to run the script.
        if not windowless_pyth_path:
            windowless_pyth_path = sys.executable

        # name of the script that Windows Task Scheduler will run.
        script_name = "schedules.py"

        curr_dir = os.getcwd()
        curr_dir = r"{}".format(curr_dir)

        tree = etree.parse(xml_template)

        for ele in tree.iter():
            # The prefix is set to
            # http://schemas.microsoft.com/windows/2004/02/mit/task
            # This should work on all versions of Windows newer than, and
            # including, Windows Vista.
            prefix = str(ele.tag)[:55]
            tag_name = str(ele.tag)[55:]

            if tag_name == "Date":
                ele.text = creation_datetime

            if tag_name == "Author":
                ele.text = dn_un

            if tag_name == "Description":
                ele.text = "Scheduled Mail Merge"

            if tag_name == "StartBoundary":
                ele.text = next_date

            if tag_name == "WeeksInterval":
                ele.text = str(week_int)

            if tag_name == "DaysOfWeek":
                for day in days:
                    etree.SubElement(ele, "{}{}".format(prefix, day))

            if tag_name == "UserId":
                ele.text = dn_un

            if tag_name == "Command":
                ele.text = r"{}".format(windowless_pyth_path)

            if tag_name == "Arguments":
                ele.text = script_name

            if tag_name == "WorkingDirectory":
                ele.text = curr_dir

        xml_name = "out_xml.xml"
        tree.write(xml_name, xml_declaration=None, encoding='UTF-16')

        # Import the generated xml into Windows task Scheduler via a subprocess
        # batch script.
        self.import_task_to_win_task_sched(xml_name, task_name)

        # Remove the generated xml after it has been imported.
        os.remove("out_xml.xml")
        return


    # pylint: disable=too-many-locals
    def generate_scheduled_merge(self):
        """Finalizes the scheduled mail merge.

        Writes a dictionary of the intance attributes to the schedules.ini
        file, and generates an xml from the template xml and imports the xml
        into Windows Task Scheduler.

        Returns:
            None.

        Raises:
            All errors raised by the method: error_check_attributes(self)
        """

        # Do a quick check for errors before generating a scheduled mail merge.
        self.error_check_attributes()

        # Make a list of dictionaries to be written to the schedules.ini file.
        list_of_dicts_of_merge_data = self.load_data_into_list_of_dicts()
        config = configparser.ConfigParser()

        # optionxform maintains upercase letters in strings for keys.
        config.optionxform = str
        config.read("scheduled_merges.ini")
        sect = "SCHEDULED_MERGES"

        # Shorten the config_key_id string if the length exceedes the maximum
        # length for a task name in Windows Task Scheduler of 232 characters.
        config_key_id = self.generate_unique_config_key_id()
        config_key_id = config_key_id[:232]
        config[sect][config_key_id] = str(list_of_dicts_of_merge_data)

        with open('scheduled_merges.ini', 'w') as config_file:
            config.write(config_file)
            config_file.close()

        # Create an xml, from the template xml, that will be used to import
        # into Windows Task Scheduler.
        self.xml_gen(str(config_key_id))
        return
