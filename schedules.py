
# This module is part of mail_merge_scheduler and is released under
# the MIT License: http://www.opensource.org/licenses/mit-license.php

"""This script is run through Windows Task Scheduler at the specified time by
the end-user by using the mail_merge_scheduler module or by manually entering
data into the scheduled_merges.ini file and manually creating a task in Windows
Task Scheduler. It is reccommended to always set up scheduled mail merges
through the mail_merge_scheduler.

This script is NOT meant to be used by the end-user.
"""

# Standard library imports
import ast
import configparser
from datetime import datetime
from datetime import timedelta
import logging
import os
# Third-party imports
from dateutil.parser import parse
from mailmerge import MailMerge
import sqlalchemy



def create_logger():
    """Creates and returns a logger. Errors are logged in schedules.log."""
    module_path = __file__
    path = os.path.split(module_path)[0]
    log_path = r"{}\schedules.log".format(path)

    new_logger = logging.getLogger(__name__)
    new_logger.setLevel(logging.ERROR)

    seperator = ["_"]*80
    seperator = ''.join(seperator)
    logger_format = "{}\n%(asctime)s\n%(message)s\n".format(seperator)
    formatter = logging.Formatter(logger_format)

    file_handler = logging.FileHandler(log_path)
    file_handler.setLevel(logging.ERROR)
    file_handler.setFormatter(formatter)

    new_logger.addHandler(file_handler)
    return new_logger



# pylint: disable=invalid-name
logger = create_logger()



class ScheduledMerge(object):
    """Object containing methods for loading and processing data found in the
    schedules_merges.ini file.

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
        output_docx_name: A string of the file name only, and not the
            full path, of the the .docx document created from the mail merge.
            This is a new document, and does not overwrite the
            template_docx_file_path.
        week_int: An integer indicating the week interval/frequency for the
            scheduled mail merge task to occur.
        sched_days: A list of strings indicating the days of the week that the
            user would like the scheduled mail merge task to occur on.
    """


    # pylint: disable=too-many-arguments
    def __init__(self, db_connection_string, db_query, template_docx_file_path,
                 output_docx_name, week_int, sched_days):

        ## Database Information
        self.db_connection_string = db_connection_string
        self.db_query = db_query

        ## Path Information
        self.template_docx_file_path = template_docx_file_path
        self.output_docx_name = output_docx_name

        ## Scheduling Information
        self.week_int = week_int
        self.sched_days = sched_days

    def generate_out_filename(self):
        """Creates a unique file name for the output docx file, that is used in
        the perform_mail_merge method below.

        Returns:
            A string for a unique name for the .docx document created from the
            mail merge.
        """

        head, tail = os.path.split(self.template_docx_file_path)
        out_docx_path = r"{}\Merged_{}".format(head, tail)

        if os.path.isfile(out_docx_path):
            file_name, ext = os.path.splitext(tail)
            num = 1

            # If the file name already exists, keep adding 1 to num, until a
            # unique file name is found.
            while True:
                out_docx_path = r"{}\Merged_{}_{}{}".format(
                    head, file_name, num, ext)

                if not os.path.isfile(out_docx_path):
                    break
                num += 1

        return out_docx_path


    def create_dict_of_data_from_vars(self):
        """Writes updated data back to the schedules_merges.ini config file.

        If any data is updated, this method converts all the instance
        attributes into a dictionary so the list can be used to
        overwrite the data for this object in the schedules_merges.ini file.

        Returns:
            dict_of_data: A dictionary with keys that correlate with
                the names of the class attributes and values of the data for
                that attribute.
        """

        days = [str(d) for d in self.sched_days]
        dict_of_data = {
            "db_connection_string":self.db_connection_string,
            "db_query":self.db_query,
            "template_docx_file_path":self.template_docx_file_path,
            "output_docx_name":self.output_docx_name,
            "week_int":self.week_int,
            "sched_days":days}
        return dict_of_data


    def compare_time_to_sched_days(self):
        """Checks self.sched_days to see if a mail merge needs to be performed.

        Compares the current datetime to all datetimes in the sched_days
        instance attribute, to check if the there is a merge scheduled for
        right now. Also in cases where the user was not logged on when a mail
         merge task was scheduled, this method will also check if any datetimes
        in sched_days occur before the current datetime, and if it does, it
        will return True, and therefore perform the mail merge.

        Returns:
            Boolean True or False.
        """

        today = datetime.today()

        for index, day in enumerate(self.sched_days):
            time_diff = (today - day).days

            if time_diff >= 0:
                self.sched_days[index] = self.update_day(day)
                return True
            else:
                continue

        return False


    def update_day(self, day):
        """Updates a datetime in sched_days by timedelta-ing it by the week
        interval, if a mail merge was performed.

        Args:
            day: A datetime.datetime object
        Returns:
            A datetime.datetime object
        """

        day += timedelta(weeks=self.week_int)
        return day


    def get_records_from_db(self):
        """Gets data from the database, based on the query given, and returns
        the data.

        Returns:
            A list of dictionaries, with key:value pairs arranged as
            key=field_name : value=record for that row.
        """

        eng = sqlalchemy.create_engine(self.db_connection_string)
        eng.connect()
        rows = eng.execute(self.db_query)
        flds = rows.keys()
        records = []

        for row in rows:
            rec = {str(fld):str(row[ind]) for ind, fld in enumerate(flds)}
            records.append(rec)

        return records


    def perform_mail_merge(self):
        """Performs a mail merge and creates a new docx file."""

        in_docx_path = self.template_docx_file_path
        if self.output_docx_name is None:
            out_docx_path = self.generate_out_filename()
        data = self.get_records_from_db()

        document = MailMerge(in_docx_path)
        document.merge_pages(data)
        document.write(out_docx_path)
        return


def write_dict_to_config(config_path, config, config_key_id, dict_of_data):
    """Writes data as a list of dictionaries back to the schedules_merges.ini
    file.
    """

    for key, value in dict_of_data.items():
        if isinstance(value, str):
            value = 'r"{}"'.format(value)
        config[config_key_id][key] = str(value)

    with open(config_path, 'w') as config_file:
        config.write(config_file)
        config_file.close()
    return

# pylint: disable=too-many-locals
# pylint: disable=broad-except
# pylint: disable=unused-variable
def check_for_scheduled_merges():
    """This function runs when the script is run.

    Loads and iterates through the dictionaries in the schedules_merges.ini
    file, and checks if there is a mail merge scheduled for right now, if so, it
    performs the mail merge and updates the datetimes in the list of
    sched_days, and overwrites that dictionary with the updated data.

    Returns:
        None
    Raises:
        All Errors raised will be written to the schedules.log file.
    """

    module_path = __file__
    path = os.path.split(module_path)[0]
    config_path = r"{}\scheduled_merges.ini".format(path)
    config = configparser.ConfigParser()

    # optionxform maintains upercase letters in strings for keys.
    config.optionxform = str
    config.read(config_path)

    try:
        if os.path.isfile(config_path) is False:
            raise FileNotFoundError(config_path)
    # If the scheduled_merges.ini file has been deleted, for whatever reason,
    # this will re-create the file.
    except Exception as exception:
        logger.exception("")
        with open(config_path, 'w') as new_config_file:
            config.write(new_config_file)
            new_config_file.close()

    for key in config.sections():
        try:
            dict_of_data = [ast.literal_eval(v) for v in config[key].values()]
            db_connection_string = dict_of_data[0]
            db_query = dict_of_data[1]
            template_docx_file_path = dict_of_data[2]
            output_docx_name = dict_of_data[3]
            week_int = dict_of_data[4]
            sched_days = [parse(item) for item in dict_of_data[5]]

            new_mer_obj = ScheduledMerge(
                db_connection_string,
                db_query,
                template_docx_file_path,
                output_docx_name,
                week_int,
                sched_days)

            # do_merge returns a boolean to determine if a merge should be done.
            do_merge = new_mer_obj.compare_time_to_sched_days()

            if do_merge is True:
                new_mer_obj.perform_mail_merge()
                dict_of_data = new_mer_obj.create_dict_of_data_from_vars()
                write_dict_to_config(
                    config_path, config, key, dict_of_data)

        # Use the general, Exception as exception, so that any error
        # that occurs has its Traceback written to the schedules.log
        # file.
        except Exception as exception:
            logger.exception("KEY_ID: %s", key)

    return



if __name__ == "__main__":
    check_for_scheduled_merges()
