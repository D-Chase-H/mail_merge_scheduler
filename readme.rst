===========================
Mail_Merge_Scheduler - Beta
===========================

Allows the user, to programmatically create a task for Windows Task Scheduler,
to take data from a database type that is supported by sqlalchemy, based on a
query, and perform a mail merge on a specified .docx or .dotx document. The
.docx document must have mail merge fields that correspond to the names of the
fields of the table that is being queried.

Database types supported:
::
 Should work with all database types supported by sqlalchemy:

   Confirmed:
     - Microsoft SQL Server
     - SQLLite

   Unconfirmed, but will most likely work:
     - MySQL
     - Oracle
     - PostgreSQL

   Unconfirmed, but may or may not work:
     - Firebird
     - Sybase

If you are using one of the unconfirmed database types, it would be very
helpful if you would let me know if it does or does not work for you.


Main Goal
=========

The main goal of this project is to automate documentation with data that is
already entered into a database. Despite the advances in data acquisition,
management, and storage today, many businesses and government organizations are
still required to keep documentation of certain things for record keeping or
legal liability purposes, especially documents that must be signed by an
employee.

This project aims to significantly reduce costs to these
organizations by reducing the time and paid work hours an employee spends
creating required documentation. In some cases, particularly with employees
who work in the field, this can enable an employee to do as much as twice the
work, while still maintaining all of their required documentation. This can
correspond to thousands or even tens of thousands of dollars saved in work
hours for a single employee in a year!

Example of Usage
================

An example would be an employee who gathers data in the field using a tablet.
When this employee returns from the field, they will often spend a notable
amount of time transferring data, that has already been gathered, into
documents and sign them.

However, with this library, one can set up a scheduled mail merge, with a query
that selects all records from a table that has a creation_date or date_modified
field that occurred a specified amount of time ago, for this example, let's say
that time is, 7 hours or less ago. Then the queried data will be taken from
the specified database and table at the scheduled time and a mail merge will be
performed a few minutes before the employee returns to the office at the end of
the day. All the employee has to do then, is review and sign their documents.
With some employees, like a building inspector working for a city, this can
eliminate 1-3 hours of documentation, as well as reducing possible errors
while creating documents.

This is just one example. This project can be used in any way your imagination
can come up with!
So feel free to be creative!

Prerequisites
=============

Requires Windows Vista or later.

Supports Python 3.x only.

Installation
============

Place the folder in the location of your choosing and do not delete or move any files in the folder.

pip install: Not yet, in-progress, and will release soon.


Dependencies
============

Third-party Python modules:
    - sqlalchemy
        * Home page: https://www.sqlalchemy.org/

        * Github: https://github.com/zzzeek/sqlalchemy/

        * Has pip install: YES, https://pypi.python.org/pypi/lxml

    - docx-mailmerge
        * GitHub: https://github.com/Bouke/docx-mailmerge

        * Has pip install: YES, https://pypi.python.org/pypi/docx-mailmerge/0.3.0

    - dateutil
        * Home page: https://dateutil.readthedocs.io/en/stable/

        * GitHub: https://dateutil.readthedocs.io/en/stable/

        * Has pip install: YES, https://pypi.python.org/pypi/python-dateutil

    - lxml
        * Home page: http://lxml.de/

        * GitHub: https://github.com/lxml/lxml

        * Has pip install: YES, https://pypi.python.org/pypi/lxml

============
============
Instructions
============
Step 1: Make An Instance of the ScheduledMerge Object
=====================================================
You will need 3 things.

1.) Your connection string to your database.

2.) The query you want to use.

3.) The path to your .docx document mail merge fields.
::
    from mail_merge_scheduler import ScheduledMerge

    database_connection_string = r"sqlite:///F:\\sql_lite\\MyDataBases\\testsqldb.db"
    query = "SELECT * FROM {} WHERE col1 < 10".format("table1")

    # Provide the full path to the template docx file.
    template_docx_path = r"F:\Scheduled Mail Merges\Daily Documentation\JohnDoe Mon-Fri 415pm\Inspections.docx"

    new_merge = ScheduledMerge(database_connection_string, query, template_docx_path, output_docx_name=None)

Optionally:
 You can pass in a name for the docx file created from the mail
 merge. If you leave it blank, the default naming convention will be
 "Merged" + the name of the template .docx document + a number In the above
 example, the docx file created from the mail merge would be named,
 Merged_Building Inspections_1.docx.
NOTE: 
 The mail merged .docx document is saved in the same folder as the template .docx document.

NOTE: 
 If there is already a Merged_Building Inspections_1.docx in that folder,
 then +1 will be added to the trailing number until it finds a file without
 that name.


Step 2: Set Up The Scheduling
=============================
Scheduling is done as a weekly scheduled task for the sake of simplicity.

* For DAILY schedules, simply set the week interval to  1, and pass in a list of all 7 days of the week.

* For Monthly schedules, simply set the week interval to a multiple of 4 and pass in your list of days.

::

    week_interval = 1
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    hour = 16
    minute = 15

    new_merge.set_schedule(week_interval, days, hour, minute, start_day=None)

Optionally:
 You can pass in a start_day for when you want Windows Task
 Scheduler to start running the scheduled task.

Step 3: Generate/Finalize
=========================
Finally, generate the scheduled mail merge and you're done!
::
    new_merge.generate_scheduled_merge()
========================================
========================================

Scheduled mail merges can be reviewed, edited, or even manually entered in the
schedules.ini file. Although you can manually enter in new scheduled mail
merges, it is recommended to use the mail_merge_scheduler.py module to add new
scheduled mail merges. This is because mail_merge_scheduler.py has error
checking for input to ensure everything will run smoothly.

The config file will look like this.
::
    [SCHEDULED_MERGES]
    Scheduled_Mail_Merge_for_Building
    Inspections.docx_at_16_15_every_1_week(s)_1 = [{'db_connection_string':
    'sqlite:///F:\\sql_lite\\MyDataBases\\testsqldb.db'}, {'db_query': 'SELECT
    * FROM table1 WHERE col1 < 10'}, {'template_docx_file_path': 'F:\Scheduled
    Mail Merges\Daily Documentation\JohnDoe Mon-Fri 415pm\Inspections.docx'},
    {'output_docx_name': None}, {'week_int': 1}, {'sched_days':
    ['2017-04-24 16:15:00', '2017-04-25 16:15:00', '2017-04-26
    16:15:00', '2017-04-27 16:15:00', '2017-04-28 16:15:00']}]
    

NOTE: 
 The script that will be called from Windows Task Scheduler is schedules.py, 
 which is not the same script that the user will use to set up a new scheduled mail merge.

Since sometimes files get deleted, the end-user may manually enter in incorrect
data to the schedules.ini file, or sometimes "things" may just change in
general, I have added a logging file called schedules.log. The above mention
script schedules.py will log any errors that may have occurred while running,
and write them to the log file for your review. This is especially nifty since
it provides the Traceback, enabling much faster debugging of potential errors!


License
=======
This project is licensed under the MIT License.

See the LICENSE file for details.
 link: https://github.com/D-Chase-H/mail_merge_scheduler/blob/master/LICENSE

Planned Features - Distant/Near Future
======================================
* Pip installation.
* A gui, which is already in the works.

Wish List
=========
* Open Office and Libre Office compatibility.

Credits
=======

This library was created by `Dustin Chase Harmon`. I go by Chase.

    * My LinkedIn: www.linkedin.com/in/dustinchaseharmon

    * My HackerRanks.com Profile: https://www.hackerrank.com/CHarmon

Contributing
============
Under normal circumstances I should get to pull requests within a few hours or by the next day.
This is my first repository, so bear with me if I can't get to your requests right away.

Please, send a pull request with your changes, and comments are appreciated.

Acknowledgments
===============

- I'd like to thank my brother Adam for encouraging me to teach myself
  programming.
- The programming community in general for providing a plethora of tutorials
  and information so anyone can teach themselves programming.
- The Google Foo.Bar challenge team and HackerRanks.com for showing me how fun
  and powerful programming can be!
- A tip of the hat to all the open source third-party libraries used in
  this project!
- Thank you to all those who contributed with pull requests!
