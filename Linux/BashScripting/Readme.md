
# Task 1: User management
A company has set up a new site and transfer staff and visitor accounts to the new site.
Your task is to write a Bash script to create user accounts for all staff and visitors. The supplied user file Usernames.txt is a text
file containing a username and its type delimited by comma per line. There are two types of users: staff and visitor. Staff users
are added to the staff group and visitor users to the visitors group.

## 1.
Write a Bash script, called createUsers.sh, to do the followings.
a) Create a group called visitors;

b) Create an account for each user and add the user to its group.
All user accounts are created with an initial password the same as their username; a home directory with the same
name as their username in the /home directory; all accounts use Bash shell program.

c) Write messages to syslog for all of the above events (new group, new user creation).

```
Note that while the current need is to handle limited number of usernames from the given user file, your script should be
able to handle an arbitrary number of usernames.
```

## 2.
Write a Bash script, called reportVisitors.sh, to report the members of visitors group to the file
/tmp/visitors.txt.

## 3.
Create a crontab entry to call the above script at 8:00AM and 9:00PM on every weekdays.

# Task 2: Web log analysis

## 1.
From the supplied (historical) NASA web server access log (NASA_access_log_Aug95.txt):
a) How many POST events appear in the log? Show the command you use to determine the number.

b) Write a one-line command to find the last nasa.gov GET request for the /history/skylab/flightsummary.txt
file. The output should list the fully qualified host name and its time of access. (Don’t worry about
removing any square brackets.)

c) Write a one-line command to compile and print all of the Expendable Launch Vehicle names listed in the log. The
ELV names are all listed in capital letters under the /elv directory. Your output should appear in the format of one
name per line, with no duplicates:

ALICE

BOB

CHARLIE

Save your output in a text file called elv.txt.
