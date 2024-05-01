# Sync CalMoodle-calendar with personal calendar

This is a workaround for the problem with having a teaching calendar running from CalMoodle that does not synchronize and auto-update your personal calendar in Outlook. This is mainly made for employees at Aalborg University. The following script(s) does two things:

1.  It checks all paths needed to automate sync
2. It evaluates CalMoodle-objects, adds them to your calendar and upon re-run, it moves all elements that has moved in CalMoodle since last run to their current location

This script is not fool-proof. You will need to manually run it at given intervals because it does not auto-run at CalMoodle changes. Due to AAU security policies, automation has proved difficult but will hopefully be added in the future.