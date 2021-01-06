# MeetingAnalysis
This is a python script to analyse zoom webinar data

Webinars have become the main mode of communication and information dissemination. 
How would you analyse webinars are effective, one evaluating the delivery, other audience measurement. I focused more on the audience measurement to see if there are any insights.

Have been measuring for almost 20+ webinars.
Few measurements,
1. metric_total_participation_quality : Plot of total duration attended and number of people
2. metric_join_leave_window : Plot of time at which a person joined and left the meeting. Are there random participants.
3. metric_attendance_ratio : No. of people attended vs Total number of people registered.
4. metric_participant_count_byminute : Count of people present in meeting by time (in minutes)

This measurements helped to observe impact of other business decisions & communications to client prospects.

This currently works for zoom data. Should be working for other meeting softwares if the expected data is available. I haven’t checked it though.

Warning: Do not become obsessed with this measurement. It is only a reference of what happened in meeting. Sample output graphs was measured in meetings run for public participants. It shows lot of randomness. 
Might be a different case if you run this for an office meeting or college/university meeting.

sample input: webinar.xls (Sheet - Attendee)
sample output: jpg

Some features to implement in future: 
1)Generate PDF report 
  generate_report()
2)Build a graphical dashboard
