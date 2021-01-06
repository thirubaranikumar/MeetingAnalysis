'''
This application provides zoom webinar meeting analytics
Following 4 metrics are currently available
1. metric_total_participation_quality
2. metric_join_leave_window
3. metric_attendance_ratio
4. metric_participant_count_byminute

Author : Barani Kumar

License : GPL3 License
'''

import pandas as pd
import numpy as np
from datetime import datetime
from matplotlib import pyplot as plt
from pandas.plotting import register_matplotlib_converters
register_matplotlib_converters()
import os
from smtplib import SMTP
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from jinja2 import Environment, FileSystemLoader
import pdfkit

class MeetingAnalytics:
    def __init__(self,workbook,sheet,save_csv=True):
        self.input = workbook
        self.sheet_name = sheet
        self.csv_output = os.path.dirname(self.input) + '_' + self.sheet_name + '.csv'
        self.process_data()
        self.save_csv = save_csv
        self.env = Environment(
            loader=FileSystemLoader('%s/templates/' % os.path.dirname(self.input)))
        return

    def process_data(self):
        self.zoom = pd.read_excel(open(self.input, 'rb'), self.sheet_name)
        return

    def join_time(self,c,zoom_yes):
        # pTimes = zoom_yes.loc[c,]
        join_time = None
        # 'Oct 23, 2020 20:41:05'
        for i in c:
            if join_time == None:
                join_time = datetime.strptime(zoom_yes.loc[i, 'Join Time'], '%b %d, %Y %H:%M:%S')
            elif join_time > datetime.strptime(zoom_yes.loc[i, 'Join Time'], '%b %d, %Y %H:%M:%S'):
                join_time = datetime.strptime(zoom_yes.loc[i, 'Join Time'], '%b %d, %Y %H:%M:%S')
        return (join_time)

    def leave_time(self,c,zoom_yes):
        # pTimes = zoom_yes.loc[c,]
        leave_time = None
        # 'Oct 23, 2020 20:41:05'
        for i in c:
            if leave_time == None:
                leave_time = datetime.strptime(zoom_yes.loc[i, 'Leave Time'], '%b %d, %Y %H:%M:%S')
            elif leave_time < datetime.strptime(zoom_yes.loc[i, 'Leave Time'], '%b %d, %Y %H:%M:%S'):
                leave_time = datetime.strptime(zoom_yes.loc[i, 'Leave Time'], '%b %d, %Y %H:%M:%S')
        return (leave_time)

    def minute_finder(self,t,attendee_df):
        count = 0
        # 'Oct 23, 2020 20:41:05'
        for i in range(len(attendee_df)):
            if t > attendee_df.loc[i, 'JoinTime'] and t < attendee_df.loc[i, 'LeaveTime']:
                count = count + 1
        return (count)

    def metric_total_participation_quality(self):
        '''
        Metric : metric_total_participation_quality
        Participants vs Total Time
        This will give overall attendee participation time overall health

        :return:
        '''
        zoom_yes = self.zoom[self.zoom.Attended == 'Yes']

        zoom_yes = zoom_yes.iloc[:, [0, 9]]
        zoom_grouped = zoom_yes.groupby(['Email']).sum()
        zoom_grouped = zoom_grouped.sort_values('Time in Session (minutes)', ascending=False)

        zoom_grouped = zoom_grouped.reset_index()
        zoom_grouped['Name'] = np.zeros(len(zoom_grouped))

        for e in zoom_grouped.Email:
            Name = self.zoom.loc[self.zoom['Email'] == e, 'User Name (Original Name)']
            zoom_grouped.loc[zoom_grouped.Email == e, 'Name'] = Name.iloc[0]

        # TODO Update to Database
        if self.save_csv:
            zoom_grouped.to_csv(self.csv_output)
        return(zoom_grouped['Time in Session (minutes)'].values)

    def metric_join_leave_window(self):
        '''
         - Combine all join and end time for a user.
         - sort all users in order of join time.
         - plot join time as line plot
         - plot end time as line plot
        This will infer individual joined meeting duration in graph
        :return:
        '''

        zoom = pd.read_excel(open(self.input, 'rb'), self.sheet_name)
        zoom_yes = zoom[zoom.Attended == 'Yes']

        zoom_yes_grp = zoom_yes.groupby(by='Email')
        attendee_df = pd.DataFrame.from_dict({k: [v] for k, v in zoom_yes_grp.groups.items()}, orient='index', dtype='object')
        attendee_df = attendee_df.reset_index()
        attendee_df.columns = ['Email', 'Positions']
        attendee_df['Positions'] = attendee_df['Positions'].apply(lambda c: c.to_list())

        attendee_df['JoinTime'] = attendee_df['Positions'].apply(self.join_time,args=(zoom_yes,))
        attendee_df['LeaveTime'] = attendee_df['Positions'].apply(self.leave_time,args=(zoom_yes,))
        attendee_df = attendee_df.sort_values(by="JoinTime")
        return(attendee_df['JoinTime'].values,attendee_df['LeaveTime'].values)

    def metric_attendance_ratio(self):
        '''
        Attended - Yes vs No Ratio
        :return:
        '''

        zoom_yes = self.zoom[self.zoom.Attended == 'Yes']
        zoom_yes_grp = zoom_yes.groupby(by='Email')
        attendee_df = pd.DataFrame.from_dict({k: [v] for k, v in zoom_yes_grp.groups.items()}, orient='index', dtype='object')
        _yes = len(attendee_df)

        zoom_no = self.zoom[self.zoom.Attended == 'No']
        zoom_no_grp = zoom_no.groupby(by='Email')
        attendee_no_df = pd.DataFrame.from_dict({k: [v] for k, v in zoom_no_grp.groups.items()}, orient='index', dtype='object')
        attendee_no_df = attendee_no_df.reset_index()
        attendee_no_df.columns = ['Email', 'Positions']
        _no = len(attendee_no_df)
        return (_yes,_no)

    def metric_participant_count_byminute(self):
        '''
        Count number of attendee at any point of minute.
        (Number of participants vs every minute data) - Calculate this data & produce a graph.
         - Combine all join and end time for a user.
         - Get the admin join time & end time
         - count for very time second or minute how many users were present
         - plot no of participants against time line in second or minute.
        This will infer attendee drop time
        This will give Maximum number of attendee at any time in a meeting
        :return:
        '''

        zoom_yes = self.zoom[self.zoom.Attended == 'Yes']
        zoom_yes_grp = zoom_yes.groupby(by='Email')
        attendee_df = pd.DataFrame.from_dict({k: [v] for k, v in zoom_yes_grp.groups.items()}, orient='index', dtype='object')
        attendee_df = attendee_df.reset_index()
        attendee_df.columns = ['Email', 'Positions']
        attendee_df['Positions'] = attendee_df['Positions'].apply(lambda c: c.to_list())

        attendee_df['JoinTime'] = attendee_df['Positions'].apply(self.join_time,args=(zoom_yes,))
        attendee_df['LeaveTime'] = attendee_df['Positions'].apply(self.leave_time,args=(zoom_yes,))
        attendee_df = attendee_df.sort_values(by="JoinTime")

        _min = min(attendee_df['JoinTime'])
        _max = max(attendee_df['LeaveTime'])
        TimePlot = pd.DataFrame(pd.date_range(_min, _max, freq="1min"))
        TimePlot['Count'] = np.zeros(len(TimePlot))
        TimePlot.columns = ['Time', 'Count']

        TimePlot['Count'] = TimePlot['Time'].apply(self.minute_finder,args=(attendee_df,))
        x = list(range(len(TimePlot['Time'])))
        return(x,TimePlot['Count'].to_list())

    def generate_image(self):
        fig,ax = plt.subplots(2,2)
        wspace = 0.4
        hspace = 0.4
        w = 19.2
        h = 12.0
        fig.set_size_inches(w, h)
        plt.subplots_adjust(wspace =wspace,hspace=hspace)
        fig.suptitle(self.sheet_name, fontsize=12)

        sessiontime = self.metric_total_participation_quality()
        ax[0,0].plot(sessiontime)
        ax[0,0].set_ylabel('Session Duration / Participant')
        ax[0,0].set_xlabel('Participants')
        ax[0,0].set_title("Total participation quality (sorted)")

        jointime,leavetime = self.metric_join_leave_window()
        ax[0,1].plot(jointime)
        ax[0,1].plot(leavetime)
        ax[0,1].set_ylabel('Time')
        ax[0,1].set_xlabel('Participants')
        ax[0,1].set_title("Each Participant Join & Leave Time")

        _yes,_no = self.metric_attendance_ratio()
        ax[1,0].set_ylabel('Participant Count')
        ax[1,0].set_xlabel('Meeting Attended Yes/No')
        ax[1,0].set_title("Attendance ratio: " + str(_yes / (_yes + _no)))
        ax[1,0].bar(['Yes','No'],[_yes, _no])

        x,y = self.metric_participant_count_byminute()
        ax[1,1].set_ylabel('Participant Count')
        ax[1,1].set_xlabel('Time Duration (Minutes)')
        ax[1,1].set_title('Participant count during the meeting')
        ax[1,1].plot(x, y)
        ax[1,1].set_xticks(x)
        plt.setp(ax[1,1].get_xticklabels(), rotation=90, horizontalalignment='right')

        #plt.show()

        fig.savefig(os.path.dirname(self.input)+'_'+self.sheet_name+'.jpg')
        return

    def get_data(self):
        data = []
        data.append(
            {
                "movies": [
                    {
                        "title": 'Terminator',
                        "description": 'One soldier is sent back to protect her from the killing machine. He must find Sarah before the Terminator can carry out its mission.'
                    },
                    {
                        "title": 'Seven Years in Tibet',
                        "description": 'Seven Years in Tibet is a 1997 American biographical war drama film based on the 1952 book of the same name written by Austrian mountaineer Heinrich Harrer on his experiences in Tibet.'
                    },
                    {
                        "title": 'The Lion King',
                        "description": 'A young lion prince is born in Africa, thus making his uncle Scar the second in line to the throne. Scar plots with the hyenas to kill King Mufasa and Prince Simba, thus making himself King. The King is killed and Simba is led to believe by Scar that it was his fault, and so flees the kingdom in shame.'
                    }
                ]
            })
        return data

    def send_mail(self,bodyContent):
        '''
        TODO: # Email - Gmail login failed (Blocks)
              Login not allowed for less secure apps
        :param bodyContent:
        :return:
        '''
        to_email = 'user@gmail.com'
        from_email = 'author@email.com'
        subject = 'Webinar Analysis Report'
        message = MIMEMultipart()
        message['Subject'] = subject
        message['From'] = from_email
        message['To'] = to_email

        message.attach(MIMEText(bodyContent, "html"))
        msgBody = message.as_string()

        server = SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, 'Author@123')
        server.sendmail(from_email, to_email, msgBody)
        server.quit()
        return

    def build_pdf_report(self,msgBody):
        '''
           Define App path - Add metrics images to html template
           Jinja - Templates to build HTML,
           wkhtmltopdf & pdfkit - convert html to PDF
        :param msgBody:
        :return:
        '''
        pdfkit.from_string(msgBody, 'analysis_report.pdf')

    def generate_report(self):
        json_data = self.get_data()
        template = self.env.get_template('child.html')
        output = template.render(data=json_data[0])
        self.send_mail(output)
        return


if __name__ ==  "__main__":
    objMA = MeetingAnalytics('Webinar.xls','Attendee')
    # CSV Timer Report generated in generate_image step automatically
    objMA.generate_image()
