import smtplib
def send_email_func(to_,subj_,msg_,from_,pass_):

    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(from_,pass_)
    msg='Subject:{}\n\n{}'.format(subj_,msg_)
    s.sendmail(from_,to_,msg)
    x=s.ehlo()
    if x[0]==250:
        return 's'
    else:
        return 'f'

    s.close()



#xlrd module used to read data from excel files
#openpyxl module used to read excel file
#latest xlrd version can read only xls file not xlsx file