# -*- coding: utf-8 -*-
"""
creer le 05/07/2018 21:40:15

auteur : Olivier Dagonneau

 gestion objets Microsoft

"""

import win32com.client as win32
import psutil
import os
import subprocess
from time import sleep
import datetime
from win32com.client import Dispatch


def test_oulook():
    for item in psutil.pids():
        p = psutil.Process(item)
        if p.name() == "OUTLOOK.EXE":
            flag = 1
            break
        else:
            flag = 0
    if (flag == 0):
        try:
            subprocess.Popen(['C:\Program Files (x86)\Microsoft Office\Office14\Outlook.exe'])
        except:
            print("Outlook didn't open successfully")

class Outlook(object):
    """class Outlook gestion mail et RDV
    dp = destinataire principal
    df = destinatataire facultatif pour un RDV ou destinataire en copie pour un mail
    su = sujet
    bd = corps du mail ou rdv
    de = date de debut pour un RDV
    fin = date de Fin pour un RDV
    exemples mail :
      mail_init=Outlook("olivier.dagonneau@posteo.net",su="test",bd="test")
      mail_init.mail()

    exemple RDV :
      rdv_init=Outlook("olivier.dagonneau@posteo.net",su="test",bd="test",de="13/07/2018 15:00:00",fi="13/07/2018 16:00:00")
      rdv_init.rdv()

    """



    def __init__(self,dp="",df="",su="",bd="",de="25/05/2018 10:00:00",fi="25/05/2018 10:00:00"):
        self.destp=dp
        self.destf=df
        self.subject=su
        self.body=bd
        self.debut=de
        self.fin=fi



    def mail(self):
        test_oulook()
        o = win32.Dispatch('outlook.application')
        email = o.CreateItem(0)
        email.To = self.destp
        email.CC = self.destf
        email.Subject = self.subject
        email.body = self.body
        email.Display(True)

    def rdv(self):
        test_oulook()
        dest_opt = self.destf.split(";")
        o = win32.Dispatch('outlook.application')
        rdvo = o.CreateItem(1)
        rdvo.MeetingStatus = 1  # rendez-vous sous la forme d’une demande de réunion
        if self.destp:
            dest_ = self.destp.split(";")
            for i in dest_:
                participants_obl = rdvo.Recipients.Add(i)
                #participants_obl.Type = 1
        if self.destf:
            dest_opt = self.destf.split(";")
            #print(dest_opt)
            for j in dest_opt:
                participants_optionnels = rdvo.Recipients.Add(j)
                participants_optionnels.Type = 2
        rdvo.Subject = self.subject
        rdvo.body = self.body
        rdvo.Start = self.debut
        rdvo.End = self.fin
        rdvo.Display(True)

if __name__ == "__main__":
    mail_init=Outlook("olivier.dagonneau@posteo.net",su="test",bd="test")
    mail_init.mail()
    rdv_init=Outlook("olivier.dagonneau@posteo.net",su="test",bd="test",de="13/07/2018 15:00:00",fi="13/07/2018 16:00:00")
    rdv_init.rdv()
