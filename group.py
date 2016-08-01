#coding:utf-8
from datetime import timedelta, date, datetime
import sys, os

if len(sys.argv)<2:
    print >> sys.stderr,"\n The group code requires the floder as the input"
    sys.exit()
    
run_folder = sys.argv[1]

if os.path.isdir(run_folder)==0:
    print >> sys.stderr,"Not a valid folder"
    sys.exit()

files = os.listdir(run_folder)
scriptFolder = os.path.dirname(os.path.abspath(sys.argv[0]))+"/"
    
sample_info_file = run_folder+"/Sample_information.txt"
sample_info = open(sample_info_file,'r')
for line in sample_info:
    if line.startswith("Sample"):    
        continue
    sep = line.split("\t")
    sID = sep[0]
    pName = sep[6]
    age = sep[8]
    if " " in age:
        age = age.strip()
        if age == "":
            age = "/"
    pID = sep[1]
    type = sep[2]
    cName = sep[3]
    if sep[4]=="Colorectal Cancer":
        cType = "T"
    else:
        cType = "F"
    CNV = ""
    input = ""
    ReportDate = date.today()
    output = "OncoGxSelect技术服务报告-"+pName+"-"+ReportDate.strftime('%Y%m%d')+".pdf"
    for file in files:
        if file.startswith(sID) and file.endswith("_interpretation.txt"):
            input = file
        if file.startswith(sID) and file.endswith("_CNV.txt"):
            CNV = file
    if input!="":
        if CNV != "":
            command = "python /home/agis/Softwares/OncoGxSelectV2_package/testPdfF2.py %s -o %s -p %s -i %s -s %s -t %s -c %s -f %s -C %s -T %s -a %s"%(run_folder+"/"+input,run_folder+"/"+output,pName,pID,sID,type,cName,run_folder,run_folder+"/"+CNV,cType,age)
        else:
            command = "python /home/agis/Softwares/OncoGxSelectV2_package/testPdfF2.py %s -o %s -p %s -i %s -s %s -t %s -c %s -f %s -T %s -a %s"%(run_folder+"/"+input,run_folder+"/"+output,pName,pID,sID,type,cName,run_folder,cType,age)
        print(command)
        os.system(command)

        
        
        
        
        