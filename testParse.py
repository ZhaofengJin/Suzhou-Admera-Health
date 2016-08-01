#coding:utf-8 
from reportlab.platypus import *
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.rl_config import defaultPageSize
import reportlab.rl_config
from reportlab.lib.units import inch,mm
from reportlab.lib import colors
from reportlab.lib.colors import HexColor, toColor
from reportlab.lib.pagesizes import letter
from reportlab.lib.enums import TA_LEFT, TA_RIGHT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

from datetime import timedelta, date, datetime
from ClinTrials.ClinTrialsUtil import *
import sys, codecs, os, argparse, re, pprint
reload(sys)
sys.setdefaultencoding("utf-8")

reportlab.rl_config.warnOnMissingFontGlyphs = 0

from reportlab.pdfbase import pdfmetrics 
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFontFamily 

scriptFolder = os.path.dirname(os.path.abspath(sys.argv[0]))
pdfmetrics.registerFont(TTFont('Calibri',scriptFolder + '/fonts/Calibri.ttf'))
pdfmetrics.registerFont(TTFont('Calibri-Bold',scriptFolder + '/fonts/Calibri-Bold.ttf'))
pdfmetrics.registerFont(TTFont('hei', scriptFolder+'/fonts/wqy-microhei.ttc'))
pdfmetrics.registerFont(TTFont('hei-Bold', scriptFolder+'/fonts/msyahei-Bold.ttf'))
from reportlab.lib import fonts,colors
fonts.addMapping('hei', 0, 0, 'hei')
fonts.addMapping('hei-Bold', 1, 0, 'hei-Bold')

def ParseArg():
	p = argparse.ArgumentParser(description = 'generate report for Select test report')
	p.add_argument('input',type = str, help = "input data file" )
	if len(sys.argv) == 1:
		print >> sys.stderr, p.print_help()
		sys.exit()
	return p.parse_args()
	
InputFile = "testDataFile.txt"
#InputFile = args.input
action_file = codecs.open(InputFile,'r',encoding = 'utf-8')

# parse the action file and store information in different variables
def ParseActionFile(action_file):
    '''
    ClinBenefits   # group 1
    LackClinBenefits   # group 2
    SideEffect   # group 3
    MutDetail   # group 4
    ClinTrial   # group 5
    GeneInfo   # group 6
    DrugInfo   # group 7
    Reference   # group 8
		if group == 1:
			if lsep[2] == "":
				result["ClinBenefits"].append(["无","","",""])
			else:
				record.append(lsep[3])
	
    '''
    # Group 1 - 8
    result={}
    Groups = ["ClinBenefits","LackClinBenefits","ClinTrial","GeneInfo","DrugInfo","Reference","Alteration_list"]
    for i in [0,1,2,3,5,6,7,8]:
        result[Groups[i]] = []
    result['ClinTrial'] = {}
    n = 0 # record the line number in action file, reset when group number changes
    prev_group = 0  # record the previous group number, help to find whether to reset n 
    prev_gene = "None" # record the previous gene name, help to find whether to reset n
    Title = []  # store title for each table
    record = []  # store the row record of each table
    Genes = []  # store the order of genes
    muDetail_n=-1 # number of mutation detail table
    geneMut_n=-1 # number of gene detail seqction
    for l in action_file.read().split("\n"):
        if l.startswith("Group") or l.strip()=="": continue # skip the first line
        lsep = l.split("\t")
        group=int(lsep[0])
        if group!=prev_group or (lsep[1]!=prev_gene and group in [4,5,6]):  # when group changes
            n=0
            Title = []
        n+=1
        if group<4:
            if lsep[2]=="":
                result[Groups[group-1]].append(["Gene","Alteration Detected","Therapies","Tumor Type","Reference"])
                result[Groups[group-1]].append(["No medically actionable mutations were detected in this category.","","",""])
            else:
                record.append(lsep[3])
                if n<=4:
                    Title.append(lsep[2])
                if n%4==0:
                    if n==4:
                        Title.insert(0,"Gene")
                        result[Groups[group-1]].append(Title)
                    record.insert(0,lsep[1])
                    result[Groups[group-1]].append(record)
                    record=[]
        elif group==4:
            if lsep[2]=="Nucleotide":
                muDetail_n+=1
                result[Groups[group-1]].append("")
                Gene_title = Paragraph("<font color='#ffc000'><b><u><i>%s</i></u></b>"%lsep[1],ParaStyle)
                Gene_cell = Paragraph("<font color='#ffffff'><b>Gene: <i>%s</i></b>"%lsep[1],ParaStyle)
                result[Groups[group-1]][muDetail_n]=[[Gene_title,"",""],[Gene_cell,"Nucleotide: %s"%lsep[3],""],["","",""]]  # create first three row of alteration details table
            elif lsep[2]=="Pathways":
                try:
                    current_line_num = len(result[Groups[group-1]][muDetail_n])
                except:
                    current_line_num = 0
                if current_line_num==3:
                        result[Groups[group-1]][muDetail_n][1][2]="Pathways: %s"%lsep[3]  # add pathway information
                else:    # if Pathway is the first line for this alteration, for Amplification or Fusion
                    muDetail_n+=1
                    assert len(result[Groups[group-1]])==muDetail_n  # confirm the current index is correct
                    result[Groups[group-1]].append("")
                    Gene_title = Paragraph("<font color='#ffc000'><b><u><i>%s</i></u></b>"%lsep[1],ParaStyle)
                    Gene_cell = Paragraph("<font color='#ffffff'><b>Gene: <i>%s</i></b>"%lsep[1],ParaStyle)
                    result[Groups[group-1]][muDetail_n]=[[Gene_title,"",""],[Gene_cell,"","Pathways: %s"%lsep[3]],["","",""]]  # create first three row of alteration details table
            elif lsep[2]=="Alteration Detected":
                result[Groups[group-1]][muDetail_n][2][0]="Alteration Detected: %s"%lsep[3]
            elif lsep[2]=="Variation Type":
                result[Groups[group-1]][muDetail_n][2][2]="Variation Type: %s"%lsep[3]
            else:  # Details
                if len(lsep)<5: # no response
                    response = ""
                elif lsep[4].startswith("Increased") or lsep[4].startswith("Potential Clinical Benefit"):  # green color
                    if lsep[4].endswith("#"):    # different cancer type benefit
                        response = Paragraph("<font color='#0067b1'><b>%s</b></font>"%lsep[4].replace("#",""),ParagraphStyle(name = 'Normal', leftIndent=15), bulletText="\xe2\x9e\xa2")
                    else:    # same cancer type benefit
                        response = Paragraph("<font color='#238943'><b>%s</b></font>"%lsep[4],ParagraphStyle(name = 'Normal', leftIndent=15), bulletText="\xe2\x9e\xa2")
                else:  # red color
                    response = Paragraph("<font color='#d89234'><b>%s</b></font>"%lsep[4],ParagraphStyle(name = 'Normal', leftIndent=15), bulletText="\xe2\x9e\xa2")
                result[Groups[group-1]][muDetail_n].append([Paragraph('<b>%s</b>'%lsep[3],ParaStyle),"",response])    
        elif group==5:
            record.append(lsep[3])
            if n<=5:
                Title.append(lsep[2])
            if n%5==0:
                if lsep[1] not in result[Groups[group-1]]:
                    Genes.append(lsep[1])
                    result[Groups[group-1]][lsep[1]]=[]
                if n==5:
                    result[Groups[group-1]][lsep[1]].append(Title)
                result[Groups[group-1]][lsep[1]].append(record)
                record=[]
        elif group==6:
            print lsep
            if lsep[2]=="Comment":
                geneMut_n+=1
                result[Groups[group-1]].append({})
                result[Groups[group-1]][geneMut_n]["Comment"]=lsep[3]
                result[Groups[group-1]][geneMut_n]['Gene']=lsep[1]
            else:
                result[Groups[group-1]][geneMut_n][lsep[3]]=lsep[4]
        elif group==7:
            result[Groups[group-1]].append([lsep[1],lsep[3]])
        elif group==8:
            result[Groups[group-1]].append(lsep[1])
        elif group==9:
            result[Groups[group-1]].extend(lsep[1:])
        prev_group=group
        prev_gene=lsep[1]
   
    return result, Genes
	
result, Genes = ParseActionFile(action_file)
#table21Data = result['']

pprint.pprint(result)
action_file.close()