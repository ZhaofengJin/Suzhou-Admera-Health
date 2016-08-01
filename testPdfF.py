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
from decimal import Decimal
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

PAGE_HEIGHT=defaultPageSize[1]
styles = getSampleStyleSheet()
HeaderStyle = styles["Heading1"]
ParaStyle = styles["Normal"]
PreStyle = styles["Code"]
StylePT = styles["BodyText"] # paragraph style for paragraph in table
StylePT.wordWrap="CJK"# allow long word wrap http://stackoverflow.com/questions/11839697/wrap-text-is-not-working-with-reportlab-simpledoctemplate
Genes = []
result = []


#the translation functions     
from microsofttranslator import Translator
gs = Translator('AGIS', '/2Qtygt7UcHNY1o7+cH67KLd+TpH+A4oZxaE6WQ7WoI=')

scriptFolder = os.path.dirname(os.path.abspath(sys.argv[0]))+"/"
trans_list={}
trans_file=codecs.open(scriptFolder+"translation.txt","r",'utf-8')
for i in trans_file.read().split("\n"):
    if i.strip()=="": continue
    i=i.split("\t")
    trans_list[i[0].strip()]=i[1]
trans_file.close()
def translate(eng_word):
    "curated list or google translation"
    if eng_word.strip()=="":
        return ""
    eng_word=eng_word.strip()
    if eng_word not in trans_list:
        ch_word = gs.translate(eng_word, 'zh')
        trans_list[eng_word]=ch_word
    return(trans_list[eng_word])

drug_transl_list={}
drug_transl_file = codecs.open(scriptFolder+"drug_translation.txt","r",'utf-8')
for i in drug_transl_file:
    if i.strip()=="": continue
    i=i.split("\t")
    if i[0].strip()!=i[1].strip():
        drug_transl_list[i[0].strip()]="%s(%s)"%(i[1],i[0].strip())
drug_transl_file.close()

description_transl_file = codecs.open(scriptFolder+"description_translation.txt","r",'utf-8')
for i in description_transl_file:
    if i.strip()=="": continue
    i=i.split("\t")
    if i[0].strip()!=i[1].strip():
        drug_transl_list[i[0].strip()]=i[1]


drug_transl_list = dict((re.escape(k), v) for k, v in drug_transl_list.iteritems())
def drugTranslation(expression):
    pattern = re.compile("|".join(drug_transl_list.keys()))
    return(pattern.sub(lambda m: drug_transl_list[re.escape(m.group(0))], expression))
    

state_to_code = {"JIANGSU":'Jiangsu',"VERMONT": "VT", "GEORGIA": "GA", "IOWA": "IA", "Armed Forces Pacific": "AP", "GUAM": "GU", "KANSAS": "KS", "FLORIDA": "FL", "AMERICAN SAMOA": "AS", "NORTH CAROLINA": "NC", "HAWAII": "HI", "NEW YORK": "NY", "CALIFORNIA": "CA", "ALABAMA": "AL", "IDAHO": "ID", "FEDERATED STATES OF MICRONESIA": "FM", "Armed Forces Americas": "AA", "DELAWARE": "DE", "ALASKA": "AK", "ILLINOIS": "IL", "Armed Forces Africa": "AE", "SOUTH DAKOTA": "SD", "CONNECTICUT": "CT", "MONTANA": "MT", "MASSACHUSETTS": "MA", "PUERTO RICO": "PR", "Armed Forces Canada": "AE", "NEW HAMPSHIRE": "NH", "MARYLAND": "MD", "NEW MEXICO": "NM", "MISSISSIPPI": "MS", "TENNESSEE": "TN", "PALAU": "PW", "COLORADO": "CO", "Armed Forces Middle East": "AE", "NEW JERSEY": "NJ", "UTAH": "UT", "MICHIGAN": "MI", "WEST VIRGINIA": "WV", "WASHINGTON": "WA", "MINNESOTA": "MN", "OREGON": "OR", "VIRGINIA": "VA", "VIRGIN ISLANDS": "VI", "MARSHALL ISLANDS": "MH", "WYOMING": "WY", "OHIO": "OH", "SOUTH CAROLINA": "SC", "INDIANA": "IN", "NEVADA": "NV", "LOUISIANA": "LA", "NORTHERN MARIANA ISLANDS": "MP", "NEBRASKA": "NE", "ARIZONA": "AZ", "WISCONSIN": "WI", "NORTH DAKOTA": "ND", "Armed Forces Europe": "AE", "PENNSYLVANIA": "PA", "OKLAHOMA": "OK", "KENTUCKY": "KY", "RHODE ISLAND": "RI", "DISTRICT OF COLUMBIA": "DC", "ARKANSAS": "AR", "MISSOURI": "MO", "TEXAS": "TX", "MAINE": "ME"}

code_to_state = {v: k for k, v in state_to_code.items()}

#function for making the header
def header(txt, style = HeaderStyle, klass = Paragraph, sep = 0.1, *args, **kwargs):
    s = Spacer(0.2*inch, sep * inch)
    para = klass(txt,style = style, *args, **kwargs)
    sect = [s, para]
    result = KeepTogether(sect)
    return result 

def pa(txt):
    return header(txt,style = ParaStyle, sep = 0.1)
'''
if not os.path.exists(scriptFolder+"OncoGxSelectV2_Panel_info.txt"):
    print "Please put the panel_info file in the same folder with the program"
    sys.exit(0)
'''

# stripe background tables
def stripe_table(data, title, color, firstCT=False, sep=0.4):
    ''' Create table and title for each type of mutations 
          color: hex color for the title line
          image: image for the first cell in the table
          firstCT: whether it is the first table of clinical trials
          sep: spacer on top        '''
    ParaStyle=ParagraphStyle("Normal", alignment=TA_CENTER)
    Title_num=1+int(firstCT)  # number of tows for titles
    tableStyle = [
                 ('GRID',(0,Title_num+1),(-1,-1),1,colors.grey),
                 ('BACKGROUND',(0,0),(-1,0),colors.HexColor(color)),
                 ('BACKGROUND',(0,Title_num),(-1,Title_num),colors.HexColor('#404040')),
                 ('TEXTCOLOR',(0,0),(-1,0),colors.HexColor('#ffffff')),
                 ('TEXTCOLOR',(0,Title_num),(-1,Title_num),colors.HexColor('#ffc000')),
                 ('TEXTCOLOR',(0,Title_num+1),(-1,-1),colors.HexColor('#000000')),
                 ('FONTSIZE',(0,0),(-1,-1),10),
                 ('FONT',(0,0),(-1,-1),'hei'),
                 ('FONT',(0,0),(-1,Title_num),'hei-Bold'),
                 ('VALIGN',(0,0),(-1,-1),"MIDDLE"),
                 ('ALIGN',(0,0),(-1,Title_num),"CENTER"),
                 ]
    if title.endswith("Trials"):
        subtitle = [""]*len(data[0])
        subtitle[0] = title.replace("Trials","")
        data.insert(0, subtitle)
        colWidths=[32*mm, 30*mm, 80*mm, 15*mm, 30*mm]
        tableStyle.extend([('ALIGN',(0,Title_num+1),(0,-1),"LEFT"),
                           ('ALIGN',(2,Title_num+1),(2,-1),"LEFT"),
                           ('ALIGN',(4,Title_num+1),(4,-1),"LEFT"),
                           ('SPAN',(0,Title_num-1),(-1,Title_num-1)),
                           ('BACKGROUND',(0,Title_num-1),(-1,Title_num-1),colors.HexColor("#af97d9")),
                           ('TEXTCOLOR',(0,Title_num-1),(-1,Title_num-1),colors.HexColor('#000000')),])
    else:
        colWidths=[19*mm, 36*mm, 60*mm, 36*mm, 36*mm]
    if firstCT or not title.endswith("Trials"):
        Title_row = [""]*len(data[0])
        Title_row[1] = title.replace("Trials","")
        if firstCT:
            Title_row[1] = "靶向治疗相关的临床试验    "
        Title_row[0] ="" #Image(PythonImage(scriptFolder+'./img/'+image),width = 11.5*mm, height = 11.5*mm)
        data.insert(0, Title_row)
        tableStyle.extend([('SPAN',(1,0),(-1-int(firstCT),0)),('ALIGN',(0,0),(0,0),"LEFT"),('FONTSIZE',(1,0),(-1,0),12)])
    if data[-1][1]=="":
        tableStyle.append(('SPAN',(0,-1),(-1,-1)))
        data[-1][0] = Paragraph("<font size=10 name='hei'><b><i>%s</b></i></font>"%data[-1][0],ParaStyle)
    else: 
        for i in range(Title_num+1,len(data)):
            for j in range(len(data[i])):
                if data[i][j].startswith("NCT"):
                    data[i][j] = Paragraph("<font size=10><link href='https://clinicaltrials.gov/ct2/show/%s' color='#4f81bd'><u>%s</u></link></font>"%(data[i][j],data[i][j]),ParaStyle)
                elif j==0 and data[1][0]=="基因":
                    data[i][j] = Paragraph("<font size=10><i>%s<i/></font>"%data[i][j],ParaStyle)  # gene name need to be italic
                elif data[Title_num][j] in ["Therapies","Title","Locations#"] and title.endswith("Trials"):
                    data[i][j] = Paragraph("<font size=10 name='hei'>%s</font>"%data[i][j],ParagraphStyle("Normal")) # don't align center for these columns in CT table
                else:
                    data[i][j] = Paragraph("<font size=10 name='hei'>%s</font>"%data[i][j].replace("||","<br /><br />"),ParaStyle)  # change type to paragraph so that it can be wrapped.
    if title.endswith("Trials"):
        data[Title_num] = [translate(x) for x in data[Title_num]]
    StripeTable = Table(data, style = tableStyle, repeatRows=2, colWidths=colWidths)
    sect = [Spacer(0.2*inch, sep*inch), StripeTable]
    result_table = KeepTogether(sect)
    return result_table

def PythonImage(png):
    jpg = png.replace("png","jpg")
    if not os.path.exists(jpg):
        tmp = open(png)
        f = open(png.replace("png","jpg"),'w')
        f.write(tmp.read())
        f.close()
        tmp.close()
    return jpg

def myFirstPage(canvas,doc):
    canvas.saveState()
    canvas.drawImage(scriptFolder+"/img/admera_health_final_logo.jpg",15*mm, 255*mm, 55*mm, 22*mm)
    canvas.drawImage(scriptFolder+"/img/header4.png",135*mm,257*mm,53*mm,4.5*mm)
    canvas.setFont("hei-Bold", 10)
    canvas.setFillColorRGB(0.33,0.61,0.75)
    canvas.drawRightString(197*mm, 258*mm, "-%s"%(args.pName))
    canvas.setStrokeColorRGB(0.33,0.61,0.75)
    canvas.line(15*mm,255*mm,200*mm,255*mm)
    canvas.setLineWidth(5)
    canvas.restoreState()
    
class NumberedCanvas(canvas.Canvas):
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        """add page info to each page (page x of y)"""
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.setFont("hei", 6)
        self.setFillColorRGB(0.17, 0.27, 0.57)#change color to #2b4490
        self.drawString(10*mm,15*mm,("苏州艾达康医疗科技有限公司（Admera Health）• 苏州工业园区星湖街218号生物纳米园B1栋7楼702"))
        self.drawString(10*mm,12*mm,("Customerservice@admerahealth.com.cn • 0512-62628766"))
        self.setFont("Helvetica", 8)
        if self._pageNumber>2:
            self.drawRightString(200*mm, 10*mm,
                "%d/%d" % (self._pageNumber-2, page_count-2))
            
def shorternNucleoChange(NucleoChange):
    '''
    If the NucleoChane string is too long, then remove the nucleotide after 'del' or 'ins'
    For example: c.2236_2250delGAATTAAGAGAAG  ->  c.2236_2250del
    '''
    if len(NucleoChange)>20:
        trimed_end = [m.end() for m in re.finditer("del|ins", NucleoChange)][-1]
        return NucleoChange[:trimed_end]
    else:
        return NucleoChange
    
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
    '''
    # Group 1 - 8
    result={}
    Groups = ["ClinBenefitsSame","ClinBenefitsDiff","LackClinBenefits","MutDetail","ClinTrial","GeneInfo","DrugInfo","Reference","Alteration_list"]
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
                result[Groups[group-1]].append(["基因","检测到的突变","治疗方案","癌症类型","参考文献"])
                result[Groups[group-1]].append(["无"])
            else:
                record.append(lsep[3])
                if n<=4:
                    Title.append(lsep[2])
                if n%4==0:
                    if n==4:
                        Title.insert(0,"Gene")
                        result[Groups[group-1]].append(["基因","检测到的突变","治疗方案","癌症类型","参考文献"])
                    record.insert(0,lsep[1])
                    record[2] = drugTranslation(record[2])   # translate the drug name
                    record[3] = translate(record[3])
                    if record[4] == "NCCN Guideline":
                       record[4] = translate(record[4])
                    
                    result[Groups[group-1]].append(record)
                    record=[]
                        
        elif group==4:
            if lsep[2]=="Nucleotide":
                muDetail_n+=1
                result[Groups[group-1]].append("")
                Gene_title = Paragraph("<font color='#ffc000'><b><u><i>%s</i></u></b></font>"%lsep[1],ParaStyle)
                Gene_cell = Paragraph("<font color='#ffffff' name='hei'>%s</font><font color='#ffffff' name='Helvetica'><b><i>%s</i></b></font>"%("基因: ".decode('utf-8'),lsep[1]),ParaStyle)
                result[Groups[group-1]][muDetail_n]=[[Gene_title,"",""],[Gene_cell,"核酸变化: ".decode('utf-8')+shorternNucleoChange(lsep[3]),""],["","",""]]  # create first three row of alteration details table
            elif lsep[2]=="Pathways":
                try:
                    current_line_num = len(result[Groups[group-1]][muDetail_n])
                except:
                    current_line_num = 0
                if current_line_num==3:
                        result[Groups[group-1]][muDetail_n][1][2]="通路: ".decode('utf-8')+lsep[3]  # add pathway information
                else:    # if Pathway is the first line for this alteration, for Amplification or Fusion
                    muDetail_n+=1
                    assert len(result[Groups[group-1]])==muDetail_n  # confirm the current index is correct
                    result[Groups[group-1]].append("")
                    Gene_title = Paragraph("<font color='#ffc000'><b><u><i>%s</i></u></b></font>"%lsep[1],ParaStyle)
                    Gene_cell = Paragraph("<font color='#ffffff' name='hei'>%s</font><font color='#ffffff' name='Helvetica'><b><i>%s</i></b></font>"%("基因: ".decode('utf-8'),lsep[1]),ParaStyle)
                    result[Groups[group-1]][muDetail_n]=[[Gene_title,"",""],[Gene_cell,"","通路: ".decode('utf-8')+lsep[3]],["","",""]]  # create first three row of alteration details table
            elif lsep[2]=="Alteration Detected":
                result[Groups[group-1]][muDetail_n][2][0]="检测到的突变: ".decode('utf-8')+lsep[3]
            elif lsep[2]=="Variation Type":
                result[Groups[group-1]][muDetail_n][2][2]="突变类型: ".decode('utf-8')+lsep[3]
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
                record[0] = drugTranslation(record[0])
                result[Groups[group-1]][lsep[1]].append(record)
                record=[]
        elif group==6:
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

def ParseArg():
    p=argparse.ArgumentParser( description = 'generate Chinese report for PGxOne genetic test')
    p.add_argument('input',type=str,help="input file with all information to generate report, *_interpretation.txt")
    p.add_argument('-o', '--outputPDF', type=str, help="output pdf file")
    p.add_argument("-p","--pName",type=str, help="Patient name", default=" ")
    p.add_argument("-i","--pID",type=str, help="Project ID", default=" ")
    p.add_argument("-s","--sID",type=str, help="Sample ID", default=" ")
    p.add_argument("-t","--Source",type=str, help="Sample Type", default="FFPE Slides")
    p.add_argument("-c","--cName",type=str, help="Name of Canser", default="")
    p.add_argument("-a","--age",type=str, help="Patient Age", default="")
    p.add_argument("-f","--folder",type=str,help="The folder for the data files",default="")
    p.add_argument("-C","--CNV",type=str, help="CNV file", default="")
    if len(sys.argv)==1:
        print >> sys.stderr, p.print_help()
        sys.exit(0)
    return p.parse_args()


#input
args = ParseArg()

input = args.sID
InputFile = args.input
action_file = codecs.open(InputFile,'r',encoding = 'utf-8')

result, Genes = ParseActionFile(action_file)
action_file.close()

ReportDate = date.today()
doc = BaseDocTemplate(args.outputPDF, pagesize = letter)

ParaStyle=ParagraphStyle("Normal", alignment=TA_CENTER)


body = ParagraphStyle(name = 'body', leftIndent=0, fistLineIndent=20, spaceBefore=10,
    alignment = TA_JUSTIFY,
    wordWrap = 'CJK',
    leading = 18)

elements = [] 
tableParaStyle = ParagraphStyle(name = 'Normal', leading = 9)

#Front Page
Title1st = Paragraph("<font size=26 color=HexColor('#2a5caa') name='hei-Bold'>临床技术服务</font>", style=ParagraphStyle("Caption", alignment=TA_CENTER))
Title1 = header([[Title1st]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
Title2nd = Paragraph("<font size=22 name='hei'>OncoGxSelect<super rise=9 size=12>TM</super>检测结果报告</font>", style=ParagraphStyle("Caption", alignment=TA_CENTER))
Title2 = header([[Title2nd]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
Page_Front = [Spacer(70*inch,2.8*inch),Title1,Spacer(10.0*inch, 0.4*inch),Title2,Spacer(20.0*inch,0.8*inch)]


data = [
        ["客户信息"],
        ["姓名",args.pName, "年龄",args.age ,"病理诊断",args.cName],
        ["项目信息"],
        ["项目编号", args.pID,"","","样品类型",args.Source],
        ["报告日期 ", ReportDate.strftime('%m/%d/%Y')],
        ]

style1 = [
        ('SPAN',(0,0),(5,0)),
        ('SPAN',(0,2),(5,2)),
        ('SPAN',(1,3),(3,3)),
        ('SPAN',(1,4),(5,4)),
        ('ALIGN',(0,1),(-1,-1),"LEFT"),
        ('FONT',(0,0),(-1,-1),"hei"),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('FONTSIZE', (0, 0), (0, 0), 12),
        ('FONTSIZE', (0, 2), (0, 2), 12),
        ('VALIGN',(0,0),(0,0),'TOP'),
        ('VALIGN',(0,2),(0,2),'TOP'),
        ('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
        ('BOX', (0,0), (-1,-1), 0.25, colors.black),
        ('TEXTCOLOR',(0,0),(0,4),colors.HexColor('#2a5caa')),
        ('TEXTCOLOR',(2,1),(2,1),colors.HexColor('#2a5caa')),
        ('TEXTCOLOR',(4,1),(4,1),colors.HexColor('#2a5caa')),
        ('TEXTCOLOR',(4,3),(4,3),colors.HexColor('#2a5caa')),
         ]

Table_patient = header(data,style=style1,klass=Table,colWidths=[25*mm,40*mm,20*mm,40*mm,20*mm,40*mm], rowHeights=[8*mm,6*mm,8*mm,6*mm,6*mm])
Page_Front.append(Table_patient)
Page_Front.append(PageBreak())
elements.extend(Page_Front)

#second page
TitleImg = Paragraph("<img src=%s/img/title.png valign='middle' width='250' height='70'/>"%scriptFolder,style=ParagraphStyle("Caption", alignment=TA_CENTER))
TitleI = header([[TitleImg]],style = [("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
Title23rd = Paragraph("<font size=13 name='hei-Bold'>1.基因检测结果</font>",style=ParagraphStyle("Caption", alignment=TA_LEFT))
Title23 = header([[Title23rd]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
Page_Second = [Spacer(5*inch,0.5*inch),TitleI,Spacer(50*inch,0.4*inch),Title23,Spacer(5.0*inch,0.2*inch)]

SList = ParagraphStyle(name = 'Normal', leftIndent=10,bulletFontSize=20, bulletIndent=0, spaceBefore=2, leading = 14, wordWrap='CJK', bulletOffsetY = -3)

TempData = []
DataForPoint = []
NameFP = []
pointName = ""
DataForID = []
NameFI = []
IDName = ""
DataForAmp = []
NameFA = []
AmpName = ""
DataForFu = []
NameFF = []
FName = ""
prev_gene = "None"
action_file = codecs.open(InputFile,'r',encoding = 'utf-8')
for l in action_file.read().split("\n"):
    if l.startswith("Group") or l.strip()=="": 
        continue # skip the first line
    lsep = l.split("\t")
    group=int(lsep[0])
    if group==4:
        if lsep[1]!=prev_gene:
            TempData=[]
            prev_gene=lsep[1]
        if lsep[2]=="Nucleotide":
            TempData.append(shorternNucleoChange(lsep[3]))
        elif lsep[2]=="Pathways":
            TempData.append(lsep[3])
        elif lsep[2]=="Alteration Detected":
            TempData.append(lsep[3])
        elif lsep[2]=="Variation Type":
            type=lsep[3]
        elif lsep[2]=="Details" and len(TempData)!=0:#finish reading one example
            TempData.insert(0,prev_gene)
            if type=="Fusion":
                fusion_File = args.folder+"/"+"Fusion_data.txt"
                action_fileF = codecs.open(fusion_File,'r',encoding = 'utf-8')
                for j in action_fileF.read().split("\n"):
                    if j.startswith("ALK"):
                        continue
                    jsep = j.split("\t")
                    if input in jsep[0]:
                        if TempData[0]=="ALK":
                            fData=["ALK",jsep[1],jsep[4]]
                            NameFF.append("ALK"+"_"+"基因融合".decode("utf-8"))
                        elif TempData[0]=="ROS1":
                            fData=["ROS1",jsep[2],jsep[5]]
                            NameFF.append("ROS1"+"_"+"基因融合".decode("utf-8"))
                        elif TempData[0]=="RET":
                            fData=["RET",jsep[3],jsep[6]]
                            NameFF.append("RET"+"_"+"基因融合".decode("utf-8"))
                        fData[1] = Decimal(fData[1]).quantize(Decimal("0.00"))
                        fData[2] = Decimal(fData[2]).quantize(Decimal("0.0000"))
                DataForFu.append(fData)
                action_fileF.close()
            elif type=="Amplification":
                amplification_File = args.CNV
                action_fileA = codecs.open(amplification_File,'r',encoding = 'utf-8')
                for j in action_fileA.read().split("\n"):
                    if j.startswith(TempData[0]):
                        jsep = j.split("\t")
                        AmpMul = Decimal(jsep[1]).quantize(Decimal("0.00"))
                        if jsep[4]<0.001:
                            pValue = "<0.001"
                        else:
                            pValue = Decimal(jsep[4]).quantize(Decimal("0.000"))
                        fData = [jsep[0],AmpMul,pValue]
                        NameFA.append(jsep[0]+"_"+"基因扩增".decode('utf-8'))
                DataForAmp.append(fData)
                action_fileA.close()
            elif type=="Deletion" or type=="Insertion": #store data into the ID table combined with indel file
                #indel_File = "S6J011-v2_summary_indel.txt"
                indel_File=args.folder+"/"+"%s_summary_indel.txt"%input
                action_fileID = codecs.open(indel_File,'r',encoding = 'utf-8')
                #read the lines in the indel file 
                for j in action_fileID.read().split("\n"):
                    if j.startswith("Gene_Name") or j.strip()=="":
                        continue
                    jsep = j.split("\t")
                    if jsep[0]==TempData[0] and jsep[5]==TempData[3]:
                        indelFrequency = Decimal(jsep[6]) * 100
                        indelFrequency = str(indelFrequency.quantize(Decimal("0.00")))+"%"
                        fData=[TempData[0],TempData[3],jsep[1]+":"+jsep[2],indelFrequency,jsep[8],TempData[2]]
                        NameFI.append(jsep[0]+"-"+TempData[3])
                DataForID.append(fData)
                action_fileID.close()
            else:
                point_File =args.folder+"/"+"%s_summary_point_mutation.txt"%input
                action_fileP = codecs.open(point_File,'r',encoding = 'utf-8')
                for j in action_fileP.read().split("\n"):
                    if j.startswith("Gene_Name") or j.strip()=="":
                        continue
                    jsep = j.split("\t")
                    if jsep[0]==TempData[0] and jsep[5]==TempData[3]:
                        pointFrequency = Decimal(jsep[8]) * 100
                        pointFrequency = str(pointFrequency.quantize(Decimal("0.00")))+"%"
                        fData=[TempData[0],TempData[3]+"("+TempData[1]+")",jsep[1],pointFrequency,jsep[9],TempData[2]]
                        NameFP.append(TempData[0]+"-"+TempData[3])
                DataForPoint.append(fData)
                action_fileP.close()
                
            TempData=[]

#点突变表格

PointData = [
            ["基因","突变形式","染色体位置","突变频率","覆盖深度","基因所在的信号通路"],
            ]
            
PointTableStyle = [
                    ('FONT',(0,0),(-1,-1),"hei"),
                    ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                    ('ALIGN',(0,0),(-1,-1),"LEFT"),
                    ('BACKGROUND',(0,0),(5,0),colors.HexColor('#9f29d9')),
                    ('FONTSIZE',(0,0),(-1,-1),11),
                    ('TEXTCOLOR',(0,0),(5,0),colors.white),
                    ]

if len(DataForPoint)>0:
    for i in range(len(DataForPoint)):
        if i>0:
            if DataForPoint[i][0]==DataForPoint[i-1][0]:
                PointTableStyle.extend([('SPAN',(0,i),(0,i+1)),('VALIGN',(0,i),(0,i+1),'MIDDLE')])
            Name=Name+"、".decode('utf-8')+NameFP[i]
        else:
            Name=NameFP[i]
        for j in range(len(DataForPoint[i]
        )):
            DataForPoint[i][j] = Paragraph("<font size=11 name='hei'>%s</font>"%DataForPoint[i][j],ParaStyle)
            
    PointData.extend(DataForPoint)
    TitleForPoint = Paragraph("<font size=13 name='hei-Bold'>点突变(%s)</font>"%Name,SList, bulletText="\xe2\x80\xa2")
    pointName = Name
    TitleP = header([[TitleForPoint]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
    Point_Table = header(PointData,PointTableStyle, klass=Table,colWidths=[25*mm,40*mm,40*mm,20*mm,20*mm,40*mm])
    Page_Second.append(TitleP)
    Page_Second.append(Spacer(5.0*inch, 0.2*inch))
    Page_Second.append(Point_Table)
    Page_Second.append(Spacer(5.0*inch, 0.2*inch))

IndelData = [
            ["基因","突变形式","染色体位置","突变频率","覆盖深度","基因所在的信号通路"]
            ]

IndelTableStyle = [
                    ('FONT',(0,0),(-1,-1),"hei"),
                    ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                    ('ALIGN',(0,0),(-1,-1),"LEFT"),
                    ('BACKGROUND',(0,0),(5,0),colors.HexColor('#9f29d9')),
                    ('FONTSIZE',(0,0),(-1,-1),11),
                    ('TEXTCOLOR',(0,0),(5,0),colors.white),
                    ]

if len(DataForID)>0:
    for i in range(len(DataForID)):
        if i>0:
            if DataForID[i][0]==DataForID[i-1][0]:
                IndelTableStyle.extend([('SPAN',(0,i),(0,i+1)),('VALIGN',(0,i),(0,i+1),'MIDDLE')])
            Name=Name+"、".decode('utf-8')+NameFI[i]
        else:
            Name=NameFI[i]
        for j in range(len(DataForPoint[i])):
            DataForID[i][j] = Paragraph("<font size=11 name='hei'>%s</font>"%DataForID[i][j],ParaStyle)
            
    IndelData.extend(DataForID)
    TitleForIndel = Paragraph("<font size=13 name='hei-Bold'>插入缺失(%s)</font>"%Name,SList, bulletText="\xe2\x80\xa2")
    IDName = Name
    TitleI = header([[TitleForIndel]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
    Point_Table = header(IndelData,IndelTableStyle, klass=Table,colWidths=[25*mm,40*mm,40*mm,20*mm,20*mm,40*mm])
    Page_Second.append(TitleI)
    Page_Second.append(Spacer(5.0*inch, 0.2*inch))
    Page_Second.append(Point_Table)
    Page_Second.append(Spacer(5.0*inch, 0.2*inch))

AmpData =  [
            ["基因","相对倍数","p值"]
            ]

AmpTableStyle = [
                    ('FONT',(0,0),(-1,-1),"hei"),
                    ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                    ('ALIGN',(0,0),(-1,-1),"LEFT"),
                    ('BACKGROUND',(0,0),(2,0),colors.HexColor('#9f29d9')),
                    ('FONTSIZE',(0,0),(-1,-1),11),
                    ('TEXTCOLOR',(0,0),(2,0),colors.white),
                    ]

if len(DataForAmp)>0:
    for i in range(len(DataForAmp)):
        if i>0:
            Name=Name+"、".decode('utf-8')+NameFA[i]
        else:
            Name=NameFA[i]
        for j in range(len(DataForAmp[i])):
            DataForAmp[i][j] = Paragraph("<font size=11 name='hei'>%s</font>"%DataForAmp[i][j],ParaStyle)
            
    AmpData.extend(DataForAmp)
    TitleForAmp = Paragraph("<font size=13 name='hei-Bold'>基因扩增(%s)</font>"%Name,SList, bulletText="\xe2\x80\xa2")
    TitleA = header([[TitleForAmp]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
    AmpName = Name
    Amp_Table = header(AmpData,AmpTableStyle, klass=Table,colWidths=[60*mm,60*mm,65*mm])
    Page_Second.append(TitleA)
    Page_Second.append(Spacer(5.0*inch, 0.2*inch))
    Page_Second.append(Amp_Table)
    Page_Second.append(Spacer(5.0*inch, 0.2*inch))

FuData =  [
              ["基因","3'/5'表达差异（Log值）","融合信号强度"]
              ]

FuTableStyle = [
                    ('FONT',(0,0),(-1,-1),"hei"),
                    ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                    ('ALIGN',(0,0),(-1,-1),"LEFT"),
                    ('BACKGROUND',(0,0),(2,0),colors.HexColor('#9f29d9')),
                    ('FONTSIZE',(0,0),(-1,-1),11),
                    ('TEXTCOLOR',(0,0),(2,0),colors.white),
                    ]
if len(DataForFu)>0:
    for i in range(len(DataForFu)):
        if i>0:
            Name=Name+"、".decode('utf-8')+NameFF[i]
        else:
            Name=NameFF[i]
        for j in range(len(DataForFu[i])):
            DataForFu[i][j] = Paragraph("<font size=11 name='hei'>%s</font>"%DataForFu[i][j],ParaStyle)
            
    FuData.extend(DataForFu)
    TitleForFu = Paragraph("<font size=13 name='hei-Bold'>基因融合(%s)</font>"%Name,SList, bulletText="\xe2\x80\xa2")
    TitleF = header([[TitleForFu]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
    FName = Name
    Fu_Table = header(FuData,FuTableStyle, klass=Table,colWidths=[60*mm,60*mm,65*mm])
    Page_Second.append(TitleF)
    Page_Second.append(Spacer(5.0*inch, 0.2*inch))
    Page_Second.append(Fu_Table)
    Page_Second.append(Spacer(5.0*inch, 0.2*inch))

#未发现小标题
SName = []
if len(DataForPoint)==0:
    SName.append("点突变")
if len(DataForID)==0:
    SName.append("插入缺失")
if len(DataForAmp)==0:
    SName.append("基因扩增")
if len(DataForFu)==0:
    SName.append("基因融合")
for i in range(len(SName)):
    if i==0:
        Name = SName[i]
    elif i==(len(SName)-1):
        Name = Name + "和".decode("utf-8")+SName[i]
    else:
        Name = Name + "、".decode("utf-8")+SName[i]
TitleForSN = Paragraph("<font size=13 name='hei-Bold'>%s</font>"%Name,SList, bulletText="\xe2\x80\xa2")
TitleSN = header([[TitleForSN]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
Page_Second.append(TitleSN)
Page_Second.append(Paragraph("<font size=13 name='hei'><b><br />未发现与靶向治疗相关的%s</b> </font>"%Name, body))

Page_Second.append(PageBreak())

#Index Page
Title10 = Paragraph("<font size=22 name='hei-Bold'>检测报告目录总览</font>",style=ParagraphStyle("Caption",alignment=TA_CENTER))
IndexTitle = header([[Title10]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])

Page_Index = [IndexTitle,Spacer(5.0*inch,0.8*inch)]

Page_Index.append(Paragraph("<font size=11 name='hei-Bold'><b>1.基因检测结果</b></font>",style=ParagraphStyle("Caption", alignment=TA_LEFT)))
Page_Index.append(Spacer(5.0*inch,0.2*inch))
if pointName!="":
    Page_Index.append(Paragraph("<font size=10 name='hei'>点突变(%s)</font>"%pointName,SList, bulletText="\xe2\x80\xa2"))
if IDName!="":
    Page_Index.append(Paragraph("<font size=10 name='hei'>插入缺失(%s)</font>"%IDName,SList, bulletText="\xe2\x80\xa2"))
if AmpName!="":
    Page_Index.append(Paragraph("<font size=10 name='hei'>基因扩增(%s)</font>"%AmpName,SList, bulletText="\xe2\x80\xa2"))
if FName!="":
    Page_Index.append(Paragraph("<font size=10 name='hei'>基因融合(%s)</font>"%FName,SList, bulletText="\xe2\x80\xa2"))

Page_Index.append(Spacer(5.0*inch,0.2*inch))
Page_Index.append(Paragraph("<font size=11 name='hei-Bold'><b>2.检测结果解读</b></font>",style=ParagraphStyle("Caption", alignment=TA_LEFT)))
Page_Index.append(Spacer(5.0*inch,0.2*inch))
Page_Index.append(Paragraph("<font size=11 name='hei-Bold'><b>3.靶向治疗相关的临床试验</b></font>",style=ParagraphStyle("Caption", alignment=TA_LEFT)))
Page_Index.append(Spacer(5.0*inch,0.2*inch))
Page_Index.append(Paragraph("<font size=11 name='hei'>参考文献</font>",style=ParagraphStyle("Caption", alignment=TA_LEFT)))
Page_Index.append(Spacer(5.0*inch,0.2*inch))
Page_Index.append(Paragraph("<font size=11 name='hei'><b>附录1：检测到的癌症靶向基因突变相关背景知识</b></font>",style=ParagraphStyle("Caption", alignment=TA_LEFT)))
Page_Index.append(Spacer(5.0*inch,0.2*inch))
Page_Index.append(Paragraph("<font size=11 name='hei'><b>附录2：与癌症治疗相关重要基因位点及关联药物列表</b></font>",style=ParagraphStyle("Caption", alignment=TA_LEFT)))
Page_Index.append(Spacer(5.0*inch,0.2*inch))
Page_Index.append(Paragraph("<font size=11 name='hei'><b>附表3：OncoGxSelect™涵盖的基因</b></font>",style=ParagraphStyle("Caption", alignment=TA_LEFT)))
Page_Index.append(Spacer(5.0*inch,0.2*inch))
Page_Index.append(Paragraph("<font size=11 name='hei'><b>检测说明</b></font>",style=ParagraphStyle("Caption", alignment=TA_LEFT)))

Page_Index.append(PageBreak())
elements.extend(Page_Index)
#IndexPage ends

elements.extend(Page_Second)


#Third page

#The title of the third page
Title31st = Paragraph("<font size=13 name='hei-Bold'>2.    检测结果解读</font>",style=ParagraphStyle("Caption", alignment=TA_LEFT))
Title31 = header([[Title31st]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
Page_Third = [Title31,Spacer(5.0*inch,0.2*inch)]
#increase table
IncreaseData = [
                ["敏感性可能增加的靶向药物"],
                ["基因突变","治疗方案","癌症类型","参考文献"],
                ]

IncreaseTableStyle = [
                ('TEXTCOLOR',(0,0),(0,0),colors.white),
                ('FONT',(0,0),(0,0),"hei"),
                ('FONT',(0,1),(-1,1),"hei"),
                ('BACKGROUND',(0,1),(-1,1),colors.HexColor('#404040')),
                ('TEXTCOLOR',(0,1),(-1,1),colors.HexColor('#ffc000')),
                ('ALIGN',(0,1),(-1,1),"CENTER"),
                ('ALIGN',(0,0),(0,0),"CENTER"),
                ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                ('BACKGROUND',(0,0),(0,0),colors.HexColor('#238943')),
                ('SPAN',(0,0),(3,0)),
                ('FONTSIZE',(0,0),(-1,-1),10.5),
                ]

#insert the data into the table
result["ClinBenefitsSame"].pop(0)
result["ClinBenefitsDiff"].pop(0)
if len(result["ClinBenefitsSame"][0])==1 and len(result["ClinBenefitsDiff"][0])==1:
    result["ClinBenefitsDiff"]=[]
elif len(result["ClinBenefitsDiff"][0])==1:
    result["ClinBenefitsDiff"]=[]
elif len(result["ClinBenefitsSame"][0])==1:
    result["ClinBenefitsSame"]=[]

DataForIncrease =sorted(result["ClinBenefitsSame"]+result["ClinBenefitsDiff"],key=lambda ge:ge[0])

i = 0
l = 0
if DataForIncrease!=[] and len(DataForIncrease[0])>1:
    for gene in (result["ClinBenefitsSame"]+result["ClinBenefitsDiff"]):
        if i>0 and (DataForIncrease[i][0]==DataForIncrease[i-1][0]) and (DataForIncrease[i][1]==DataForIncrease[i-1][1]):
            i+=1
            continue
        for j in range(len(DataForIncrease)-i):
            if (DataForIncrease[j+i][0]==DataForIncrease[i][0]) and (DataForIncrease[j+i][1]==DataForIncrease[i][1]):
                l+=1
        if l>1:
            IncreaseTableStyle.extend([('SPAN',(0,i+2),(0,i+l+1)),('VALIGN',(0,i+2),(0,i+l+1),'MIDDLE')])
        l = 0
        i+=1

i = 0
if DataForIncrease!=[] and len(DataForIncrease[0])>1:
    for gene in DataForIncrease:
        DataForIncrease[i][0]=Paragraph("<font size=10 name='hei'>%s</font>"%(DataForIncrease[i][0]+'-'+DataForIncrease[i][1]),ParaStyle)
        DataForIncrease[i][1]=Paragraph("<font size=10 name='hei'>%s</font>"%DataForIncrease[i][2],ParaStyle)
        DataForIncrease[i][2]=Paragraph("<font size=10 name='hei'>%s</font>"%DataForIncrease[i][3],ParaStyle)
        DataForIncrease[i][3]=Paragraph("<font size=10 name='hei'>%s</font>"%DataForIncrease[i][4],ParaStyle)
        DataForIncrease[i].pop(4)
        i+=1
else:
    DataForIncrease[0][0]=Paragraph("<font size=10 name='hei'>%s</font>"%DataForIncrease[0][0],ParaStyle)
    for i in range(3):
        DataForIncrease[0].append(Paragraph("<font size=10 name='hei'>\</font>",ParaStyle))

IncreaseData.extend(DataForIncrease)


Increase_Table = header(IncreaseData,IncreaseTableStyle, klass=Table,colWidths=[30*mm,75*mm,40*mm,40*mm])
Page_Third.append(Increase_Table)
Page_Third.append(Spacer(5.0*inch,0.2*inch))

DecreaseData = [
                ["敏感性可能降低的靶向药物"],
                ["基因突变","治疗方案","癌症类型","参考文献"],
                ]

DecreaseTableStyle = [
                ('TEXTCOLOR',(0,0),(0,0),colors.white),
                ('FONT',(0,0),(0,0),"hei"),
                ('FONT',(0,1),(-1,1),"hei"),
                ('ALIGN',(0,0),(0,0),"CENTER"),
                ('BACKGROUND',(0,1),(-1,1),colors.HexColor('#404040')),
                ('TEXTCOLOR',(0,1),(-1,1),colors.HexColor('#ffc000')),
                ('ALIGN',(0,1),(-1,1),"CENTER"),
                ('GRID',(0,0),(-1,-1),0.5,colors.grey),
                ('BACKGROUND',(0,0),(0,0),colors.HexColor('#d89234')),
                ('SPAN',(0,0),(3,0)),
                ('FONTSIZE',(0,0),(-1,-1),10.5)
                ]

result["LackClinBenefits"].pop(0)
DataForDiff = result["LackClinBenefits"]

i = 0
l = 0
if len(DataForDiff[0])>1:
    for gene in result["LackClinBenefits"]:
        if i>0 and (result["LackClinBenefits"][i][0]==result["LackClinBenefits"][i-1][0]) and (result["LackClinBenefits"][i][1]==result["LackClinBenefits"][i-1][1]):
            i+=1
            continue
        for j in range(len(result["LackClinBenefits"])-i):
            if (result["LackClinBenefits"][j+i][0]==result["LackClinBenefits"][i][0]) and (result["LackClinBenefits"][j+i][1]==result["LackClinBenefits"][i][1]):
                l+=1
        if l>1:
            DecreaseTableStyle.extend([('SPAN',(0,i+2),(0,i+l+1)),('VALIGN',(0,i+2),(0,i+l+1),'MIDDLE')])
        l = 0
        i+=1

i = 0
if len(DataForDiff[0])>1:
    for gene in result["LackClinBenefits"]:
        DataForDiff[i][0]=Paragraph("<font size=10 name='hei'>%s</font>"%(result["LackClinBenefits"][i][0]+'-'+result["LackClinBenefits"][i][1]),ParaStyle)
        DataForDiff[i][1]=Paragraph("<font size=10 name='hei'>%s</font>"%result["LackClinBenefits"][i][2],ParaStyle)
        DataForDiff[i][2]=Paragraph("<font size=10 name='hei'>%s</font>"%result["LackClinBenefits"][i][3],ParaStyle)
        DataForDiff[i][3]=Paragraph("<font size=10 name='hei'>%s</font>"%result["LackClinBenefits"][i][4],ParaStyle)
        DataForDiff[i].pop(4)
        i+=1
else:
    DataForDiff[0][0]=Paragraph("<font size=10 name='hei'>%s</font>"%result["LackClinBenefits"][0][0],ParaStyle)
    for i in range(3):
        DataForDiff[0].append(Paragraph("<font size=10 name='hei'>\</font>",ParaStyle))

DecreaseData.extend(DataForDiff)
Decrease_Table = header(DecreaseData,DecreaseTableStyle, klass=Table,colWidths=[30*mm,75*mm,40*mm,40*mm])
Page_Third.append(Decrease_Table)


Page_Third.append(PageBreak())
elements.extend(Page_Third)


#fourth page
Title41st = Paragraph("<font size = 13 name = 'hei-Bold'><b>3.靶向治疗相关的临床试验</b></font>",style=ParagraphStyle("Caption", alignment=TA_LEFT))
Title_CT = header([[Title41st]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[195*mm])
ClinTrials=[]
for Gene in Genes:
    firstCT=False
    #if Gene==Genes[0]:
    #    firstCT=True
    ClinTrialTable = stripe_table(result['ClinTrial'][Gene], "与 ".decode('utf-8')+Gene+" 相关的临床试验Trials".decode('utf-8'), '#472c77', firstCT, sep=0.2)
    ClinTrials.append(ClinTrialTable)
Page_Fourth = [Title_CT]+ClinTrials
#print(ClinTrials)
if len(Genes)==0:
    Page_Fourth.append(Paragraph("<font size=10 name='hei-Bold'><b><br />&nbsp&nbsp&nbsp&nbsp 未检测到与临床试验推荐相关的基因突变、基因融合及基因扩增 </b></font>",ParaStyle))
CT_annotation = Paragraph("<font size=10 name='hei'><i><br />* 注：仅显示正在招募患者的部分临床实验信息，可访问下面的网站寻找更多信息: www.ClinicalTrials.gov.</i></font>",style=ParagraphStyle("Normal", alignment=TA_LEFT))
Page_Fourth.extend([CT_annotation,PageBreak()])
elements.extend(Page_Fourth)

#ref page
myTitle8 = Paragraph("<font size=14 name='hei-Bold'>参考文献</font>",style=ParagraphStyle("Caption", alignment=TA_CENTER))
Title8 = header([[myTitle8]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])
Page_ref = [Title8,Spacer(5.0*inch, 0.2*inch)]
List = ParagraphStyle(name = 'Normal', leftIndent=10, bulletIndent=0, spaceBefore=2, fontSize = 9, leading = 14, wordWrap='CJK')
Color = '#1111ee'  # link color
Page_ref.extend([
      Paragraph("NCCN Biomarkers Compendium at: <font color='%s'><u>http://www.nccn.org/professionals/biomarkers/content/</u></font>"%Color, List, bulletText="\xe2\x80\xa2"),
      Paragraph("U.S. Food and Drug Administration, Table of Pharmacogenomic Biomarkers in Drug Labeling. Available online at: <br /><font color='%s'><u>http://www.fda.gov/Drugs/ScienceResearch/ResearchAreas/Pharmacogenetics/ucm083378.htm</u></font>"%Color, List, bulletText="\xe2\x80\xa2"),
      Paragraph("My Cancer Genome at: <font color='%s'><u>http://www.mycancergenome.org/</u></font>"%Color, List, bulletText="\xe2\x80\xa2"),
      Paragraph("Knowledge Base of Precision Oncology at: <font color='%s'><u>https://pct.mdanderson.org/</u></font>"%Color, List, bulletText="\xe2\x80\xa2"),
      Paragraph("Catalogue Of Somatic Mutations In Cancer (COSMIC) at: <font color='%s'><u>cancer.sanger.ac.uk</u></font>"%Color, List, bulletText="\xe2\x80\xa2"),
])

for ref in result["Reference"]:
    Page_ref.append(Paragraph(ref, List, bulletText="\xe2\x80\xa2"))

Page_ref.append(PageBreak())
elements.extend(Page_ref)

# Page About mutation information

n_Line=0 # record the Line number
myTitle5 = Paragraph("<font size=14 name='hei-Bold'>检测到突变的基因相关背景知识</font>",style=ParagraphStyle("Caption", alignment=TA_LEFT))
Title5 = header([[myTitle5]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[190*mm])

GeneInfo_Dic = {}
Prevelence_Dic = {}
GeneInfotrans_file = codecs.open(scriptFolder+"GeneInfo_translation.txt","r",'utf-8')
for j in GeneInfotrans_file.read().split("\n"):
    jsep = j.split("\t")
    GeneInfo_Dic.update({jsep[0]:jsep[2]})
GeneInfotrans_file.close()

for GeneInfo in result["GeneInfo"]:
    sentence = ""
    jsep = GeneInfo['Mutation prevalence'].split("|")
    for i in range(len(jsep)):
        Alteration = jsep[i][0:(jsep[i].find("mutation"))]
        disease = translate(jsep[i][(jsep[i].find("in")+2):jsep[i].find(":")])
        fre = jsep[i][(jsep[i].find(":")+1):len(jsep[i])]
        if i==0:
            sentence = Alteration + "在".decode("utf-8")+disease+"中的突变频率为".decode("utf-8")+fre
        else:
            sentence = sentence + "<br />" + Alteration + "在".decode("utf-8")+disease+"的所有".decode("utf-8")+GeneInfo['Gene']+"突变中的突变频率为".decode("utf-8")+fre
    Prevelence_Dic.update({GeneInfo['Gene']:sentence})

GeneInfo_string="<br />".join(["<font size=12 color='#0067b1'><b><i>%s</i></b></font><font size=10 name='hei'><br />%s<br /><br /><u>%s</u><br />%s<br /><br /><u>%s</u><br />%s<br /><br /><u>%s</u><br />%s<br /><br /></font>"%(GeneInfo['Gene'],GeneInfo_Dic[GeneInfo['Gene']],"突变发生的位置".decode('utf-8'),drugTranslation(GeneInfo['Mutation location in gene and/or protein']),"突变发生的频率".decode('utf-8'),Prevelence_Dic[GeneInfo['Gene']],"突变的影响".decode('utf-8'),drugTranslation(GeneInfo['Effect of mutation'].replace("|","<br />"))) for GeneInfo in result["GeneInfo"]])#'GeneInfo']])

if len(result["GeneInfo"])==0:#0 = 'GeneInfo'
    GeneInfo_string = "<font size=11 name='hei'><b><br />&nbsp&nbsp&nbsp&nbsp 未检测到相关基因突变、基因融合及基因扩增 </b></font>"
GeneInfo_box = Paragraph(GeneInfo_string, body)
Page_four = [Title5, Spacer(5.0*inch, 0.2*inch),GeneInfo_box,PageBreak()]
elements.extend(Page_four)

#  Information Page
myTitle7 = Paragraph("<font size=14 name='hei-Bold'>附表1: OncoGxSelectV2\xe2\x84\xa2 涵盖的基因</font>",style=ParagraphStyle("Caption", alignment=TA_LEFT))
Title7 = header([[myTitle7]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[190*mm])
Info_para = Paragraph("<font size=10 name='hei'>OncoGxSelectV2\xe2\x84\xa2 是基于扩增的癌症基因检测的Panel，提供灵敏、精确地检测方法来检测癌症中基因突变的频率，其涵盖了EGFR、 KRAS、BRAF、NRAS、ERBB2、ALK、PIK3CA、MAP2K1、KIT、MET、RET 和ROS1这12个基因。基于新一代测序的技术平台，检测这12个基因的点突变、插入缺失、基因扩增及基因融合，各基因对应的检测区域及突变类型见下表。</font>",body)

Info_file = open(scriptFolder+"/OncoGxSelectV2_Panel_info_CN2.txt",'r')
Data_info = [x.split("\t") for x in Info_file.read().split("\n") if len(x.split("\t"))>1]
tableStyle = [
        ('TEXTCOLOR',(0,0),(-1,0),colors.HexColor("#ffc000")),
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#404040")),
        ('GRID',(0,1),(-1,-1),0.5,colors.black),
        ('FONT',(0,1),(0,-1),"hei"),
        ('FONT',(1,1),(1,-1),"hei"),
        ('FONT',(2,1),(2,-1),"hei"),
        ('FONT',(0,0),(-1,0),"hei-Bold"),
        ('SPAN',(0,10),(0,13)),
        ('VALIGN',(0,10),(0,13),'MIDDLE'),
        ('SPAN',(0,14),(0,20)),
        ('VALIGN',(0,14),(0,20),'MIDDLE'),
        ('SPAN',(0,21),(0,25)),
        ('VALIGN',(0,21),(0,25),'MIDDLE'),
        ('SPAN',(2,11),(2,13)),
        ('VALIGN',(2,11),(2,13),'MIDDLE'),
        ('SPAN',(2,14),(2,20)),
        ('VALIGN',(2,14),(2,20),'MIDDLE'),
        ('SPAN',(2,22),(2,25)),
        ('VALIGN',(2,22),(2,25),'MIDDLE'),
]
Table_info = header(Data_info,style =tableStyle, klass=Table, sep=0.2, colWidths=[35*mm, 70*mm, 50*mm], rowHeights=[6.5*mm]*len(Data_info))
Page_info = [Title7, Spacer(5.0*inch, 0.2*inch), Info_para, Table_info]
elements.extend(Page_info)

#Second Additional Info
'''
myTitleA =  Paragraph("<font size=14 name='hei-Bold'>附表2: 与癌症治疗相关重要基因位点及关联药物</font>",style=ParagraphStyle("Caption", alignment=TA_LEFT))
TitleA = header([[myTitleA]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[190*mm])

Page_Addtional = [TitleA, Spacer(5.0*inch, 0.2*inch)]
'''
AdditionalData = [
                ["附表2: 与癌症治疗相关重要基因位点及关联药物"],
                [""],
                ["基因","突变形式","突变类型","关联药物"],
                ]

AdditionalTableStyle=[
                    ('VALIGN',(0,0),(0,0),'TOP'),
                    ('SPAN',(0,0),(3,0)),
                    ('FONT',(0,0),(0,0),'hei-Bold'),
                    ('FONTSIZE',(0,0),(0,0),13),
                    ('FONT',(0,2),(3,2),'hei'),
                    ('TEXTCOLOR',(0,2),(-1,2),colors.white),
                    ('BACKGROUND',(0,2),(-1,2),colors.HexColor("#426ab3")),
                    ('GRID',(0,2),(-1,-1),0.5,colors.black),
                    ]
                    
additional_File = scriptFolder+"/sites_and_drugs.txt"
action_fileA = codecs.open(additional_File,'r',encoding = 'utf-8')

dataForA=[]
x = 0
for j in action_fileA.read().split("\n"):
    if j.startswith("基因") or j.strip()=="":
        continue
    jsep = j.split("\t")
    dataForA.append([jsep[0],jsep[1],jsep[2],jsep[3]])
    x += 1

action_fileA.close()

l = 0
for i in range(len(dataForA)):
        if i>0 and (dataForA[i][0]==dataForA[i-1][0]):
            continue
        for j in range(len(dataForA)-i):
            if (dataForA[j+i][0]==dataForA[i][0]):
                l+=1
            else:
                break
        if l>1:
            AdditionalTableStyle.extend([('SPAN',(0,i+3),(0,i+l+2)),('VALIGN',(0,i+3),(0,i+l+2),'MIDDLE')])
            AdditionalTableStyle.extend([('SPAN',(3,i+3),(3,i+l+2)),('VALIGN',(3,i+3),(3,i+l+2),'MIDDLE')])
        l = 0

#print(dataForA)
for i in range(len(dataForA)):
    for j in range(4):
        dataForA[i][j] = Paragraph("<font size=10 name='hei'>%s</font>"%dataForA[i][j],ParaStyle)
    
AdditionalData.extend(dataForA)
Table_Additional = header(AdditionalData,AdditionalTableStyle, klass=Table,colWidths=[25*mm,50*mm,40*mm,70*mm])
Page_Addtional=[Table_Additional]
Page_Addtional.append(PageBreak())
elements.extend(Page_Addtional)

#last page
myTitle9 = Paragraph("<font size=14 name='hei-Bold'>检测说明</font>",style=ParagraphStyle("Caption", alignment=TA_CENTER))
Title9 = header([[myTitle9]],style =[("ALIGN",(0,0),(0,0),"LEFT")], klass=Table, colWidths=[193*mm])

Page_last = [Title9,Spacer(5.0*inch, 0.2*inch)]
Page_last.append(Paragraph("<font size=13 name='hei'><b><br />检测方法</b> </font>", body))
Page_last.append(Paragraph("<font size=10 name='hei'>基于新一代测序平台MiSeq（Illumina），OncoGxSelectV2\xe2\x84\xa2 对样品DNA和RNA获得的文库开展测序，其分析结果准确、可靠。OncoGxSelectV2\xe2\x84\xa2 选择第5节中列出的12个基因的部分区域，深度检测其点突变、插入缺失、基因融合或基因扩增信息。</font>", body))
Page_last.append(Paragraph("<font size=13 name='hei'><b><br />检测资质</b> </font>", body))
Page_last.append(Paragraph("<font size=10 name='hei'>OncoGxSelectV2\xe2\x84\xa2 基因突变检测服务由Admera Health开发，包括检测性能参数的确定和验证。由于现阶段并不需要经过FDA认可，故检测尚未取得FDA批准。OncoGxSelectV2\xe2\x84\xa2 检测适用于临床样品。Admera Health美国临床实验室已获得CLIA认可和CAP认证，具备开展高难度临床实验检测的资质。</font>", body))
Page_last.append(Paragraph("<font size=13 name='hei'><b><br />检测局限性</b> </font>", body))
Page_last.append(Paragraph("<font size=10 name='hei'>与所有实验方法一样，检测错误的可能性总是存在的。受样品来源和质量等综合因素影响，OncoGxSelectV2\xe2\x84\xa2 不能保证可检测到所有影响药效和药物安全性的突变，实验结果未发现某种突变并不排除病人实际存在变异表型的可能性；另外，药效和药物安全还可能受到非基因因素的影响，因此基因检测并不能替代临床和治疗性药物监测等检测手段。</font>", body))
Page_last.append(Paragraph("<font size=13 name='hei'><b><br />免责声明</b> </font>", body))
Page_last.append(Paragraph("<font size=10 name='hei'>报告中提供的信息是技术服务的一种形式，而非医学建议，<u>检测报告仅供科学研究使用</u>。本检测报告的内容基于出具日期时已出版的研究报道，不能反映其后新发表的数据及修订的用药方案。报告仅对其中的检测数据的准确性和客观真实性负责，但不作为其他任何保证的凭据，包括但不限于特定用途的商业性和适用性提示担保。临床医学建议必然是根据个例的实际情况量体裁制，因此，临床医生对患者的治疗负全部责任，这包括那些基于病人基因型信息而制定的方案。因此，Admera Health及其工作人员并不对任何个人或实体因使用这份报告而形成的直接或间接损失、损伤负责。</font>", body))

elements.extend(Page_last)


print(len(elements))
action_file.close()
doc.addPageTemplates([PageTemplate("First",[Frame(14*mm,17*mm,187*mm,240*mm)],onPage=myFirstPage),
                      PageTemplate("Later",[Frame(14*mm,17*mm,187*mm,228*mm)],onPage=myFirstPage),
                     ])

doc.build(elements, canvasmaker=NumberedCanvas)