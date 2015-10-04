__author__ = 'Shiva'
import openpyxl
class excel:

    def getSprintDetails(self):
        wb = openpyxl.load_workbook('Defect.xlsx')
        sheetnames = wb.get_sheet_names() # to get sheet names in the excel
        #print(sheetnames)

        sheet = wb.get_sheet_by_name(sheetnames[1])
        #print(sheet.title)

        last_row = sheet.max_row
        #print(last_row)
        sprintname =[]
        sprintstartdate= []
        sprintenddate=[]
        sprintdetail = []
        for i in range(2,last_row+1,1):
            sprintname.append(sheet.cell(row=i,column=1).value)
            sprintstartdate.append(sheet.cell(row = i,column = 2).value)
            sprintenddate.append(sheet.cell(row = i,column = 3).value)

        for i in range(len(sprintname)):
            sprintdetail.append([sprintname[i],sprintstartdate[i],sprintenddate[i]])
            pass

        #print(sprintdetail)
        #
        # sprintstartdate.sort()
        # sprintenddate.sort()
        # for i in range(len(sprintstartdate)):
        #     print('SPRINT NAME '+ str(sprintname[i]))
        #     print ('START DATE '+ str(sprintstartdate[i]))
        #     print ('END DATE '+ str(sprintenddate[i]))
        return (sprintdetail)

    def getDefectDetails(self):
        wb = openpyxl.load_workbook('Defect.xlsx')
        sheetnames = wb.get_sheet_names() # to get sheet names in the excel
        sheet = wb.get_sheet_by_name(sheetnames[0])
        last_row = sheet.max_row
        print(last_row)
        defect_date = []
        defect_priority = []
        for i in range(2,last_row+1):
            defect_date.append(sheet.cell(row=i, column=2).value)
            defect_priority.append(sheet.cell(row=i,column=3).value)
        #print(defect_date)
        # for figuring out defects in each month

        jan,feb,mar,apr,may,june,july,aug,sep,oct,nov,dec = 0,0,0,0,0,0,0,0,0,0,0,0
        #month_index = []
        for i in range(len(defect_date)):
            for month_range in range(1,13,1):
                if(defect_date[i].month == month_range):
                    if (month_range == 1):
                        jan +=1
                    elif (month_range ==2):
                        feb +=1
                    elif (month_range ==3):
                        mar +=1
                    elif (month_range ==4):
                        apr +=1
                    elif (month_range ==5):
                        may +=1
                    elif (month_range ==6):
                        june +=1
                    elif (month_range ==7):
                        july +=1
                    elif (month_range ==8):
                        aug +=1
                    elif (month_range ==9):
                        sep +=1
                    elif (month_range ==10):
                        oct +=1
                    elif (month_range ==11):
                        nov +=1
                    elif (month_range ==12):
                        dec +=1

        print(jan,feb,mar,apr,may,june,july,aug,sep,oct,nov,dec)

        #for defects based on priority
        p1,p2,p3 = 0,0,0
        for i in range(len(defect_priority)):
            for priority_range in range(1,4,1):
                if(defect_priority[i] == priority_range):
                    if (priority_range ==1):
                        p1 +=1
                    elif (priority_range ==2):
                        p2 +=1
                    elif (priority_range ==3):
                        p3 +=1
        print(p1,p2,p3)

    def getDefectInSprint(self):
        sprintdetails = excel.getSprintDetails(self)
        #print(len(sprintdetails))
        #print(sprintdetails)


        wb = openpyxl.load_workbook('Defect.xlsx')
        sheetnames = wb.get_sheet_names() # to get sheet names in the excel
        sheet = wb.get_sheet_by_name(sheetnames[0])
        last_row = sheet.max_row
        print(last_row)
        defect_date = []

        for i in range(2,last_row+1,1):
            defect_date.append(sheet.cell(row=i, column=2).value)

        print(len(defect_date))
        for i in range(len(defect_date)):
            #print('This is defect date '+str(defect_date[i]))
            for j in range(len(sprintdetails)):
                if(defect_date[i]>= sprintdetails[j][1] and defect_date[i]<=sprintdetails[j][2]):
                 #if (sprintdetails[j][1] >= defect_date[i] <= sprintdetails[j][2]):
                     print(sprintdetails[j][0])
            #print('----------------------------------------------')
        pass


if __name__ == '__main__':
    e = excel()
    #e.getSprintDetails()
    #e.getDefectDetails()
    e.getDefectInSprint()
