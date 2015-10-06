__author__ = 'Shiva'
import openpyxl
class excel:

    def getSprintDetails(self):
        wb = openpyxl.load_workbook('Defect.xlsx')
        sheetnames = wb.get_sheet_names() # to get sheet names in the excel
        sheet = wb.get_sheet_by_name(sheetnames[1])
        last_row = sheet.max_row #last row in the sheet
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
        return (sprintdetail)

    def getDefectDetails(self):
        wb = openpyxl.load_workbook('Defect.xlsx')
        sheetnames = wb.get_sheet_names() # to get sheet names in the excel
        sheet = wb.get_sheet_by_name(sheetnames[0])
        last_row = sheet.max_row
        print(last_row)
        defect_date = []
        defect_priority = []
        defect_deatil = []
        for i in range(2,last_row+1):
            defect_date.append(sheet.cell(row=i, column=2).value)
            defect_priority.append(sheet.cell(row=i,column=3).value)
        for i in range(len(defect_date)):
            defect_deatil.append([defect_date[i],defect_priority[i]])
            pass
        return (defect_deatil)

    def getDefectInSprint(self):
        sprintdetails = excel.getSprintDetails(self)
        defectdetails = excel.getDefectDetails(self)

        wb = openpyxl.load_workbook('Defect.xlsx')
        sheetnames = wb.get_sheet_names() # to get sheet names in the excel
        sheet = wb.get_sheet_by_name(sheetnames[0])
        last_row = sheet.max_row
        print(last_row)

        print(len(defectdetails))
        for i in range(len(defectdetails)):
            for j in range(len(sprintdetails)):
                if(defectdetails[i][0]>= sprintdetails[j][1] and defectdetails[i][0]<=sprintdetails[j][2]):
                     #print(sprintdetails[j][0])
                     value = [sprintdetails[j][0]]
                     #print(value)
                     defectdetails[i].extend(value)
        months = ['January', 'February', 'March', 'April', 'May', 'June', 'July','August', 'September', 'October', 'November', 'December']
        monthly_defect = []
        month_year_name = []
        defect_count_month = []
        priority_1 = []
        priority_2 = []
        priority_3 = []
        for i in range(len(defectdetails)):
            for year in range(2015,2017,1):
                for month_val in range(1,13,1):
                    if defectdetails[i][0].month == month_val and defectdetails[i][0].year == year:
                        month_year_val = [str((str(months[month_val-1]) + ' '+str(year)))]
                        defectdetails[i].extend(month_year_val)
                        pass
        excel.writeinsheet(self,defectdetails)
        for i in range(len(defectdetails)):
            #print(defectdetails[i])
            pass
        print('----------------------------------------------------------------------------------------------')

        for year in range(2015,2017,1):
            for month_val in range(1,13,1):
                value_to_search = str((str(months[month_val-1]) + ' '+str(year)))
                count =0
                p1,p2,p3=0,0,0
                print('-------')
                for i in range(len(defectdetails)):
                    if value_to_search in defectdetails[i]:
                        count +=1
                        print(count)
                        if defectdetails[i][1] ==1:
                            p1 +=1
                        elif defectdetails[i][1] ==2:
                            p2+=1
                        elif defectdetails[i][1] ==3:
                            p3+=1
                        else:
                            pass
                if count>0:
                    month_year_name.append(value_to_search)
                    defect_count_month.append(count)
                if p1 >=0 or p2>=0 or p3>=0:
                    priority_1.append(p1)
                    priority_2.append(p2)
                    priority_3.append(p3)
            pass

        for i in range(len(month_year_name)):
            monthly_defect.append([month_year_name[i],defect_count_month[i],priority_1[i],priority_2[i],priority_3[i]])

        for i in range(len(monthly_defect)):
            pass
            print(monthly_defect[i])
        excel.writeinsheet(self,monthly_defect)

        #For counting defects monthly based on priority




    def writeinsheet(self,anylist):
        anylist = anylist
        wb = openpyxl.load_workbook('Defect.xlsx')
        sheetnames = wb.get_sheet_names() # to get sheet names in the excel
        sheet = wb.get_sheet_by_name(sheetnames[2])
        wb.remove_sheet(sheet)

        wb.create_sheet(2,title='Temp')
        wb.save('Defect.xlsx')
        sheet = wb.get_sheet_by_name('Temp')

        for rows in range(len(anylist)):
            for columns in range(len(anylist[rows])):
                sheet.cell(row=rows+1,column=columns+1, value=str(anylist[rows][columns]))
                print(anylist[rows][columns])
        wb.save('Defect.xlsx')


if __name__ == '__main__':
    e = excel()
    #e.getSprintDetails()
    #e.getDefectDetails()
    e.getDefectInSprint()
