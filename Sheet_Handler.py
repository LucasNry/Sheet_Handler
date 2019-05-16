import xlsxwriter

class SheetHandler:

    name = ""
    paramcontainer = []
    row = 1
    col = 1
    def __init__(self):
        self.name = input("What's the name of the file(exclude file format.Ex:.xlsx)?: ") + ".xlsx"
        self.paramcontainer = self.get_param()

    def isGettingInput (self): # Needs Rework
        answer = input("Is there more data?(y/n) ")
        isGettInput = True
        if answer.lower() == "n" :
            isGettInput = False
            return isGettInput
        elif answer.lower() == "y":
            isGettInput = True
            return isGettInput

    def get_param(self):
        paramContainer = []
        counter = 1
        print("What parameters are gonna be stored in the sheet?")
        while self.isGettingInput():
            paramet = input("Param. " + str(counter) + ": ")
            paramContainer.append(paramet)
            counter += 1
        return paramContainer;

    def get_data(self):
        data = []
        dataInterm = []
        print("Provide the data: ")
        while self.isGettingInput():
            for param in self.paramcontainer:
                dataInterm.append(input(param + ": "))
            data.append(dataInterm)
            dataInterm = []

        return data;

    def create_and_write_sheet (self):
        workbook = xlsxwriter.Workbook(self.name)
        worksheet = workbook.add_worksheet()
        data = self.get_data()

        for param in self.paramcontainer:
            worksheet.write(0,self.col,param)
            self.col += 1
        self.col = 1
        for array in data:
            for item in array:
                worksheet.write(self.row,self.col,item)
                self.col += 1
            self.col = 1
            self.row += 1

        workbook.close()

    def add_to_sheet (self, name): #Hasn't been tested
        workbook = xlsxwriter.Workbook(name + ".xlsx")
        worksheet = workbook.get_worksheet_by_name(self.name)
        data = self.get_data()
        for array in data:
            for item in array:
                worksheet.write(self.row, self.col, item)
                self.col += 1
            self.col = 1
            self.row += 1






