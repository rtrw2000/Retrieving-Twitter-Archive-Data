class Column():
    headerList = []
    dictHeader = {}


    def createColumnHeader(self,profile):
        self.dictHeader.clear()
        self.dictHeader["Name"]=""
        self.dictHeader["Created At"]=""
        self.dictHeader["Profile Url"]=""

        self.dictHeader["Latitude"]=""
        self.dictHeader["Longitude"]=""
        self.dictHeader["Google Maps"]=""
        self.dictHeader["Tweet"]=""



    def getColumnHeader(self):
        return self.headerList

