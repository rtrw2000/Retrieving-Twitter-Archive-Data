# from Interval import Interval
from Profile import Profile
# from Creds import Creds
class Row():

    # def __init__(self):
    #     # self.splitter = splitter
    #     # self.splitter.getAllNameFromSheet()

    # def getRowByName(self, name):
    #     x = self.splitter.nameDictionary[name]
    #     print(x)




    def getAllProfileListFromRow(self,splitter):
        print("Getting all profile list from excel..")
        profileList = []
        for key, value in splitter.nameDictionary.items():
            profile = Profile(key)
            profileList.append(profile)
        return profileList


