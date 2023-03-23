# from Interval import Interval


class Profile():
    def __init__(self, username):
        self.username = username

    def setListTweet(self,listTweet):
        self.lisTweet = listTweet

    def setListCreated(self,createdAt):
        self.created_at = createdAt

    def getListCreated(self):
        return self.created_at

    def setListLatitude(self,listLatitude):
        self.listLatitude = listLatitude

    def getListLatitude(self):
        return self.listLatitude

    def setListLongitude(self,listLongitude):
        self.listLongitude = listLongitude

    def getListLongitude(self):
        return self.listLongitude

    def getListTweeet(self):
        return self.lisTweet


    def setListTime(self, listTime):
        self.listTime = listTime

    def setListPlusCode(self, listPlusCode):
        self.listPlusCode = listPlusCode

    def setCreds(self, creds):
        self.creds = creds

    def setPlusCodeGroup(self, plusCodeGroup):
        self.plusCodeGroup = plusCodeGroup

    def getPlusCodeGroupSize(self):
        return len(self.plusCodeGroup)

    def getPlusCodeGroup(self):
        return self.plusCodeGroup

    def setInterval(self, Interval):
        self.interval = Interval

    def setTweet(self,Tweet):
        self.tweet = Tweet

    def getTweet(self):
        return self.tweet

    def setProfileUrl(self,url):
        self.url = url

    def getProfileUrl(self):
        return self.url

    def setLat(self,lat):
        self.lat = lat

    def getLat(self):
        return self.lat

    def setMaps(self,maps):
        self.maps = maps

    def getMaps(self):
        return self.maps

    def setLongitude(self,longitude):
        self.longitude = longitude

    def getLongitude(self):
        return self.longitude

    def getInterval(self):
        sorted_interval = dict(sorted(self.interval.items()))

        for key, value in sorted_interval.items():
            print("KEY INTERVAL " +self.username+" "+ str(key))
        return self.interval


    def mappingCredsToInterval(self):
        self.sorted_creds = dict(sorted(self.creds.items()))
        self.sorted_interval = dict(sorted(self.interval.items()))
        for key in self.sorted_interval.keys():
            credsList = []
            for value2 in self.sorted_creds.values():
                if(value2.interval == str(key)):
                    credsList.append(value2)

            intervalObj = Interval(key,credsList,self.plusCodeGroup)
            self.sorted_interval[key] = intervalObj







    def getMappingCredsInterval(self):
        for key,value in self.sorted_interval.items():
            print("Interval adalah 2 " +" "+ self.username + " " + str(key) + " " + str(value.id)+" "+str(len(value.creds)) )









    def hitungSebaran(self):
        self.dictSebaran = {}
        sorted_interval = dict(sorted(self.creds.items()))

        # sorted_interval = dict(sorted(self.interval.items()))
        for key, value in sorted_interval.items():
            print("Hell yeah dari " + self.username + " " + str(key) + " sebaran" + " " + str(value.sebaran))

