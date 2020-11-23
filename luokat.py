# -*- coding: utf-8 -*-
"""
Created on Sat Sep 12 10:21:38 2020

@author: rikus
"""



import pandas as pd
import datetime
import numpy as np
import operator
import tkinter as tk
from PIL import Image, ImageTk # PIL is the Python Imaging Library
from itertools import count



class ImageLabel(tk.Label):
    """
    Luo tkinter Labelin, joka ottaa kaikki normaalit parametrit.
    Jos parametriksi antaa kuvan, joka on animaatio, toistaa animaation kerran.
    Luokka lainattu täältä: https://stackoverflow.com/questions/43770847/play-an-animated-gif-in-python-with-tkinter
    """
    def load(self, im):
        if isinstance(im, str):
            im = Image.open(im) # Image is a class which is used to represent PIL images
        self.loc = 0
        self.frames =  []
        self.last_frame = 0
        
        try:
            for i in count(1): # count(x) is an infinite iterator from itertools module
                self.frames.append(ImageTk.PhotoImage(im.copy()))
                im.seek(i)
        except EOFError:
            self.last_frame = len(self.frames)-1
        
        try:
#            self.delay = im.info["duration"]
            self.delay = 15 # vaihdettu success_animation.gif varten
        except:
            self.delay = 100
            
        if len(self.frames) == 1:
            self.config(image=self.frames[0])
        else:
            self.next_frame()
            
    def unload(self):
        self.config(image="")
        self.frames = None
        
    def next_frame(self):
        if self.frames: # if frames not empty
            if self.loc == self.last_frame:
                self.config(image=self.frames[0])
            else:
                self.loc += 1
                self.loc %= len(self.frames)
                self.config(image=self.frames[self.loc])
                self.after(self.delay, self.next_frame)



class Report(object):
    """
    Pohjaluokka Proteukselle ja SAP :ille, johon ei välttämättä tule mitään.
    """
    pass



class Proteus(Report):
    
    def __init__(self, file_location, day):
        """
        file_location: Proteus-raportin (CSV) lokaatio levyllä.
        day: kokonaisluku; halutaanko raportti tältä päivältä (0), huomiselta (1)
        vai ylihuomiselta (2)
        """
        assert type(file_location) == str, "__init__ argument file_location is faulty"
        assert type(day) == int and 0 <= day <= 2, "__init__ argument day is faulty"
        
        # Asettaa Pandasin näyttämään kaikki DF-objektin pylväät konsolissa
        pd.set_option('display.expand_frame_repr', False)
        
        self.day = None
        
        self.file_original = None # alkuperäinen versio Proteus-raportista (CSV)
        self.file_edit = None # muokattu versio Proteus-raportista (DF)
        self.file_PT_tilaukset = None # pivot table, joka sisältää kaikki halutun päivän tilaukset (DF)
        self.file_PT_virheet = None # pivot table, joka sisältää virhe-tilassa olevat tilaukset (DF)
        
        self.set_file(file_location)
        self.fix_column_P()
        self.remove_wrong_dates(day)
        self.create_PT_tilaukset()
        self.create_PT_virheet()


    def get_date(self, day):
        """
        Palauttaa tämän, huomisen tai ylihuomisen päivämäärän (str) esim. "2020-03-09"
        day: kokonaisluku 0-2; tänään huomenna tai ylihuomenna
        """
        assert type(day) == int and 0 <= day <= 2, "get_date day argument is faulty"
        now = datetime.datetime.now()
        delta = datetime.timedelta(day)
        temp = now + delta
        date = str(temp)[0:10]
#        print("date1: " + str(date))
        assert type(date) == str and len(date) == 10, "get_date return statement date is faulty"
        return date
#        return "2020-03-09" # testaamista varten


    def get_time(self):
        """
        Palauttaa ajan tunnit-minuutit-sekunnit (str)
        """
        time = str(datetime.datetime.now())[11:-7]
        time = time.replace(":", "-")
        assert type(time) == str and len(time) == 8, "get_time return statement time is faulty"
        return time


    def set_file(self, file_location):
        """
        Hakee tiedoston tiedostopolusta file_location, siivoaa sitä hieman,
        ja tallentaa siitä DF-kopion kahteen muuttujaan: file_original, file_edit
        """
        assert type(file_location) == str, "set_file file_location argument is faulty"
        
        temp_DF = pd.read_csv(file_location, encoding='cp1252', sep=";")
        assert type(temp_DF) == pd.core.frame.DataFrame, "set_file variable temp_DF is faulty"
        
        # Tehdään uusi DF-objekti vain halutuilla pylväillä
        temp_DF = temp_DF[["PM tilnro", "Tila O - L", "Jono", "Toim.os nimi2", "Til. Tyyppi", 
                         "Toim. Tapa", "Qq", "Toiv. Toim pvm", "Tuote", "Kuvaus", "Tilattu "]]
        
        # Korjataan yksi hassusti nimetty header
        temp_DF = temp_DF.rename(columns = {"Tilattu ":"Tilattu"})
        
        # Muutetaan päivämäärät datetime-objekteiksi
        temp_DF["Toiv. Toim pvm"] = pd.to_datetime(
                temp_DF["Toiv. Toim pvm"], 
                infer_datetime_format = True, 
                dayfirst=True)
        
        self.file_original = temp_DF.copy(deep=True)
        self.file_edit = temp_DF.copy(deep=True)
        
        
    def fix_column_P(self):
        """
        Korjaa DF-objektin file_edit P-pylvään kolmeen osaan merkin "-" perusteella.
        Nimeää pylväät column_1, column_2 ja column_3.
        """
        
        column_1 = "Material" # materiaalikoodi esim. 9061007
        column_2 = "Plant" # esim. FIBL, FIBQ, FIBS
        column_3 = "Sloc" # esim. VRAH, VPAL, VOUL
        
        # new data frame with split value columns 
        temp_DF = self.file_edit["Tuote"].str.split("-", n = -1, expand = True) 
          
        # making separate "Material" column from new data frame 
        self.file_edit[column_1]= temp_DF[0] 
          
        # making separate "Plant" column from new data frame 
        self.file_edit[column_2]= temp_DF[1] 
        
        # making separate "Sloc" column from new data frame 
        self.file_edit[column_3]= temp_DF[2] 
        
        # Dropping old "Tuote" column 
        self.file_edit.drop(columns =["Tuote"], inplace = True) 
        
        # drops all rows containing at least one field with missing data
        self.file_edit = self.file_edit.dropna()
        
        
    def remove_wrong_dates(self, day):
        """
        Poistaa väärän pvm :n rivit DF-objektista file_edit.
        """
        assert type(day) == int and 0 <= day <= 2, "remove_wrong_dates argument day is faulty"
        
        date = self.get_date(day)
        
        # Get names of indexes for which column "Toiv. Toim pvm" does not have value "2020-03-06"
        indexNames = self.file_edit[ self.file_edit["Toiv. Toim pvm"] != date ].index
        
        # Delete these row indexes from dataFrame
        self.file_edit.drop(indexNames , inplace=True)
        
    
    def create_PT_tilaukset(self):
        """
        Palauttaa DF-objektista report_edit tehdyn pivot tablen.
        Kyseinen pivot table näyttää päivän työjonon niin, että sieltä on poistettu
        johdot, moduulit yms.
        """
        
        temp_DF = self.file_edit.copy(deep=True)
        assert type(temp_DF) == pd.core.frame.DataFrame, "create_PT_tilaukset variable temp_DF is faulty"
        
        # Get names of indexes for which column "Toiv. Toim pvm" does not have value "2020-03-06"
        indexNames1 = temp_DF[ temp_DF["Sloc"] == "VADS" ].index
        indexNames2 = temp_DF[ temp_DF["Sloc"] == "ASFP" ].index
        
        # Delete these row indexes from dataFrame
        temp_DF.drop(indexNames1 , inplace=True)
        temp_DF.drop(indexNames2 , inplace=True)
        
        temp_DF = temp_DF[["Toim.os nimi2", "Sloc", "Kuvaus", "Tilattu"]]
        self.file_PT_tilaukset = temp_DF.pivot_table(
                index=["Kuvaus", "Sloc", "Toim.os nimi2"], 
                aggfunc=np.sum)
        print(self.file_PT_tilaukset)
        
        
    def create_PT_virheet(self):
        """
        Palauttaa DF-objektista file_original tehdyn pivot tablen.
        Kyseinen pivot table näyttää kaikki O OSA- ja O POI-tilassa olevat
        rivit.
        """
        
        indexNames = self.file_original[ self.file_original["Jono"] == " " ].index
        temp_DF = self.file_original.copy(deep=True)
        assert type(temp_DF) == pd.core.frame.DataFrame, "create_PT_virheet variable temp_DF is faulty"
        temp_DF.drop(indexNames , inplace=True)
        temp_DF = temp_DF[["PM tilnro", "Tila O - L", "Jono", "Toim.os nimi2", "Toim. Tapa", 
                        "Toiv. Toim pvm", "Tuote", "Kuvaus", "Tilattu"]]
        self.file_PT_virheet = temp_DF.pivot_table(
                index=["Toiv. Toim pvm", "Toim.os nimi2", "PM tilnro", "Tuote", "Kuvaus"], 
                columns=["Jono"], 
                aggfunc=np.sum)
        print(self.file_PT_virheet)
        
        
    def get_file_original(self):
        assert type(self.file_original) == pd.core.frame.DataFrame, "get_file_original return statement is faulty"
        return self.file_original   
    
    
    def get_file_edit(self):
        assert type(self.file_edit) == pd.core.frame.DataFrame, "get_file_edit return statement is faulty"
        return self.file_edit
    
    
    def get_file_PT_tilaukset(self):
        assert type(self.file_PT_tilaukset) == pd.core.frame.DataFrame, "get_file_PT_tilaukset return statement is faulty"
        return self.file_PT_tilaukset
    
    
    def get_file_PT_virheet(self):
        assert type(self.file_PT_virheet) == pd.core.frame.DataFrame, "get_file_PT_virheet return statement is faulty"
        return self.file_PT_virheet



class Sap(Report):

    def __init__(self, työjono_location):
        """
        työjono_location: tiedostopolku SAP :in työjonoraporttiin (xlsx)
        
        Ohjelman kotikansiossa pitää aina olla tiedosto "Työtuotteiden ajat.xlsx",
        josta noudetaan ajat ja tunnukset.
        
        ajat: excel (sheet1) jossa on määritetty työtuotteiden kestot ja mitkä työtuotteet halutaan
        mukaan tämän algoritmin laskentaan.
        
        tunnukset: excel (sheet2) jossa on määritetty minkä SAP-tunnuksien tekemät
        tallennukset otetaan mukaan tämän algoritmin laskentaan.
        """
        assert type(työjono_location) == str, "__init__ argument työjono_location is faulty"
        
        self.työjono = pd.read_excel(työjono_location)
        self.ajat = pd.read_excel("Työtuotteiden ajat.xlsx", sheet_name="Sheet1")
        self.tunnukset = pd.read_excel("Työtuotteiden ajat.xlsx", sheet_name="Sheet2")
        self.frequencies = self.retrieve_frequencies(self.työjono, self.ajat, self.tunnukset)
        self.total_time = self.calculate_total_time(self.frequencies, self.ajat)
        self.lines = self.minutes_per_twenty(self.total_time)


    def retrieve_frequencies(self, työjono, ajat, tunnukset):
        """
        Tekee dictin "päivän työjono.xlsx" -tiedoston pylväiden solupareista "Material":"Order Quantity".
        Jos Materiaali esiintyy useammin kuin kerran, incrementoi "Order Quantity" verran.
        Hyväksyy materiaalin joko ehdolla 1) käyttäjänimi ja materiaali on hyväksytty "x"-kirjaimella
        tai 2) materiaali on pakotettu "a"-kirjaimella.
        """
        assert type(työjono) == pd.core.frame.DataFrame, "retrieve frequencies argument työjono is faulty"
        assert type(ajat) == pd.core.frame.DataFrame, "retrieve frequencies argument ajat is faulty"
        assert type(tunnukset) == pd.core.frame.DataFrame, "retrieve frequencies argument tunnukset is faulty"
        
        usernames = self.retrieve_usernames(tunnukset)
        materials_X = self.retrieve_materials_x(ajat)
        materials_A = self.retrieve_materials_a(ajat)
        frequencies = {}
        length = työjono.shape[0]
        for n in range(length):
            try:
                material = työjono["Material"][n]
                username = työjono["Created By"][n]
                quantity = työjono["Order Quantity"][n]
                if material in frequencies and material in materials_X and username in usernames or material in frequencies and material in materials_A:
                    frequencies[material] += quantity
                elif material not in frequencies and material in materials_X and username in usernames or material not in frequencies and material in materials_A:
                    frequencies[material] = quantity
            except:
                material = työjono["Material"][n]
                username = työjono["Created By"][n]
                quantity = työjono["Delivery quantity"][n]
                if material in frequencies and material in materials_X and username in usernames or material in frequencies and material in materials_A:
                    frequencies[material] += quantity
                elif material not in frequencies and material in materials_X and username in usernames or material not in frequencies and material in materials_A:
                    frequencies[material] = quantity
                    
        assert type(frequencies) == dict, "retrieve frequencies return statement is faulty"
        return frequencies
    
    
    def retrieve_usernames(self, tunnukset):
        """
        Noutaa "Työtuotteiden ajat.xlsx" -tiedoston Sheet2 :sta kaikki ne tunnukset,
        joidenka vierestä on ruksattu "Seurataan" -solu.
        Palauttaa listan.
        """
        assert type(tunnukset) == pd.core.frame.DataFrame, "retrieve usernames argument tunnukset is faulty"
        
        usernames = []
        length = tunnukset.shape[0]
        for n in range(length):
            if tunnukset["Seurataan"][n] == "x":
                usernames.append(tunnukset["Tunnus"][n])
                
        assert type(usernames) == list, "retrieve_usernames return statement is faulty"
        return usernames


    def retrieve_materials_a(self, ajat):
        """
        apufunktio: retrieve_frequencies
        """
        assert type(ajat) == pd.core.frame.DataFrame, "retrieve_materials_a argument ajat is faulty"
        
        picks = []
        length = ajat.shape[0]
        for n in range(length):
            if ajat["Seurataan"][n] == "a":
                picks.append(ajat["SAP no"][n])
        
        assert type(picks) == list, "retrieve_materials_a return statement is faulty"
        return picks
        
        
    def retrieve_materials_x(self, ajat):
        """
        apufunktio: retrieve_frequencies
        """
        assert type(ajat) == pd.core.frame.DataFrame, "retrieve_materials_x argument ajat is faulty"
        
        picks = []
        length = ajat.shape[0]
        for n in range(length):
            if ajat["Seurataan"][n] == "x":
                picks.append(ajat["SAP no"][n])
        
        assert type(picks) == list, "retrieve_materials_x return statement is faulty"
        return picks
        
        
    def calculate_total_time(self, frequencies, ajat):
        """
        Ottaa dictin jossa key-value on material-frequency.
        Kertoo kaikki frequencyt asennusajalla ja incrementoi tulot palautettavaan
        muuttujaan TotalTime
        """
        assert type(frequencies) == dict, "calculate_total_time argument frequencies is faulty"
        assert type(ajat) == pd.core.frame.DataFrame, "calculate_total_time argument ajat is faulty"
        
        temp = frequencies.copy()
        times = self.retrieve_times(ajat)
        totalTimes = {}
        totalTime = 0
        for material in temp:
            totalTimes[material] = temp[material] * times[material]
        for value in totalTimes.values():
            totalTime += value
        
        totalTime = int(totalTime)
        assert type(totalTime) == int, "calculate_total_time return statement is faulty"
        return totalTime
        
        
    def retrieve_times(self, ajat):
        """
        Tekee dictin "Työtuotteiden ajat.xlsx" -tiedoston pylväiden solupareista "SAP no":"Kesto".
        """
        assert type(ajat) == pd.core.frame.DataFrame, "retrieve_times argument ajat is faulty"
        
        times = {}
        length = ajat.shape[0]
        for n in range(length):
            material = ajat["SAP no"][n]
            time = ajat["Kesto"][n]
            times[material] = time
            
        assert type(times) == dict, "retrieve_times return statement is faulty"
        return times
        
        
    def minutes_to_hours(self, x):
        """
        Ottaa luvun (x)
        Palauttaa stringin muodossa <tunnit:minuutit:sekunnit>
        """
        assert type(x) == int or type(x) == float, "minutes_to_hours argument x is faulty"
        
        minutes = int(x % 60)
        hours = int(x // 60)
        temp = str(hours).zfill(2) + ":" + str(minutes).zfill(2) + ":00"
        
        assert type(temp) == str and len(temp) == 8, "minutes_to_hours return statement is faulty"
        return temp
        
        
    def minutes_per_twenty(self, x):
        """
        Ottaa kokonaisluvun x (minuutit), jakaa sen 1-10 esiasentajan kesken,
        muuttaa minuutit tunneiksi ja palauttaa tulokset listana.
        """
        assert type(x) == int, "minutes_per_twenty argument x is faulty"
        
        aList = []
        for n in range(2,22,2):
            temp = "Työjono per " + str(int(n/2)) + " esiasentaja(a)" + ": " + str(self.minutes_to_hours(x/n))
            aList.append(temp)
        note = "\nHuom. Tunnit on laskettu oletuksella, että\njokainen esiasentaja asentaa kerrallaan kaksi\nlaitetta ts. käyttää kahta konsolipiuhaa. Jos\nhalutaan veloitettavat tunnit, pitää halutun\nrivin tunnit kertoa vielä nyt kahdella."
        aList.append(note)
        
        assert type(aList) == list, "minutes_per_twenty return statement is faulty"
        return aList


    def get_työjono(self):
        assert type(self.työjono) == pd.core.frame.DataFrame, "get_työjono return statement is faulty"
        return self.työjono
    
    
    def get_ajat(self):
        assert type(self.ajat) == pd.core.frame.DataFrame, "get_ajat return statement is faulty"
        return self.ajat
    
    
    def get_tunnukset(self):
        assert type(self.tunnukset) == pd.core.frame.DataFrame, "get_tunnukset return statement is faulty"
        return self.tunnukset
    
    
    def get_frequencies(self):
        assert type(self.frequencies) == dict, "get_frequencies return statement is faulty"
        return self.frequencies
    
    
    def get_total_time(self):
        assert type(self.total_time) == int, "get_total_time return statement is faulty"
        return self.total_time
    
    
    def get_times(self):
        assert type(self.lines) == list, "get_times return statement is faulty"
        return self.lines



class Teraterm(object):
    
    def __init__(self, konffit):
        """
        konffit: konffit, jotka halutaan muuttaa TeraTerm-scriptin lukemaan muotoon (str)
        """
        self.konffit_copypaste = konffit
        self.konffit_teraterm = None
        self.convertConfig(konffit)
    
    
    def listToStr(self, inputStr):
        """
        Ottaa konffit stringinä.
        Palauttaa rivit listana.
        """
        outputList = []
        temp = inputStr.split("\n")
        for line in temp:
            if len(line) != 0:
                outputList.append(line)
        #print(outputList)
        return outputList


    def makeList(self, x):
        """
        input: string
        output: lista (stringin jokainen merkki listan objektina)
        """
    #    print("makeList:" + str(x))
        aList = []
        for n in range(len(x)):
            aList.append(x[n])
        return aList


    def removeStuff(self, x):
        """
        input: lista
        output: palauttaa listan, josta poistettu edestä ja lopusta välilyönnit.
        """
    #    print("removeStuff: " + str(x))
        aList = x.copy()
        while True:
            if aList[0] == " ":
                del aList[0]
            else:
                break
        while True:
            if aList[-1] == " ":
                del aList[-1]
            else:
                break
        return aList


    def isExclamation(self, x):
        """
        input: lista
        output: True, jos lista on huutomerkki; False muuten
        """
    #    print("isExclamation: " + str(x))
        aList = x.copy()
        if len(aList) <= 3 and "!" in aList:
            return True
        else:
            return False


    def addStuff(self, x):
        """
        input: lista
        output: lista, johon lisätty alkuun ja loppuun halutut merkit
        """
    #    print("addStuff " + str(x))
        aList = []
        start = "sendln '"
        middle = x.copy()
        end = "'"
        for char in start:
            aList.append(char)
        for char in middle:
            aList.append(char)
        for char in end:
            aList.append(char)
        return aList


    def makeString(self, x):
        """
        input: lista
        output: string (listan jokainen objekti yhdistettynä stringiksi)
        """
    #    print("makeString: " + str())
        aString = ""
        aList = x.copy()
        aString = "".join(str(x) for x in aList)
        return aString


    def convertConfig(self, inputStr):
        """
        Muuttaa konffit.txt sisällön TeraTermin lukemaan muotoon.
        """
        text = self.listToStr(inputStr)
        output = ""
        for line in text:
            temp = self.makeList(line)
            temp2 = self.removeStuff(temp)
            if self.isExclamation(temp2) == False:
                temp3 = self.addStuff(temp2)
                temp4 = self.makeString(temp3)
                output = output + temp4 + "\n"
        self.konffit_teraterm = output
    
    
    def get_konffit_copypaste(self):
        return self.konffit_copypaste
    
    
    def get_konffit_teraterm(self):
        return self.konffit_teraterm

