import xlrd
from threading import Thread
import time


def recuperation_serveur():
    wb = xlrd.open_workbook('config.xlsx')
    sh = wb.sheet_by_name(u'Configuration_OPC')
    for rownum in range(sh.nrows):
        print(sh.row_values(rownum))

    colonne2 = sh.col_values(1)
    # print("OPC : ",colonne2[0])
    # print("Noeud : ",colonne2[1])
    serveur = colonne2[0]
    noeud = colonne2[1]

    return serveur, noeud


def recuperation_equipement():
    wb = xlrd.open_workbook('config.xlsx')
    liste_feuille = wb.sheet_names()
    del liste_feuille[0]
    print(liste_feuille)
    compteur = 0
    equipement = list()


    for feuille in liste_feuille:
        sh = wb.sheet_by_name(feuille)
        equipement.append(sh.col_values(0))
        del equipement[compteur][0]

        compteur += 1
    return equipement



class surveillance_donnees(Thread):
    def __init__(self, liste_thread):
        Thread.__init__(self)
        self.wb = xlrd.open_workbook('config.xlsx')

        """
        self.sh = self.wb.sheet_by_name(u'Actionneur')
        self.liste_thread = liste_thread
        self.compteur = 0
        self.config = list()
        for rownum in range(self.sh.nrows):
            print(self.sh.row_values(rownum))
            self.config.append(self.sh.row_values(rownum))
            print(self.compteur)
            self.compteur += 1
        self.config_actuel = self.config[:]
        self.start_run = True
        """

        self.liste_feuille = self.wb.sheet_names()
        del self.liste_feuille[0]
        print("liste feuille : ",self.liste_feuille)
        self.config = list()
        self.compteur = 0
        compteur = 0
        for feuille in self.liste_feuille:
            self.sh = self.wb.sheet_by_name(feuille)
            for rownum in range(self.sh.nrows):
                if rownum != 0 :
                    self.config.append(self.sh.row_values(rownum))
            pass

            compteur += 1
        print("config = ", self.config)
    pass

    def run(self):
        while self.start_run:
            self.wb = xlrd.open_workbook('config.xlsx')
            self.sh = self.wb.sheet_by_name(u'Actionneur')
            #print()
            #print()

            self.compteur = 0
            for rownum in range(self.sh.nrows):
                self.config_actuel[self.compteur] = self.sh.row_values(rownum)
                #print("config_actuel =",self.config_actuel[self.compteur])
                #print("config_precedente =",self.config[self.compteur])
                if self.config_actuel[self.compteur] != self.config[self.compteur]:
                    if self.config_actuel[self.compteur][1] != self.config[self.compteur][1]:
                        self.liste_thread[self.compteur-1].set_anomalie()
                        self.config[self.compteur][1] = self.config_actuel[self.compteur][1]

                    if self.config_actuel[self.compteur][2] != self.config[self.compteur][2]:
                        self.liste_thread[self.compteur-1].set_default()
                        self.config[self.compteur][2] = self.config_actuel[self.compteur][2]

                    if self.config_actuel[self.compteur][3] != self.config[self.compteur][3]:
                        if self.config_actuel[self.compteur][3] == 1:
                            self.liste_thread[self.compteur-1].simulation_fdc(1)
                        else:
                            self.liste_thread[self.compteur-1].simulation_fdc(0)
                        pass
                        self.config[self.compteur][3] = self.config_actuel[self.compteur][3]

                    if self.config_actuel[self.compteur][4] != self.config[self.compteur][4]:
                        if self.config_actuel[self.compteur][4] == 1:
                            self.liste_thread[self.compteur-1].fdc_ouv_methode(1)
                        else:
                            self.liste_thread[self.compteur-1].fdc_ouv_methode(0)
                        pass
                        self.config[self.compteur][4] = self.config_actuel[self.compteur][4]

                    if self.config_actuel[self.compteur][5] != self.config[self.compteur][5]:
                        if self.config_actuel[self.compteur][5] == 1:
                            self.liste_thread[self.compteur-1].fdc_fer_methode(1)
                        else:
                            self.liste_thread[self.compteur-1].fdc_fer_methode(0)
                        pass
                        self.config[self.compteur][5] = self.config_actuel[self.compteur][5]

                self.compteur += 1
            time.sleep(2)
            pass
        pass


if __name__ == '__main__':
    recuperation_serveur()
    equipement = recuperation_equipement()
    for i in equipement:
        for x in i:
            print(x)
    pass