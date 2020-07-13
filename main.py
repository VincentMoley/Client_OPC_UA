import opcua
from opcua import ua
import time
from threading import Thread
import xlrd
from random import *


class Actionneur(Thread):

    def __init__(self, client, noeud, equipement, numero_equipement):
        Thread.__init__(self, name=equipement)
        print("Actionneur cree : ", equipement, " numero_equipement : ", numero_equipement)
        self.numero_equipement = numero_equipement
        self.equipement = equipement
        self.client = client
        self.fdc_ouv = client.get_node(noeud + equipement + ".Fdc_Ouv")
        self.fdc_fer = client.get_node(noeud + equipement + ".Fdc_Fer")
        self.anomalie = client.get_node(noeud + equipement + ".Anomalie")
        self.default = client.get_node(noeud + equipement + ".Defaut")
        self.commande = client.get_node(noeud + equipement + ".Commande")
        self.etat = client.get_node(noeud + equipement + ".Etat")
        self.commande_value = self.commande.get_value()

        self.simulation_fin_de_course = True
        # Iniitalisation des fin de courses
        if self.commande.get_value():
            self.fdc_ouv.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
            self.fdc_fer.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
        else:
            self.fdc_ouv.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
            self.fdc_fer.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
        pass
        self.wb = xlrd.open_workbook('config.xlsx')
        self.sh = self.wb.sheet_by_name(u'Actionneur')

        self.compteur = 0
        self.config = list()
        self.config = (self.sh.row_values(self.numero_equipement))

        self.config_actuel = self.config[:]
        self.start_run = True

    def run(self):
        while self.start_run:
            if self.simulation_fin_de_course:
                if self.commande.get_value() != self.commande_value:
                    self.commande_value = self.commande.get_value()
                    if self.commande.get_value():
                        self.fdc_ouv.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))
                        self.fdc_fer.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                    else:
                        self.fdc_ouv.set_attribute(ua.AttributeIds.Value, ua.DataValue(False))
                        self.fdc_fer.set_attribute(ua.AttributeIds.Value, ua.DataValue(True))

                pass

            # surveillance de donnees excel
            self.surveillance_donnees_excel()
            time.sleep(0.1)
        pass

    def stop(self):
        #print("arret thread :", self.name)
        self.start_run = False

        pass

    def set_value(self, value):
        #print("set_value ", value)
        self.fdc_ouv.set_attribute(ua.AttributeIds.Value, ua.DataValue(value))
        pass

    def set_anomalie(self, value):
        self.etat.set_attribute(ua.AttributeIds.Value, ua.DataValue(ua.Variant(value, ua.VariantType.UInt16)))
        pass

    def set_default(self, value):
        self.etat.set_attribute(ua.AttributeIds.Value, ua.DataValue(ua.Variant(value, ua.VariantType.UInt16)))
        pass

    def simulation_fdc(self, value):
        print("equipement : ", self.equipement, " ,fdc simulation : ", value)
        self.simulation_fin_de_course = value
        pass

    def fdc_ouv_methode(self, value):
        #print("equipement : ", self.equipement, " ,fdc_ouv value : ", value)
        self.fdc_ouv.set_attribute(ua.AttributeIds.Value, ua.DataValue(bool(value)))
        pass

    def fdc_fer_methode(self, value):
        #print("equipement : ", self.equipement, " ,fdc_fer value : ", value)
        self.fdc_fer.set_attribute(ua.AttributeIds.Value, ua.DataValue(bool(value)))
        pass

    def surveillance_donnees_excel(self):

        self.wb = xlrd.open_workbook('config.xlsx')
        self.sh = self.wb.sheet_by_name(u'Actionneur')

        self.config_actuel = self.sh.row_values(self.numero_equipement)
        if self.config_actuel != self.config:
            if self.config_actuel[1] != self.config[1]:
                if self.config_actuel[1]:
                    self.set_anomalie(3)
                elif self.config_actuel[1] == False and self.config_actuel[2] == False:
                    self.set_anomalie(0)
                elif self.config_actuel[1] == False and self.config_actuel[2] == True:
                    self.set_anomalie(4)
                pass
                if self.config_actuel[1] and self.config_actuel[2]:
                    self.set_anomalie(5)
            if self.config_actuel[2] != self.config[2]:
                if self.config_actuel[2]:
                    self.set_default(4)
                elif self.config_actuel[1] == False and self.config_actuel[2] == False:
                    self.set_anomalie(0)
                elif self.config_actuel[1] == False and self.config_actuel[2] == True:
                    self.set_default(3)
                pass
                if self.config_actuel[1] and self.config_actuel[2]:
                    self.set_default(5)
            if self.config_actuel[3] != self.config[3]:
                if self.config_actuel[3] == 0:
                    self.simulation_fdc(0)
                else:
                    self.simulation_fdc(1)
                pass
            if self.config_actuel[4] != self.config[4]:
                if self.config_actuel[4] == 0:
                    self.fdc_ouv_methode(0)
                else:
                    self.fdc_ouv_methode(1)
                pass
            if self.config_actuel[5] != self.config[5]:
                if self.config_actuel[5] == 0:
                    self.fdc_fer_methode(0)
                else:
                    self.fdc_fer_methode(1)
                pass
            self.config[:] = self.config_actuel
            pass
        pass


class Capteurs(Thread):
    def __init__(self, client, noeud, equipement, numero_equipement):
        Thread.__init__(self, name=equipement)
        print("Capteur cree : ", equipement, " numero equipement : ", numero_equipement)
        self.numero_equipement = numero_equipement
        self.equipement = equipement
        self.valeur = client.get_node(noeud + equipement + ".Mesure_A")

        self.client = client

        self.wb = xlrd.open_workbook('config.xlsx')
        self.sh = self.wb.sheet_by_name(u'Capteur')


        self.compteur = 0
        self.config = list()
        self.config = self.sh.row_values(self.numero_equipement)
        self.config_actuel = self.config[:]
        self.start_run = True
        self.control_simulation_rampe = False

    def run(self):
        while self.start_run == True:
            self.surveillance_donnees_excel()
            time.sleep(1)
            pass
        pass

    def stop(self):
        self.start_run = False
        pass

    def surveillance_donnees_excel(self):

        self.wb = xlrd.open_workbook('config.xlsx')
        self.sh = self.wb.sheet_by_name(u'Capteur')

        self.config_actuel = self.sh.row_values(self.numero_equipement)
        if self.config_actuel != self.config:
            if self.config_actuel[1] != self.config[1]:
                self.set_value(int(self.config_actuel[1]))

            elif self.config_actuel[2] != self.config[2]:
                if self.config_actuel[2] == 1:
                    self.control_simulation_rampe = True
                    self.simulation_rampe(self.config_actuel[4], self.config_actuel[5], self.config_actuel[3],
                                          self.config_actuel[6])

                else:
                    self.control_simulation_rampe = False

            self.config[:] = self.config_actuel
            pass
        pass

    def set_value(self, value):
        try:
            self.valeur.set_attribute(ua.AttributeIds.Value, ua.DataValue(ua.Variant(value, ua.VariantType.Float)))
        except:
            print(value)

        pass

    def simulation_rampe(self, min, max, incremente, temps_incrementation):
        while self.control_simulation_rampe:
            self.mesure = self.valeur.get_value()
            if self.mesure >= max:

                self.set_value(min)
            elif self.mesure < min:
                self.mesure(min)
            else:
                if self.mesure + incremente <= max:
                    #self.set_value(self.mesure + incremente)
                    self.set_value(round(uniform(0, 100)))
                else:
                    self.set_value(max)
            time.sleep(temps_incrementation)
        pass


class Sequencement():
    def __init__(self, client, noeud):
        self.wb = xlrd.open_workbook('config.xlsx')
        self.sh = self.wb.sheet_by_name(u'Sequencement')
        self.client = client
        self.noeud = noeud
        #self.liste_sequencement = self.sh.row_values(self.numero_equipement)
        liste_mouvement = list()
        compteur = 0
        for colnum in range(self.sh.ncols):
            liste_mouvement.append(self.sh.col_values(colnum))
            compteur += 1
        for mouvement in liste_mouvement:
            del mouvement[0]
        self.liste_mouvement = liste_mouvement
        self.sequence()
        pass

    def sequence(self):
        for y in self.liste_mouvement:
            compteur = 0
            for x in y:
                self.etat = self.client.get_node(self.noeud + self.liste_mouvement[0][compteur] + ".Etat")
                if x == 0:
                    """
                    self.ouverture.set_attribute(ua.AttributeIds.Value, ua.DataValue(bool(False)))
                    self.fermeture.set_attribute(ua.AttributeIds.Value, ua.DataValue(bool(True)))
                    """
                    try:
                        self.etat.set_attribute(ua.AttributeIds.Value,
                                                ua.DataValue(ua.Variant(1, ua.VariantType.UInt16)))
                        # self.fdc_ouv.set_attribute(ua.AttributeIds.Value, ua.DataValue(bool(False)))
                        # self.fdc_fer.set_attribute(ua.AttributeIds.Value, ua.DataValue(bool(True)))
                    except Exception as e:
                        print(e)
                        print("erreur set_value : ", self.liste_mouvement[0][compteur])
                        pass
                else:
                    """
                    self.fermeture.set_attribute(ua.AttributeIds.Value, ua.DataValue(bool(False)))
                    self.ouverture.set_attribute(ua.AttributeIds.Value, ua.DataValue(bool(True)))
                    """

                    try:
                        self.etat.set_attribute(ua.AttributeIds.Value,
                                                ua.DataValue(ua.Variant(2, ua.VariantType.UInt16)))
                        # self.fdc_fer.set_attribute(ua.AttributeIds.Value, ua.DataValue(bool(False)))
                        # self.fdc_ouv.set_attribute(ua.AttributeIds.Value, ua.DataValue(bool(True)))
                    except Exception as e:
                        print(e)
                        print("erreur set_value : ", self.liste_mouvement[0][compteur])
                        pass

                pass
                compteur += 1
                pass
            time.sleep(0.5)

        pass

    pass

class Connexion:
    def __init__(self, url):
        self.client = None
        self.url = url
        self.connexion()
        pass

    def connexion(self):
        self.client = opcua.Client(self.url)
        self.client.connect()
        self.client.load_type_definitions()
        pass

    def deconnexion(self):
        self.client.disconnect()
        pass

    pass


if __name__ == '__main__':
    import config_excel

    liste_thread = list()  # liste des threads des equipements
    serveur, noeud = config_excel.recuperation_serveur()
    client_connexion = Connexion(serveur)
    equipements = config_excel.recuperation_equipement()
    nombre_type_equipement = len(equipements)
    compteur = 0

    while compteur < nombre_type_equipement:
        compteur_equipement = 1

        if compteur == 0:
            for i in equipements[compteur]:
                actionneur = Actionneur(client_connexion.client, noeud, i, compteur_equipement)
                compteur_equipement += 1
                actionneur.start()
                liste_thread.append(actionneur)
                pass
            pass
        elif compteur == 1:
            for i in equipements[compteur]:
                capteur = Capteurs(client_connexion.client, noeud, i, compteur_equipement)
                compteur_equipement += 1
                capteur.start()
                liste_thread.append(capteur)
            pass

        if compteur == 2:
            for i in equipements[compteur]:
                sequencement = Sequencement(client_connexion.client, noeud)

            pass

        compteur += 1
        pass
