import xlrd
import config_excel
from main import Connexion
from opcua import ua
import time
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
        print(liste_mouvement)
        print(len(liste_mouvement))
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

if __name__ == '__main__':
    liste_thread = list()  # liste des threads des equipements
    serveur, noeud = config_excel.recuperation_serveur()
    client_connexion = Connexion(serveur)
    equipements = config_excel.recuperation_equipement()
    nombre_type_equipement = len(equipements)
    compteur = 0
    sequencement = Sequencement(client_connexion.client,noeud)
    pass