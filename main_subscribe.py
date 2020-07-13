import opcua
from opcua import ua
from opcua.common.node import Node
import time

class SubHandler(object):

    """
    Subscription Handler. To receive events from server for a subscription
    data_change and event methods are called directly from receiving thread.
    Do not do expensive, slow or network operation there. Create another
    thread if you need to do such a thing
    """
    def __init__(self, actionneur):
        self.actionneur = actionneur
        pass
    def datachange_notification(self, node, val, data):
        print("Python: New data change event", node, val)
        if val:
            self.actionneur.set_value(True)
            pass
        else:
            self.actionneur.set_value(False)
            pass


    def event_notification(self, event):
        print("Python: New event", event)

class Actionneur():
    def __init__(self, client):
        self.client = client
        self.fdc_ouv = client.get_node("ns=2;s=Siemens_TCP_IP_simulateur.Automate_HE.VB-536-114.Fdc_Ouv")
        self.commande = client.get_node("ns=2;s=Siemens_TCP_IP_simulateur.Automate_HE.VB-536-114.Commande")
        pass

    def set_value(self, value):
        print("set_value ", value)
        self.fdc_ouv.set_attribute(ua.AttributeIds.Value, ua.DataValue(value))


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
    client_connexion = Connexion("opc.tcp://192.168.10.123:49320")
    actionneur = Actionneur(client_connexion.client)

    handler = SubHandler(actionneur)
    sub = client_connexion.client.create_subscription(2000, handler)
    sub.subscribe_data_change(actionneur.commande)


    pass
