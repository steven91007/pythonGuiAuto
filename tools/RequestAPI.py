import requests
import json


class API:

    @staticmethod
    def GetApplications():

        response = requests.get('https://msplatform.deltaww.com/rpa/productmaking/applications')
        data = response.json()
        return data

    @staticmethod
    def PostApplicationData(id, form_data):

        response = requests.post('https://msplatform.deltaww.com/rpa/productmaking/applications/{}'.format(str(id)), files=form_data)
        print(response.text)

    @staticmethod
    def GetCheckedApplications():
        response = requests.get('https://msplatform.deltaww.com/rpa/productmaking/applications/checked')
        data = response.json()
        return data

    @staticmethod
    def PostApplicationPrinted(id):
        response = requests.post('https://msplatform.deltaww.com/rpa/productmaking/applications/{}/pass'.format(str(id)))
        print(response.text)

if __name__ == '__main__':
    # GetApplications()

    # filename = r'C:\Users\amo.cy.hsu\PycharmProjects\pythonGuiAuto\2870686002.pdf'
    #
    # multipart_form_data = {
    # 'file': (filename, open(filename,'rb')),
    # 'revision': (None,'00'),
    # 'state': (None, 1),
    # 'reason': (None, '')
    # }
    # PostApplicationData(3, multipart_form_data)
    API.PostApplicationPrinted(3)
    API.GetCheckedApplications()