import logging
import googlemaps
import ipdb
import pandas as pd
from pprint import pformat

# Inserte aqui su API Key
KEY = ''

class ApiGoogleMaps:
    def __init__(self, address):
        self.client = googlemaps.Client(key=KEY)
        self.address = address
        self.gm_res = self.client.geocode(self.address)
    
    def get_lat_lng(self):
        if self.gm_res:
            geometry = self.gm_res[0].get('geometry', False)
            if geometry and geometry.get('location', False):
                return geometry['location']['lat'], geometry['location']['lng']
        return '', ''

    def get_formatted_address(self):
        if self.gm_res:
            return self.gm_res[0].get('formatted_address', 'no_address')

    def get_partial_match(self):
        if self.gm_res:
            return self.gm_res[0].get('partial_match', False)


def run():
    # Insertar aqui tu ruta de archivo excel
    fpath = '/home/jrojas/Descargas/arch202105271343-YJKZ30WE46.xls'
    wb = pd.read_excel(fpath, sheet_name=0)
    file_addresses = []
    gm_addresses = []
    gm_lat = []
    gm_lng = []
    gm_partial_match = []
    print("Procesando Archivo")
    for i, row in wb.iterrows():
        print(f'{i}%')
        # Insertar aqui tu formato de direccón
        address = f"{row.at['DIRECCION ENTREGA']} {row.at['DISTRITO ENTREGA']} {row.at['PROVINCIA ENTREGA']} {row.at['REGION ENTREGA']}"

        file_addresses.append(address)
        
        gm = ApiGoogleMaps(address)
        gm_addresses.append(gm.get_formatted_address())

        lat, lng = gm.get_lat_lng()
        gm_lat.append(lat)
        gm_lng.append(lng)

        partial_match = gm.get_partial_match() 
        if partial_match is None:
            partial_match = True
        gm_partial_match.append(partial_match)
        

    df = pd.DataFrame({
        'Direcciones': file_addresses,
        'Direcciones (GM)': gm_addresses,
        'Latitud (GM)': gm_lat,
        'Longitud (GM)': gm_lng,
        'Coincidencias Parciales (GM)': gm_partial_match
    })
    tcp = ['' for v in gm_partial_match]
    cp = len(list(filter(lambda s: s is True, gm_partial_match)))
    ce = len(list(filter(lambda s: s is False, gm_partial_match)))
    tcp[0] = f'Coincidencias Parciales {cp}'
    tcp[1] = f'Coincidencias Exactas {ce}'
    tcp[2] = f'Eficiencia al {ce/len(gm_partial_match)}%'
    df['Totales'] = tcp
    print('Completo!')
    df.to_excel('googlem_maps_address.xls')

run()