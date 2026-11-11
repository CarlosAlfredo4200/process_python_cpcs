########### Python 3.2 #############
import urllib.request, json
id_colegio = 10
try:
    url = "https://api.q10.com/v1/estudiantes/colegios/{212027}"

    hdr ={
            "Cache-Control": "no-cache",
            "Accept": "application/json",
            'Api-Key': '23117b580a1048eab7fa33021ad255b4',
    }

    req = urllib.request.Request(url, headers=hdr)

    req.get_method = lambda: 'GET'
    response = urllib.request.urlopen(req)
    print(response.getcode())
    print(response.read())
except Exception as e:
    print(e)
####################################