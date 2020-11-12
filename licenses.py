import json
import uuid
import xlrd
import os
import xlwt
import xlsxwriter
import openpyxl
import os.path
import datetime
from datetime import datetime
import requests

class licenses():
    def __init__(self,id):
        self.id=id

    def licencePrint(self, ermName, ermOrgid, ermOrgName, ermDescription, aliases, fileName):
        Ordarchivo=open("macewan\licenses\macewan_licenses.json", 'a')
        license ={
            "id": self.id,
            #"dateCreated": "2020-05-29T20:40:32Z",
            "links": [],
            "customProperties": {},
            "contacts": [],
            "description": ermDescription,
            #"lastUpdated": "2020-05-29T20:40:32Z",
            "docs": [],
            "name": ermName,
            "status": {"id": "", "value": "active","label": "Active"},
            "supplementaryDocs": [],
            "_links": {
                "linkedResources": {
                    "href": "/licenses/licenseLinks?filter=owner.id%"+self.id,
                }
            },
            "openEnded": False,
            "amendments": [],
            "orgs": [{ "id": "","org": {"id": "","orgsUuid": ermOrgid,"name": ermOrgName},"owner": {"id": self.id},"role": {"id": "","value": "licensor","label": "Licensor"}}],
            "alternateNames": [aliases],
            "type": {"id": "2c9180857445294501756738c3be0028","value": "unspecified","label": "Unspecified"},
            }
        #json_ord = json.dumps(order,indent=2)
        json_ord = json.dumps(license)
        print('Datos en formato JSON', json_ord)
        Ordarchivo.write(json_ord+"\n")


def getOrgId(orgname):
        dic={}
        #pathPattern="/organizations-storage/organizations" #?limit=9999&query=code="
        pathPattern="/organizations/organizations" #?limit=9999&query=code="
        okapi_url="https://okapi-liverpool.folio.ebsco.com"
        okapi_token="eyJhbGciOiJIUzI1NiJ9.eyJzdWIiOiJhZG1pbiIsInVzZXJfaWQiOiI2NGEyZWY0Yy04YjBkLTRlMjYtYmU3Yy1jOWNkNmM4MTYwYmMiLCJpYXQiOjE2MDI3NzM1NDMsInRlbmFudCI6ImZzMDAwMDEwNDUifQ.RoGwG6x9ivM7tJJ1d7vNZmFh5k6rfmMdT-9pLC47t_0"
        okapi_tenant="fs00001045"
        okapi_headers = {"x-okapi-token": okapi_token,"x-okapi-tenant": okapi_tenant,"content-type": "application/json"}
        length="1"
        start="1"
        element="organizations"
        query=f"query=name=="
        #/organizations-storage/organizations?query=code==UMPROQ
        paging_q = f"?{query}"+orgname
        path = pathPattern+paging_q
        #data=json.dumps(payload)
        url = okapi_url + path
        req = requests.get(url, headers=okapi_headers)
        idorg=[]
        if req.status_code != 201:
            json_str = json.loads(req.text)
            total_recs = int(json_str["totalRecords"])
            if (total_recs!=0):
                rec=json_str[element]
                #print(rec)
                l=rec[0]
                if 'id' in l:
                    idorg.append(l['id'])
                    idorg.append(l['name'])

        return idorg
###################################################################################################
#
#
###################################################################################################
def readSpreadsheet(spreadsheet,org):
    wb = xlrd.open_workbook(spreadsheet)
    fileN=org
    worksheet = wb.sheet_by_name("Sheet1")
    print("no filas:", worksheet.nrows)
    print("no filas:", worksheet.ncols)
    f = open("licenses_error.txt", "a")
    #read orders
    count=0
    oldlicense=""
    for p in range(worksheet.nrows):
        if (p!=0):
            if oldlicense!=worksheet.cell_value(p,0):
                oldlicense=worksheet.cell_value(p,0)
                ermLicenceYear="Year: "+str(worksheet.cell_value(p,0))
                ermLicenceYear=ermLicenceYear.replace(".0","")
                ermDataPackageName=str(worksheet.cell_value(p,1).strip())
                orgname=worksheet.cell_value(p,2).strip()
                vendor=getOrgId(orgname)
                if len(vendor)==0:
                    vendor.append("0b3ffc1e-1fa5-4d40-8c93-e592aa94ab57")
                    vendor.append("EBSCO")
                erm_aliases=worksheet.cell_value(p,4).strip()

                uuidpol=str(uuid.uuid4())
            ##################ermName, ermOrgid,ermOrgName
                org=licenses(uuidlic)
                org.licencePrint(ermDataPackageName,vendor[0], vendor[1], ermLicenceYear, erm_aliases, fileN)

if __name__ == "__main__":

    """This is the Starting point for the script"""    
    #readSpreadsheet("macewan\licenses\MacEwanLicense.xlsx", "Macewan_licenses")

    readSpreadsheet("Liverpool\licenses\Liverpool_License_SourceDataFinal.xlsx", "Liverpool")
    